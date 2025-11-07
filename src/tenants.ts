import { createHash, timingSafeEqual } from 'node:crypto';

import { CONFIG, type ApiKeyConfig, type TenantConfig, type SiteType } from './config.js';
import { logger } from './logger.js';
import { resolveSecretSource } from './utils/secrets.js';

export interface TenantSiteSettings {
  siteDisplayName: string;
  siteType: SiteType;
  host?: string;
  sitePath?: string;
}

export interface TenantRecord {
  id: string;
  name: string;
  active: boolean;
  sharepoint: TenantSiteSettings;
  metadata?: Record<string, unknown>;
}

export interface ApiKeyView {
  id: string;
  name: string;
  tenantId: string;
  roles: readonly string[];
  active: boolean;
  createdAt?: string;
  lastRotatedAt?: string;
  expiresAt?: string;
  metadata?: Record<string, unknown>;
}

export interface ApiKeySecretView {
  id?: string;
  active: boolean;
  notBefore?: string;
  expiresAt?: string;
}

interface ApiKeySecretInternal {
  hash: Buffer;
  view: ApiKeySecretView;
  notBefore?: number;
  expiresAt?: number;
}

interface ApiKeyInternal {
  view: ApiKeyView;
  active: boolean;
  expiresAt?: number;
  secrets: ApiKeySecretInternal[];
}

export type ApiKeyAuthFailureCode =
  | 'invalid'
  | 'api_key_inactive'
  | 'api_key_expired'
  | 'secret_inactive'
  | 'secret_not_yet_valid'
  | 'secret_expired'
  | 'tenant_inactive'
  | 'tenant_missing';

export type ApiKeyAuthResult =
  | {
      ok: true;
      tenant: TenantRecord;
      apiKey: ApiKeyView;
      secret: ApiKeySecretView;
      presentedHash: string;
    }
  | {
      ok: false;
      code: ApiKeyAuthFailureCode;
      message: string;
      tenantId?: string;
      apiKeyId?: string;
      secretId?: string;
    };

const tenantsById = new Map<string, TenantRecord>();
const apiKeyRegistry: ApiKeyInternal[] = [];

function normaliseTenant(config: TenantConfig): TenantRecord {
  const siteDisplayName =
    config.sharepoint.siteDisplayName || CONFIG.sharepoint.siteDisplayName || config.name;
  const siteType: SiteType = config.sharepoint.siteType ?? CONFIG.sharepoint.siteType;
  const host = config.sharepoint.host ?? CONFIG.sharepoint.host;
  const sitePath = config.sharepoint.sitePath;

  return {
    id: config.id,
    name: config.name,
    active: config.active !== false,
    sharepoint: {
      siteDisplayName,
      siteType,
      host,
      sitePath,
    },
    metadata: config.metadata,
  };
}

function toTimestamp(value?: string): number | undefined {
  if (!value) {
    return undefined;
  }
  const ts = Date.parse(value);
  return Number.isNaN(ts) ? undefined : ts;
}

function normaliseApiKey(config: ApiKeyConfig): ApiKeyInternal {
  const active = config.active !== false;
  const expiresAt = toTimestamp(config.expiresAt);

  const view: ApiKeyView = {
    id: config.id,
    name: config.name,
    tenantId: config.tenantId,
    roles: [...config.roles],
    active,
    createdAt: config.createdAt,
    lastRotatedAt: config.lastRotatedAt,
    expiresAt: config.expiresAt,
    metadata: config.metadata,
  };

  const secrets: ApiKeySecretInternal[] = config.secrets.map((secret, index) => {
    const label = `API key ${config.id} secret ${secret.id ?? index}`;
    let material = resolveSecretSource(label, {
      value: secret.value,
      env: secret.env,
      file: secret.file,
    });

    const buffer = Buffer.from(material, 'utf8');
    material = '';
    const hash = createHash('sha256').update(buffer).digest();
    buffer.fill(0);

    const notBefore = toTimestamp(secret.notBefore);
    const secretExpiresAt = toTimestamp(secret.expiresAt);
    const activeSecret = secret.active !== false;

    return {
      hash,
      notBefore,
      expiresAt: secretExpiresAt,
      view: {
        id: secret.id,
        active: activeSecret,
        notBefore: secret.notBefore,
        expiresAt: secret.expiresAt,
      },
    };
  });

  return {
    view,
    active,
    expiresAt,
    secrets,
  };
}

(function initialiseTenants() {
  CONFIG.tenants.forEach((tenantConfig) => {
    const tenant = normaliseTenant(tenantConfig);
    tenantsById.set(tenant.id, tenant);
  });

  if (tenantsById.size === 0) {
    logger.warn('No tenants configured; default SharePoint settings will be used if available');
  }
})();

(function initialiseApiKeys() {
  CONFIG.apiKeys.forEach((apiKeyConfig) => {
    const tenant = tenantsById.get(apiKeyConfig.tenantId);
    if (!tenant) {
      logger.error('API key configured for unknown tenant', {
        apiKeyId: apiKeyConfig.id,
        tenantId: apiKeyConfig.tenantId,
      });
      return;
    }

    try {
      const normalised = normaliseApiKey(apiKeyConfig);
      apiKeyRegistry.push(normalised);
    } catch (error) {
      logger.error('Failed to initialise API key secret material', {
        apiKeyId: apiKeyConfig.id,
        tenantId: apiKeyConfig.tenantId,
        error: (error as Error).message,
      });
    }
  });

  if (apiKeyRegistry.length === 0) {
    logger.warn('No API keys configured; authenticated routes will reject requests');
  }
})();

export function listTenants(): TenantRecord[] {
  return Array.from(tenantsById.values());
}

export function getTenantById(id: string): TenantRecord | undefined {
  return tenantsById.get(id);
}

export function getDefaultTenant(): TenantRecord | undefined {
  return Array.from(tenantsById.values()).find((tenant) => tenant.active);
}

export function authenticateApiKey(rawKey: string | undefined): ApiKeyAuthResult {
  if (!rawKey || rawKey.trim().length === 0) {
    return { ok: false, code: 'invalid', message: 'API key missing' };
  }

  const candidate = rawKey.trim();
  const hash = createHash('sha256').update(candidate, 'utf8').digest();
  const hashHex = hash.toString('hex');
  const now = Date.now();

  for (const entry of apiKeyRegistry) {
    for (const secret of entry.secrets) {
      if (secret.hash.length === hash.length && timingSafeEqual(secret.hash, hash)) {
        const tenant = tenantsById.get(entry.view.tenantId);
        if (!tenant) {
          return {
            ok: false,
            code: 'tenant_missing',
            message: 'Tenant is not configured for this API key',
            apiKeyId: entry.view.id,
            tenantId: entry.view.tenantId,
            secretId: secret.view.id,
          };
        }

        if (!tenant.active) {
          return {
            ok: false,
            code: 'tenant_inactive',
            message: 'Tenant is inactive',
            apiKeyId: entry.view.id,
            tenantId: tenant.id,
            secretId: secret.view.id,
          };
        }

        if (!entry.active) {
          return {
            ok: false,
            code: 'api_key_inactive',
            message: 'API key is inactive',
            apiKeyId: entry.view.id,
            tenantId: tenant.id,
            secretId: secret.view.id,
          };
        }

        if (entry.expiresAt && now >= entry.expiresAt) {
          return {
            ok: false,
            code: 'api_key_expired',
            message: 'API key has expired',
            apiKeyId: entry.view.id,
            tenantId: tenant.id,
            secretId: secret.view.id,
          };
        }

        if (!secret.view.active) {
          return {
            ok: false,
            code: 'secret_inactive',
            message: 'Secret version is inactive',
            apiKeyId: entry.view.id,
            tenantId: tenant.id,
            secretId: secret.view.id,
          };
        }

        if (secret.notBefore && now < secret.notBefore) {
          return {
            ok: false,
            code: 'secret_not_yet_valid',
            message: 'Secret version is not yet valid',
            apiKeyId: entry.view.id,
            tenantId: tenant.id,
            secretId: secret.view.id,
          };
        }

        if (secret.expiresAt && now >= secret.expiresAt) {
          return {
            ok: false,
            code: 'secret_expired',
            message: 'Secret version has expired',
            apiKeyId: entry.view.id,
            tenantId: tenant.id,
            secretId: secret.view.id,
          };
        }

        return {
          ok: true,
          tenant,
          apiKey: entry.view,
          secret: secret.view,
          presentedHash: hashHex,
        };
      }
    }
  }

  return { ok: false, code: 'invalid', message: 'API key not recognised' };
}
