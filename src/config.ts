import 'dotenv/config';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { z } from 'zod';

import { logger } from './logger.js';
import { readSecretFromEnv } from './utils/secrets.js';

export type SiteType = 'team' | 'communication';

const ALLOWED_ROLES = ['admin', 'api-access', 'read-only', 'share-manager'] as const;

const ApiKeySecretSchema = z
  .object({
    id: z.string().min(1).optional(),
    value: z.string().min(1).optional(),
    env: z.string().min(1).optional(),
    file: z.string().min(1).optional(),
    active: z.boolean().default(true),
    notBefore: z.string().datetime().optional(),
    expiresAt: z.string().datetime().optional(),
  })
  .refine((secret) => secret.value || secret.env || secret.file, {
    message: 'Secret definition must provide value, env, or file',
  });

export type ApiKeySecretConfig = z.infer<typeof ApiKeySecretSchema>;

const ApiKeySchema = z.object({
  id: z.string().min(1),
  name: z.string().min(1),
  tenantId: z.string().min(1),
  roles: z
    .array(z.enum(ALLOWED_ROLES))
    .min(1, 'API key must include at least one role')
    .default(['api-access']),
  active: z.boolean().default(true),
  createdAt: z.string().datetime().optional(),
  lastRotatedAt: z.string().datetime().optional(),
  expiresAt: z.string().datetime().optional(),
  secrets: z.array(ApiKeySecretSchema).min(1, 'API key must define at least one secret version'),
  metadata: z.record(z.unknown()).optional(),
});

export type ApiKeyConfig = z.infer<typeof ApiKeySchema>;

const TenantSharePointSchema = z
  .object({
    siteDisplayName: z.string().min(1).optional(),
    siteType: z.enum(['team', 'communication']).optional(),
    host: z.string().min(1).optional(),
    sitePath: z.string().min(1).optional(),
  })
  .default({});

const TenantSchema = z.object({
  id: z.string().min(1),
  name: z.string().min(1),
  active: z.boolean().default(true),
  sharepoint: TenantSharePointSchema,
  metadata: z.record(z.unknown()).optional(),
});

export type TenantConfig = z.infer<typeof TenantSchema>;

const DEFAULT_CONFIG = {
  environment: 'development',
  server: {
    port: 8080,
  },
  azure: {
    tenantId: '',
    clientId: '',
    clientSecret: '',
    scope: 'https://graph.microsoft.com/.default',
  },
  sharepoint: {
    siteDisplayName: 'Elion Studio',
    siteType: 'team' as SiteType,
    host: undefined as string | undefined,
  },
  openai: {
    apiKey: '',
    model: 'gpt-4.1',
  },
  redis: {
    url: undefined as string | undefined,
    tls: false,
    keyPrefix: 'elion:sp-agent',
  },
  auth: {
    tokenCacheTtlMinutes: 55,
    tokenFallbackMs: 5 * 60 * 1000,
  },
  tenants: [] as TenantConfig[],
  apiKeys: [] as ApiKeyConfig[],
} as const;

const ConfigSchema = z
  .object({
    environment: z
      .enum(['development', 'test', 'staging', 'production'])
      .default(DEFAULT_CONFIG.environment),
    server: z.object({
      port: z
        .number({ invalid_type_error: 'PORT must be a number' })
        .int('PORT must be an integer')
        .min(1, 'PORT must be between 1 and 65535')
        .max(65535, 'PORT must be between 1 and 65535')
        .default(DEFAULT_CONFIG.server.port),
    }),
    azure: z.object({
      tenantId: z.string().min(1, 'AZURE_TENANT_ID is required'),
      clientId: z.string().min(1, 'AZURE_CLIENT_ID is required'),
      clientSecret: z.string().min(1, 'AZURE_CLIENT_SECRET is required'),
      scope: z.string().min(1).default(DEFAULT_CONFIG.azure.scope),
    }),
    sharepoint: z
      .object({
        siteDisplayName: z
          .string({ invalid_type_error: 'SITE_DISPLAY_NAME must be a string' })
          .min(1, 'SITE_DISPLAY_NAME cannot be empty')
          .default(DEFAULT_CONFIG.sharepoint.siteDisplayName),
        siteType: z
          .enum(['team', 'communication'])
          .default(DEFAULT_CONFIG.sharepoint.siteType),
        host: z
          .string()
          .transform((value) => value.trim())
          .transform((value) => value.replace(/\/+$/, ''))
          .optional(),
      })
      .default(DEFAULT_CONFIG.sharepoint),
    openai: z.object({
      apiKey: z.string().min(1, 'OPENAI_API_KEY is required'),
      model: z.string().min(1).default(DEFAULT_CONFIG.openai.model),
    }),
    redis: z
      .object({
        url: z
          .string()
          .trim()
          .transform((value) => (value.length ? value : undefined))
          .refine((value) => !value || value.startsWith('redis://') || value.startsWith('rediss://'), {
            message: 'REDIS_URL must start with redis:// or rediss://',
          })
          .optional(),
        tls: z.boolean().default(DEFAULT_CONFIG.redis.tls),
        keyPrefix: z
          .string({ invalid_type_error: 'REDIS_KEY_PREFIX must be a string' })
          .min(1, 'REDIS_KEY_PREFIX cannot be empty')
          .default(DEFAULT_CONFIG.redis.keyPrefix),
      })
      .default(DEFAULT_CONFIG.redis),
    auth: z
      .object({
        tokenCacheTtlMinutes: z
          .number({ invalid_type_error: 'AUTH_TOKEN_CACHE_TTL_MINUTES must be a number' })
          .int('AUTH_TOKEN_CACHE_TTL_MINUTES must be an integer')
          .min(1, 'Token cache TTL must be at least 1 minute')
          .max(55, 'Token cache TTL cannot exceed 55 minutes')
          .default(DEFAULT_CONFIG.auth.tokenCacheTtlMinutes),
        tokenFallbackMs: z
          .number({ invalid_type_error: 'AUTH_TOKEN_FALLBACK_MS must be a number' })
          .int('AUTH_TOKEN_FALLBACK_MS must be an integer')
          .min(0, 'Token fallback window must be non-negative')
          .default(DEFAULT_CONFIG.auth.tokenFallbackMs),
      })
      .default(DEFAULT_CONFIG.auth),
    tenants: z.array(TenantSchema).default(DEFAULT_CONFIG.tenants),
    apiKeys: z.array(ApiKeySchema).default(DEFAULT_CONFIG.apiKeys),
  })
  .superRefine((config, ctx) => {
    if (config.sharepoint.siteType === 'communication' && !config.sharepoint.host) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'SHAREPOINT_HOST is required when SITE_TYPE is "communication"',
        path: ['sharepoint', 'host'],
      });
    }

    const tenantIds = new Set<string>();

    config.tenants.forEach((tenant, index) => {
      if (tenantIds.has(tenant.id)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `Duplicate tenant id '${tenant.id}'`,
          path: ['tenants', index, 'id'],
        });
      }
      tenantIds.add(tenant.id);

      const siteType = tenant.sharepoint.siteType ?? config.sharepoint.siteType;
      if (siteType === 'communication') {
        const host = tenant.sharepoint.host ?? config.sharepoint.host;
        if (!host) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            message: 'Communication site tenants must configure sharepoint.host',
            path: ['tenants', index, 'sharepoint', 'host'],
          });
        }
      }
    });

    const apiKeyIds = new Set<string>();

    config.apiKeys.forEach((apiKey, index) => {
      if (apiKeyIds.has(apiKey.id)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `Duplicate API key id '${apiKey.id}'`,
          path: ['apiKeys', index, 'id'],
        });
      }
      apiKeyIds.add(apiKey.id);

      if (!tenantIds.has(apiKey.tenantId)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `API key references unknown tenant '${apiKey.tenantId}'`,
          path: ['apiKeys', index, 'tenantId'],
        });
      }

      if (!apiKey.secrets.some((secret) => secret.active !== false)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: 'API key must have at least one active secret version',
          path: ['apiKeys', index, 'secrets'],
        });
      }
    });
  });

type ConfigInput = z.input<typeof ConfigSchema>;
type AppConfig = z.output<typeof ConfigSchema>;

function parseJsonValue(raw: string, label: string): unknown {
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error(`${label} must contain valid JSON: ${(error as Error).message}`);
  }
}

function loadJsonValue(envVar: string, pathVar: string, label: string): unknown {
  const inline = process.env[envVar];
  if (inline && inline.trim().length > 0) {
    return parseJsonValue(inline, label);
  }

  const filePath = process.env[pathVar];
  if (filePath && filePath.trim().length > 0) {
    try {
      const absolute = resolve(filePath);
      const raw = readFileSync(absolute, 'utf8');
      return parseJsonValue(raw, `${label} file`);
    } catch (error) {
      logger.error('Failed to read JSON configuration file', {
        label,
        path: filePath,
        error: (error as Error).message,
      });
      throw error;
    }
  }

  return undefined;
}

function parseBoolean(value: string | undefined): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (value === 'true' || value === '1') {
    return true;
  }
  if (value === 'false' || value === '0') {
    return false;
  }
  return undefined;
}

function loadRuntimeOverrides(): Record<string, unknown> {
  const overrides: Record<string, unknown> = {};
  const path = process.env.RUNTIME_SECRETS_PATH;

  if (!path) {
    return overrides;
  }

  try {
    const absolute = resolve(path);
    const raw = readFileSync(absolute, 'utf8');
    const parsed = JSON.parse(raw);
    if (typeof parsed === 'object' && parsed !== null) {
      return parsed as Record<string, unknown>;
    }
    throw new Error('Runtime secrets JSON must be an object');
  } catch (error) {
    logger.error('Failed to load runtime secrets overrides', {
      path,
      error: (error as Error).message,
    });
    throw error;
  }
}

function mergeConfig<T extends Record<string, any>>(...sources: Array<T | undefined>): T {
  const target: Record<string, any> = {};

  for (const source of sources) {
    if (!source) continue;
    for (const [key, value] of Object.entries(source)) {
      if (value === undefined) {
        continue;
      }

      if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
        target[key] = mergeConfig(target[key] ?? {}, value);
      } else {
        target[key] = value;
      }
    }
  }

  return target as T;
}

const envOverrides: ConfigInput = {
  environment: (process.env.NODE_ENV as AppConfig['environment']) ?? undefined,
  server: {
    port: process.env.PORT ? Number(process.env.PORT) : undefined,
  },
  azure: {
    tenantId: process.env.AZURE_TENANT_ID ?? undefined,
    clientId: process.env.AZURE_CLIENT_ID ?? undefined,
    clientSecret: readSecretFromEnv('AZURE_CLIENT_SECRET') ?? undefined,
    scope: process.env.AZURE_SCOPE ?? undefined,
  },
  sharepoint: {
    siteDisplayName: process.env.SITE_DISPLAY_NAME ?? undefined,
    siteType: (process.env.SITE_TYPE as SiteType | undefined) ?? undefined,
    host: process.env.SHAREPOINT_HOST ?? undefined,
  },
  openai: {
    apiKey: readSecretFromEnv('OPENAI_API_KEY') ?? undefined,
    model: process.env.OPENAI_MODEL ?? undefined,
  },
  redis: {
    url: process.env.REDIS_URL ?? undefined,
    tls: parseBoolean(process.env.REDIS_TLS),
    keyPrefix: process.env.REDIS_KEY_PREFIX ?? undefined,
  },
  auth: {
    tokenCacheTtlMinutes: process.env.AUTH_TOKEN_CACHE_TTL_MINUTES
      ? Number(process.env.AUTH_TOKEN_CACHE_TTL_MINUTES)
      : undefined,
    tokenFallbackMs: process.env.AUTH_TOKEN_FALLBACK_MS
      ? Number(process.env.AUTH_TOKEN_FALLBACK_MS)
      : undefined,
  },
  tenants: loadJsonValue('TENANTS_JSON', 'TENANTS_CONFIG_PATH', 'TENANTS_JSON') as TenantConfig[] | undefined,
  apiKeys: loadJsonValue('API_KEYS_JSON', 'API_KEYS_CONFIG_PATH', 'API_KEYS_JSON') as ApiKeyConfig[] | undefined,
};

const runtimeOverrides = (() => {
  try {
    return loadRuntimeOverrides();
  } catch {
    return undefined;
  }
})();

const mergedConfig = mergeConfig<ConfigInput>(
  DEFAULT_CONFIG as unknown as ConfigInput,
  envOverrides,
  runtimeOverrides as ConfigInput | undefined,
);

let parsedConfig: AppConfig;

try {
  parsedConfig = ConfigSchema.parse(mergedConfig);
} catch (error) {
  if (error instanceof z.ZodError) {
    const details = error.issues
      .map((issue) => `- ${issue.path.join('.') || '(root)'}: ${issue.message}`)
      .join('\n');
    throw new Error(`Configuration validation failed:\n${details}`);
  }
  throw error;
}

export const CONFIG: AppConfig = Object.freeze({
  ...parsedConfig,
  sharepoint: {
    ...parsedConfig.sharepoint,
    host:
      parsedConfig.sharepoint.host ||
      parsedConfig.tenants.find((tenant) => tenant.active && tenant.sharepoint.host)?.sharepoint.host ||
      undefined,
    siteDisplayName:
      parsedConfig.sharepoint.siteDisplayName ||
      parsedConfig.tenants.find((tenant) => tenant.active && tenant.sharepoint.siteDisplayName)?.sharepoint
        .siteDisplayName ||
      DEFAULT_CONFIG.sharepoint.siteDisplayName,
  },
  azure: {
    ...parsedConfig.azure,
    tokenEndpoint: `https://login.microsoftonline.com/${parsedConfig.azure.tenantId}/oauth2/v2.0/token`,
  },
});

export type AppConfigType = typeof CONFIG;
