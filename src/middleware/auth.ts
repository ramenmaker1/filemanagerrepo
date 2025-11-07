import type { NextFunction, Request, Response } from 'express';

import { logger } from '../logger.js';
import { authenticateApiKey, type ApiKeyAuthResult } from '../tenants.js';

export interface RequestApiKeyContext {
  id: string;
  name: string;
  roles: readonly string[];
  tenantId: string;
  secretId?: string;
  presentedHash: string;
  expiresAt?: string;
  metadata?: Record<string, unknown>;
}

const PUBLIC_PATHS = new Set(['/healthz', '/readyz']);

function buildEnvelope(req: Request, message: string, code: ApiKeyAuthResult['code'], status: number) {
  const details = { code } as Record<string, unknown>;
  return {
    error: status === 401 ? 'UnauthorizedError' : status >= 500 ? 'InternalServerError' : 'ForbiddenError',
    message,
    requestId: (req as any).id ?? 'n/a',
    timestamp: new Date().toISOString(),
    path: req.path,
    details,
  };
}

function failureStatus(code: ApiKeyAuthResult['code']): number {
  switch (code) {
    case 'invalid':
    case 'api_key_expired':
    case 'secret_expired':
      return 401;
    case 'tenant_missing':
      return 500;
    default:
      return 403;
  }
}

export function authenticationMiddleware(req: Request, res: Response, next: NextFunction) {
  if (PUBLIC_PATHS.has(req.path)) {
    return next();
  }

  const authorization = req.headers.authorization;
  let candidate: string | undefined;

  if (typeof authorization === 'string') {
    const bearer = authorization.match(/^Bearer\s+(.+)$/i);
    if (bearer) {
      candidate = bearer[1];
    }
  }

  if (!candidate) {
    const apiKeyHeader = req.headers['x-api-key'];
    if (typeof apiKeyHeader === 'string') {
      candidate = apiKeyHeader;
    } else if (Array.isArray(apiKeyHeader) && apiKeyHeader.length > 0) {
      candidate = apiKeyHeader[0];
    }
  }

  const authResult = authenticateApiKey(candidate);

  if (!authResult.ok) {
    const status = failureStatus(authResult.code);
    const message =
      status >= 500
        ? 'Authentication subsystem misconfigured'
        : authResult.message || 'Authentication failed';

    const envelope = buildEnvelope(req, message, authResult.code, status);

    logger.warn('API authentication failed', {
      path: req.path,
      status,
      code: authResult.code,
      tenantId: authResult.tenantId,
      apiKeyId: authResult.apiKeyId,
    });

    if (status === 401) {
      res.setHeader('WWW-Authenticate', 'Bearer realm="elion-studio-api", error="invalid_token"');
    }

    return res.status(status).json(envelope);
  }

  const context: RequestApiKeyContext = {
    id: authResult.apiKey.id,
    name: authResult.apiKey.name,
    roles: authResult.apiKey.roles,
    tenantId: authResult.tenant.id,
    secretId: authResult.secret.id,
    presentedHash: authResult.presentedHash,
    expiresAt: authResult.apiKey.expiresAt,
    metadata: authResult.apiKey.metadata,
  };

  (req as any).apiKey = context;
  (req as any).tenant = authResult.tenant;
  res.locals.tenantId = authResult.tenant.id;
  res.locals.apiKeyId = authResult.apiKey.id;
  res.setHeader('X-Tenant-Id', authResult.tenant.id);
  res.setHeader('X-Api-Key-Id', authResult.apiKey.id);

  return next();
}
