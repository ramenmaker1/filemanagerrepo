import fetch, { type RequestInit } from 'node-fetch';

import { CONFIG } from '../config.js';
import { logger } from '../logger.js';
import { getRedisValue, setRedisValue } from '../redis.js';

export class HttpError extends Error {
  status: number;

  body: unknown;

  constructor(status: number, statusText: string, body: unknown) {
    super(`Request failed: ${status} ${statusText}`);
    this.status = status;
    this.body = body;
  }
}

interface TokenCacheEntry {
  token: string;
  expiresAt: number;
  cachedAt: number;
}

const TOKEN_BUFFER_MS = 60_000;
const REDIS_CACHE_KEY = `${CONFIG.redis.keyPrefix}:graph:token:${CONFIG.azure.tenantId}:${CONFIG.azure.clientId}`;

let inMemoryToken: TokenCacheEntry | null = null;
let consecutiveTokenFailures = 0;
let fallbackMode = false;

function isTokenValid(entry: TokenCacheEntry | null, allowStale = false): boolean {
  if (!entry) {
    return false;
  }

  const now = Date.now();
  if (allowStale) {
    return now - entry.expiresAt < CONFIG.auth.tokenFallbackMs;
  }

  return entry.expiresAt - TOKEN_BUFFER_MS > now;
}

async function readTokenFromRedis(): Promise<TokenCacheEntry | null> {
  const raw = await getRedisValue(REDIS_CACHE_KEY);
  if (!raw) {
    return null;
  }

  try {
    const parsed = JSON.parse(raw) as TokenCacheEntry;
    if (typeof parsed.token === 'string' && typeof parsed.expiresAt === 'number') {
      return parsed;
    }
  } catch (error) {
    logger.error('Failed to parse token cache entry from Redis', { error: (error as Error).message });
  }
  return null;
}

async function writeTokenCache(entry: TokenCacheEntry): Promise<void> {
  const ttlMs = Math.min(
    CONFIG.auth.tokenCacheTtlMinutes * 60_000,
    Math.max(60_000, entry.expiresAt - Date.now()),
  );
  await setRedisValue(REDIS_CACHE_KEY, JSON.stringify(entry), ttlMs);
}

async function acquireNewToken(): Promise<TokenCacheEntry> {
  const response = await fetch(CONFIG.azure.tokenEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({
      client_id: CONFIG.azure.clientId,
      client_secret: CONFIG.azure.clientSecret,
      grant_type: 'client_credentials',
      scope: CONFIG.azure.scope,
    }).toString(),
  });

  const text = await response.text();
  let body: any = null;
  if (text) {
    try {
      body = JSON.parse(text);
    } catch (error) {
      logger.warn('Token endpoint returned non-JSON response', { error: (error as Error).message });
      body = text;
    }
  }

  if (!response.ok) {
    throw new Error(
      `Token request failed: ${response.status} ${response.statusText} ${(body && body.error_description) || ''}`.trim(),
    );
  }

  const expiresIn = Number(body?.expires_in ?? 3600);
  const now = Date.now();
  const entry: TokenCacheEntry = {
    token: body.access_token as string,
    expiresAt: now + expiresIn * 1000,
    cachedAt: now,
  };
  inMemoryToken = entry;
  await writeTokenCache(entry);
  return entry;
}

async function getAccessToken(): Promise<string> {
  if (isTokenValid(inMemoryToken)) {
    return inMemoryToken!.token;
  }

  const redisEntry = await readTokenFromRedis();
  if (isTokenValid(redisEntry)) {
    inMemoryToken = redisEntry;
    return redisEntry!.token;
  }

  try {
    const entry = await acquireNewToken();
    consecutiveTokenFailures = 0;
    fallbackMode = false;
    return entry.token;
  } catch (error) {
    consecutiveTokenFailures += 1;
    logger.error('Failed to acquire Microsoft Graph token', {
      attempts: consecutiveTokenFailures,
      error: (error as Error).message,
    });

    if (consecutiveTokenFailures >= 3) {
      logger.error('Graph token acquisition repeated failures; consider investigating credentials or network issues');
    }

    if (consecutiveTokenFailures >= 5 && !fallbackMode) {
      fallbackMode = true;
      logger.error('Entering token fallback mode - using last known token if available');
    }

    if (fallbackMode && isTokenValid(inMemoryToken, true)) {
      logger.warn('Using stale Microsoft Graph token from fallback cache', {
        ageMs: Date.now() - (inMemoryToken?.cachedAt ?? Date.now()),
      });
      return inMemoryToken!.token;
    }

    throw error;
  }
}

async function fetchWithToken(
  baseUrl: string,
  path: string,
  init: RequestInit = {},
  defaultHeaders: Record<string, string> = {},
): Promise<any> {
  const token = await getAccessToken();

  const response = await fetch(`${baseUrl}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      ...defaultHeaders,
      ...(init.headers ?? {}),
    },
  });

  const text = await response.text();
  let body: any = null;
  if (text) {
    try {
      body = JSON.parse(text);
    } catch (error) {
      logger.warn('Failed to parse response as JSON', {
        baseUrl,
        path,
        text,
        error: (error as Error).message,
      });
      body = text;
    }
  }

  if (!response.ok) {
    logger.error('Request failed', {
      baseUrl,
      path,
      status: response.status,
      body,
    });
    throw new HttpError(response.status, response.statusText, body);
  }

  return body;
}

export async function graphFetch(path: string, init: RequestInit = {}): Promise<any> {
  return fetchWithToken('https://graph.microsoft.com/v1.0', path, init, {
    'Content-Type': 'application/json',
  });
}

export async function sharepointFetch(
  path: string,
  init: RequestInit = {},
  options: { host?: string } = {},
): Promise<any> {
  const host = options.host ?? CONFIG.sharepoint.host;
  if (!host) {
    throw new Error('SharePoint host must be configured to call SharePoint REST APIs');
  }

  return fetchWithToken(host, path, init, {
    'Content-Type': 'application/json;odata=verbose',
    Accept: 'application/json;odata=verbose',
  });
}
