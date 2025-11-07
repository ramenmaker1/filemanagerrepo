import express, { type NextFunction, type Request, type Response } from 'express';

import pkg from '../package.json' assert { type: 'json' };

import { CONFIG } from './config.js';
import { logger } from './logger.js';
import { graphFetch, sharepointFetch } from './msGraph/auth.js';
import { getRedisClient } from './redis.js';
import { authenticationMiddleware } from './middleware/auth.js';
import { getDefaultTenant } from './tenants.js';
import catalog from './routes/catalog.js';
import list from './routes/list.js';
import provision from './routes/provision.js';
import share from './routes/share.js';
import agent from './routes/agent.js';

const app = express();
app.use(express.json());

function buildErrorEnvelope(error: Error, _status: number, req: Request) {
  return {
    error: error.name || 'Error',
    message: error.message || 'Unexpected error',
    requestId: (req as any).id ?? 'n/a',
    timestamp: new Date().toISOString(),
    path: req.path,
    details: (error as any).details ?? undefined,
    trace_id: (req as any).traceId ?? undefined,
  };
}

app.get('/healthz', (_req, res) => {
  res.json({
    status: 'ok',
    version: pkg.version ?? '0.0.0',
    uptime: process.uptime(),
    environment: CONFIG.environment,
  });
});

type CheckResult = {
  name: string;
  status: 'pass' | 'fail' | 'skipped';
  latency_ms: number;
  updated_at: string;
  message?: string;
};

async function runCheck(name: string, fn: () => Promise<void>): Promise<CheckResult> {
  const startedAt = Date.now();
  try {
    await fn();
    return {
      name,
      status: 'pass' as const,
      latency_ms: Date.now() - startedAt,
      updated_at: new Date().toISOString(),
    };
  } catch (error) {
    return {
      name,
      status: 'fail' as const,
      latency_ms: Date.now() - startedAt,
      message: (error as Error).message,
      updated_at: new Date().toISOString(),
    };
  }
}

function skippedCheck(name: string, message: string): CheckResult {
  return {
    name,
    status: 'skipped',
    latency_ms: 0,
    message,
    updated_at: new Date().toISOString(),
  };
}

app.get('/readyz', async (_req, res) => {
  const overallStart = Date.now();

  const defaultTenant = getDefaultTenant();
  const sharepointHost = defaultTenant?.sharepoint.host ?? CONFIG.sharepoint.host;

  const checks = await Promise.all([
    runCheck('graph', async () => {
      await graphFetch('/sites/root?$select=id');
    }),
    sharepointHost
      ? runCheck('sharepoint', async () => {
          await sharepointFetch('/_api/web?$select=Id', {}, { host: sharepointHost });
        })
      : Promise.resolve(skippedCheck('sharepoint', 'SHAREPOINT_HOST not configured')),
    CONFIG.redis.url
      ? runCheck('redis', async () => {
          const client = await getRedisClient();
          if (!client) {
            throw new Error('Redis unavailable');
          }
          await client.ping();
        })
      : Promise.resolve(skippedCheck('redis', 'Redis cache not configured')),
  ]);

  const hasFailure = checks.some((check) => check.status === 'fail');
  const status = hasFailure ? 'degraded' : 'ready';

  res.status(hasFailure ? 503 : 200).json({
    status,
    checks,
    overall_latency_ms: Date.now() - overallStart,
  });
});

app.use(authenticationMiddleware);

app.use('/provision', provision);
app.use('/share', share);
app.use('/catalog', catalog);
app.use('/list', list);
app.use('/agent', agent);

app.use((err: Error, req: Request, res: Response, _next: NextFunction) => {
  logger.error('Unhandled error', { error: err.message, stack: err.stack });
  const status = res.statusCode >= 400 ? res.statusCode : 500;
  res.status(status).json(buildErrorEnvelope(err, status, req));
});

app.listen(CONFIG.server.port, () => {
  logger.info(`Elion SP Agent listening on :${CONFIG.server.port}`);
});
