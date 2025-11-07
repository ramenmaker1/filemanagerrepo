import Redis from 'ioredis';

import { CONFIG } from './config.js';
import { logger } from './logger.js';

let redisClient: Redis | null = null;
let initializing: Promise<Redis | null> | null = null;

export async function getRedisClient(): Promise<Redis | null> {
  if (!CONFIG.redis.url) {
    return null;
  }

  if (redisClient) {
    return redisClient;
  }

  if (initializing) {
    return initializing;
  }

  initializing = (async () => {
    try {
      const client = new Redis(CONFIG.redis.url, {
        enableAutoPipelining: true,
        tls: CONFIG.redis.tls ? {} : undefined,
      });

      client.on('error', (error) => {
        logger.error('Redis connection error', { error: error.message });
      });

      client.on('end', () => {
        logger.warn('Redis connection closed');
        redisClient = null;
        initializing = null;
      });

      await client.ping();
      redisClient = client;
      logger.info('Redis connection established');
      return client;
    } catch (error) {
      logger.error('Failed to connect to Redis', { error: (error as Error).message });
      redisClient = null;
      initializing = null;
      return null;
    }
  })();

  return initializing;
}

export async function setRedisValue(key: string, value: string, ttlMs: number): Promise<void> {
  const client = await getRedisClient();
  if (!client) {
    return;
  }

  try {
    await client.set(key, value, 'PX', ttlMs);
  } catch (error) {
    logger.error('Failed to set Redis value', {
      key,
      error: (error as Error).message,
    });
  }
}

export async function getRedisValue(key: string): Promise<string | null> {
  const client = await getRedisClient();
  if (!client) {
    return null;
  }

  try {
    return await client.get(key);
  } catch (error) {
    logger.error('Failed to read Redis value', {
      key,
      error: (error as Error).message,
    });
    return null;
  }
}
