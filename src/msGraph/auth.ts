import { ConfidentialClientApplication } from '@azure/msal-node';
import fetch, { type RequestInit } from 'node-fetch';

import { CONFIG } from '../config.js';
import { logger } from '../logger.js';

const scopes = ['https://graph.microsoft.com/.default'];

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    clientSecret: CONFIG.clientSecret
  }
});

export async function graphFetch(path: string, init: RequestInit = {}): Promise<any> {
  const token = await cca.acquireTokenByClientCredential({ scopes });
  if (!token?.accessToken) {
    throw new Error('Failed to acquire Microsoft Graph token');
  }

  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token.accessToken}`,
      'Content-Type': 'application/json',
      ...(init.headers ?? {})
    }
  });

  const text = await response.text();
  let body: any = null;
  if (text) {
    try {
      body = JSON.parse(text);
    } catch (error) {
      logger.warn('Failed to parse Graph response as JSON', { path, text, error: (error as Error).message });
      body = text;
    }
  }

  if (!response.ok) {
    logger.error('Graph request failed', { path, status: response.status, body });
    throw new Error(`Graph request failed: ${response.status}`);
  }

  return body;
}
