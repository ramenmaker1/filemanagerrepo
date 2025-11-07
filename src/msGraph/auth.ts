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

export class HttpError extends Error {
  status: number;

  body: unknown;

  constructor(status: number, statusText: string, body: unknown) {
    super(`Request failed: ${status} ${statusText}`);
    this.status = status;
    this.body = body;
  }
}

async function fetchWithToken(
  baseUrl: string,
  path: string,
  init: RequestInit = {},
  defaultHeaders: Record<string, string> = {}
): Promise<any> {
  const token = await cca.acquireTokenByClientCredential({ scopes });
  if (!token?.accessToken) {
    throw new Error('Failed to acquire Microsoft Graph token');
  }

  const response = await fetch(`${baseUrl}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token.accessToken}`,
      ...defaultHeaders,
      ...(init.headers ?? {})
    }
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
        error: (error as Error).message
      });
      body = text;
    }
  }

  if (!response.ok) {
    logger.error('Request failed', {
      baseUrl,
      path,
      status: response.status,
      body
    });
    throw new HttpError(response.status, response.statusText, body);
  }

  return body;
}

export async function graphFetch(path: string, init: RequestInit = {}): Promise<any> {
  return fetchWithToken('https://graph.microsoft.com/v1.0', path, init, {
    'Content-Type': 'application/json'
  });
}

export async function sharepointFetch(path: string, init: RequestInit = {}): Promise<any> {
  if (!CONFIG.sharepointHost) {
    throw new Error('SHAREPOINT_HOST must be configured to call SharePoint REST APIs');
  }

  return fetchWithToken(CONFIG.sharepointHost, path, init, {
    'Content-Type': 'application/json;odata=verbose',
    Accept: 'application/json;odata=verbose'
  });
}
