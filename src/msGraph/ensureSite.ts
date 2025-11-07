import { graphFetch, sharepointFetch, HttpError } from './auth.js';
import { logger } from '../logger.js';
import { CONFIG, type SiteType } from '../config.js';

interface Group {
  id: string;
  displayName: string;
  mailNickname: string;
}

interface SiteResponse {
  id: string;
  webUrl: string;
  displayName?: string;
}

export interface SiteContext {
  siteId: string;
  siteUrl: string;
  siteType: SiteType;
  groupId?: string;
}

const SITE_PROVISION_ATTEMPTS = 10;
const SITE_PROVISION_DELAY_MS = 5000;

function slugify(name: string): string {
  return name
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .substring(0, 60);
}

async function pollForSite(path: string): Promise<SiteResponse> {
  for (let attempt = 0; attempt < SITE_PROVISION_ATTEMPTS; attempt += 1) {
    try {
      const site = (await graphFetch(path)) as SiteResponse;
      if (site?.id) {
        return site;
      }
    } catch (error) {
      logger.warn('Waiting for site provisioning', {
        attempt,
        error: (error as Error).message
      });
    }
    await new Promise((resolve) => setTimeout(resolve, SITE_PROVISION_DELAY_MS));
  }
  throw new Error('Timed out waiting for SharePoint site provisioning');
}

export async function ensureTeamSite(displayName: string): Promise<SiteContext> {
  const escaped = displayName.replace(/'/g, "''");
  const existing = await graphFetch(`/groups?$filter=displayName eq '${escaped}'`);
  if (existing.value?.length) {
    const group = existing.value[0] as Group;
    const site = (await graphFetch(`/groups/${group.id}/sites/root`)) as SiteResponse;
    logger.info('Found existing team site', { displayName, siteId: site.id });
    return { groupId: group.id, siteId: site.id, siteUrl: site.webUrl, siteType: 'team' };
  }

  const group = (await graphFetch('/groups', {
    method: 'POST',
    body: JSON.stringify({
      displayName,
      mailNickname: displayName.toLowerCase().replace(/\s+/g, '-'),
      groupTypes: ['Unified'],
      mailEnabled: true,
      securityEnabled: false
    })
  })) as Group;

  logger.info('Created unified group, waiting for SharePoint site provisioning', { groupId: group.id });
  const site = await pollForSite(`/groups/${group.id}/sites/root`);
  return { groupId: group.id, siteId: site.id, siteUrl: site.webUrl, siteType: 'team' };
}

async function getSiteByPath(hostname: string, relativePath: string): Promise<SiteResponse | null> {
  try {
    return (await graphFetch(`/sites/${hostname}:${relativePath}`)) as SiteResponse;
  } catch (error) {
    if (error instanceof HttpError && error.status === 404) {
      return null;
    }
    throw error;
  }
}

function resolveCommunicationPath(displayName: string, sitePath?: string): string {
  if (sitePath) {
    const trimmed = sitePath.startsWith('/') ? sitePath : `/${sitePath}`;
    if (!trimmed.startsWith('/sites/')) {
      throw new Error('Communication site paths must start with /sites/');
    }
    return trimmed;
  }

  const slug = slugify(displayName || 'site');
  return `/sites/${slug}`;
}

export async function ensureCommunicationSite(
  displayName: string,
  host: string,
  sitePath?: string,
): Promise<SiteContext> {
  let url: URL;
  try {
    url = new URL(host);
  } catch (error) {
    throw new Error('SharePoint host must be a valid absolute URL');
  }

  const relativePath = resolveCommunicationPath(displayName, sitePath);

  const existing = await getSiteByPath(url.hostname, relativePath);
  if (existing?.id) {
    logger.info('Found existing communication site', { displayName, siteId: existing.id });
    return { siteId: existing.id, siteUrl: existing.webUrl, siteType: 'communication' };
  }

  logger.info('Creating communication site via SPSiteManager', { displayName, relativePath });
  await sharepointFetch(
    '/_api/SPSiteManager/Create',
    {
      method: 'POST',
      body: JSON.stringify({
        request: {
          Title: displayName,
          Url: `${url.origin}${relativePath}`,
          Lcid: 1033,
          ShareByEmailEnabled: true,
          Description: 'Elion Studio communication site',
          WebTemplate: 'SITEPAGEPUBLISHING#0',
          SiteDesignId: '00000000-0000-0000-0000-000000000000'
        }
      })
    },
    { host },
  );

  const site = await pollForSite(`/sites/${url.hostname}:${relativePath}`);
  return { siteId: site.id, siteUrl: site.webUrl, siteType: 'communication' };
}

export interface EnsureSiteOptions {
  displayName: string;
  siteType: SiteType;
  host?: string;
  sitePath?: string;
}

export async function ensureSite(options: EnsureSiteOptions): Promise<SiteContext> {
  if (options.siteType === 'communication') {
    const host = options.host ?? CONFIG.sharepoint.host;
    if (!host) {
      throw new Error('SharePoint host must be configured for communication sites');
    }
    return ensureCommunicationSite(options.displayName, host, options.sitePath);
  }

  return ensureTeamSite(options.displayName);
}

