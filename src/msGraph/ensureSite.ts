import { graphFetch } from './auth.js';
import { logger } from '../logger.js';

export interface Group {
  id: string;
  displayName: string;
  mailNickname: string;
  sites?: { value: Array<{ id: string }> };
}

export async function ensureTeamSite(displayName: string): Promise<{ group: Group; siteId: string }>
{
  const escaped = displayName.replace(/'/g, "''");
  const existing = await graphFetch(`/groups?$filter=displayName eq '${escaped}'`);
  if (existing.value?.length) {
    const group = existing.value[0] as Group;
    const site = await graphFetch(`/groups/${group.id}/sites/root`);
    logger.info('Found existing team site', { displayName, siteId: site.id });
    return { group, siteId: site.id as string };
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

  // Poll for site creation
  let siteId = '';
  for (let attempt = 0; attempt < 10; attempt += 1) {
    try {
      const site = await graphFetch(`/groups/${group.id}/sites/root`);
      if (site?.id) {
        siteId = site.id as string;
        break;
      }
    } catch (error) {
      logger.warn('Waiting for site provisioning', { attempt, error: (error as Error).message });
    }
    await new Promise((resolve) => setTimeout(resolve, 5000));
  }

  if (!siteId) {
    throw new Error('Timed out waiting for SharePoint site provisioning');
  }

  return { group, siteId };
}
