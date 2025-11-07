import { graphFetch } from './auth.js';
import { logger } from '../logger.js';

const EDITORS_GROUP = 'Elion-Editors';
const VIEWERS_GROUP = 'Elion-Viewers';

async function ensureSecurityGroup(displayName: string): Promise<{ id: string }> {
  const escaped = displayName.replace(/'/g, "''");
  const existing = await graphFetch(`/groups?$filter=displayName eq '${escaped}'`);
  if (existing.value?.length) {
    return existing.value[0] as { id: string };
  }

  const group = await graphFetch('/groups', {
    method: 'POST',
    body: JSON.stringify({
      displayName,
      mailEnabled: false,
      securityEnabled: true,
      mailNickname: displayName.toLowerCase().replace(/\s+/g, '-')
    })
  });
  logger.info('Created security group', { displayName, groupId: group.id });
  return group as { id: string };
}

async function findDriveByName(siteId: string, libraryName: string): Promise<{ id: string } | null> {
  const drives = await graphFetch(`/sites/${siteId}/drives`);
  return (drives.value as Array<{ id: string; name: string }>).find((drive) => drive.name === libraryName) ?? null;
}

export async function ensurePermissions(siteId: string): Promise<void> {
  const editors = await ensureSecurityGroup(EDITORS_GROUP);
  const viewers = await ensureSecurityGroup(VIEWERS_GROUP);

  const legalDrive = await findDriveByName(siteId, 'Legal & Finance');
  if (!legalDrive) {
    logger.warn('Legal & Finance drive not found to set permissions');
    return;
  }

  try {
    await graphFetch(`/drives/${legalDrive.id}/root/listItem/breakRoleInheritance`, {
      method: 'POST',
      body: JSON.stringify({ retainInheritedPermissions: false })
    });
  } catch (error) {
    logger.warn('Failed to break inheritance (may already be broken or unsupported)', {
      error: (error as Error).message
    });
  }

  const grantPermissions = async (groupId: string, roles: string[]) => {
    await graphFetch(`/drives/${legalDrive.id}/items/root/invite`, {
      method: 'POST',
      body: JSON.stringify({
        requireSignIn: true,
        sendInvitation: false,
        roles,
        recipients: [
          {
            objectId: groupId
          }
        ]
      })
    });
  };

  await grantPermissions(editors.id, ['write']);
  await grantPermissions(viewers.id, ['read']);
}

export { EDITORS_GROUP, VIEWERS_GROUP };
