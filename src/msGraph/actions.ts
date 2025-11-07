import { ensureTeamSite } from './ensureSite.js';
import { ensureLibraries } from './libraries.js';
import { ensurePermissions } from './permissions.js';
import { ensureCatalogPage } from './catalog.js';
import { createExpiringLink } from './sharing.js';
import { listFolder, resolveDriveId, resolveItemId } from './listing.js';
import { logger } from '../logger.js';
import { CONFIG } from '../config.js';
import { type MsGraphAction } from './schema.js';

export async function dispatchGraphAction(
  displayName: string,
  payload: MsGraphAction
): Promise<Record<string, unknown>> {
  switch (payload.action) {
    case 'ensure_site': {
      const { siteId, group } = await ensureTeamSite(payload.siteName ?? displayName);
      await ensureLibraries(siteId);
      await ensurePermissions(siteId);
      return { siteId, groupId: group.id };
    }
    case 'ensure_libraries': {
      const { siteId } = await ensureTeamSite(displayName);
      await ensureLibraries(siteId);
      return { siteId };
    }
    case 'ensure_groups_permissions': {
      const { siteId } = await ensureTeamSite(displayName);
      await ensurePermissions(siteId);
      return { siteId };
    }
    case 'create_catalog_page': {
      const { siteId } = await ensureTeamSite(displayName);
      const url = await ensureCatalogPage(siteId, 'Catalog', payload.catalogLinks);
      return { catalogUrl: url };
    }
    case 'share_deliverable': {
      const { siteId } = await ensureTeamSite(displayName);
      const driveId = await resolveDriveId(siteId, 'Deliverables');
      const itemId = await resolveItemId(siteId, driveId, payload.driveItemPath);
      const webUrl = await createExpiringLink({
        driveId,
        itemId,
        type: payload.shareType ?? 'view',
        expiresAt: payload.expiresAt
      });
      return { link: webUrl };
    }
    case 'list_folder': {
      const { siteId } = await ensureTeamSite(displayName);
      const driveId = await resolveDriveId(siteId, payload.libraryName);
      const items = await listFolder(siteId, driveId, payload.driveItemPath ?? '/');
      return { items };
    }
    case 'link_repo_and_base44': {
      logger.info('Linking repo/Base44 metadata', {
        repoUrl: payload.repoUrl,
        base44Url: payload.base44Url,
        sharepointUrl: payload.sharepointUrl
      });
      return {
        message: 'Update GitHub README and Base44 metadata with provided SharePoint URL',
        repoUrl: payload.repoUrl,
        base44Url: payload.base44Url,
        sharepointUrl: payload.sharepointUrl
      };
    }
    default:
      return { message: 'Unsupported action' };
  }
}

export const DEFAULT_DISPLAY_NAME = CONFIG.siteDisplayName;
