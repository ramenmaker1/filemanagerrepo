import { ensureSite, type SiteContext } from './ensureSite.js';
import { ensureLibraries } from './libraries.js';
import { ensurePermissions } from './permissions.js';
import { ensureCatalogPage } from './catalog.js';
import { createExpiringLink } from './sharing.js';
import { listFolder, resolveDriveId, resolveItemId } from './listing.js';
import { logger } from '../logger.js';
import { CONFIG, type SiteType } from '../config.js';
import { type MsGraphAction } from './schema.js';

export async function dispatchGraphAction(
  displayName: string,
  payload: MsGraphAction
): Promise<Record<string, unknown>> {
  const resolveContext = async (override?: { siteName?: string; siteType?: SiteType }): Promise<SiteContext> => {
    return ensureSite(override?.siteName ?? displayName, override?.siteType ?? CONFIG.siteType);
  };

  switch (payload.action) {
    case 'ensure_site': {
      const context = await resolveContext({
        siteName: payload.siteName ?? displayName,
        siteType: payload.siteType
      });
      await ensureLibraries(context.siteId);
      await ensurePermissions(context.siteId);
      return {
        siteId: context.siteId,
        groupId: context.groupId,
        siteUrl: context.siteUrl,
        siteType: context.siteType
      };
    }
    case 'ensure_libraries': {
      const context = await resolveContext();
      await ensureLibraries(context.siteId);
      return { siteId: context.siteId, siteUrl: context.siteUrl, siteType: context.siteType };
    }
    case 'ensure_groups_permissions': {
      const context = await resolveContext();
      await ensurePermissions(context.siteId);
      return { siteId: context.siteId, siteUrl: context.siteUrl, siteType: context.siteType };
    }
    case 'create_catalog_page': {
      const context = await resolveContext();
      const url = await ensureCatalogPage(context.siteId, 'Catalog', payload.catalogLinks ?? {});
      return { catalogUrl: url, siteId: context.siteId, siteUrl: context.siteUrl };
    }
    case 'share_deliverable': {
      const context = await resolveContext();
      const driveId = await resolveDriveId(context.siteId, 'Deliverables');
      const itemId = await resolveItemId(context.siteId, driveId, payload.driveItemPath);
      const webUrl = await createExpiringLink({
        driveId,
        itemId,
        type: payload.shareType ?? 'view',
        expiresAt: payload.expiresAt
      });
      return { link: webUrl, siteId: context.siteId, siteUrl: context.siteUrl };
    }
    case 'list_folder': {
      const context = await resolveContext();
      const driveId = await resolveDriveId(context.siteId, payload.libraryName);
      const items = await listFolder(context.siteId, driveId, payload.driveItemPath ?? '/');
      return { items, siteId: context.siteId, siteUrl: context.siteUrl };
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
