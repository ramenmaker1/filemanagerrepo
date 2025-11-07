import { ensureSite, type EnsureSiteOptions, type SiteContext } from './ensureSite.js';
import { ensureLibraries } from './libraries.js';
import { ensurePermissions } from './permissions.js';
import { ensureCatalogPage } from './catalog.js';
import { createExpiringLink } from './sharing.js';
import { listFolder, resolveDriveId, resolveItemId } from './listing.js';
import { logger } from '../logger.js';
import { CONFIG, type SiteType } from '../config.js';
import { type TenantRecord } from '../tenants.js';
import { type MsGraphAction } from './schema.js';

function buildSiteOptions(
  tenant: TenantRecord,
  override?: { siteName?: string; siteType?: SiteType },
): EnsureSiteOptions {
  const displayName =
    override?.siteName || tenant.sharepoint.siteDisplayName || CONFIG.sharepoint.siteDisplayName || tenant.name;
  return {
    displayName,
    siteType: override?.siteType ?? tenant.sharepoint.siteType ?? CONFIG.sharepoint.siteType,
    host: tenant.sharepoint.host ?? CONFIG.sharepoint.host,
    sitePath: tenant.sharepoint.sitePath,
  };
}

export async function dispatchGraphAction(
  tenant: TenantRecord,
  payload: MsGraphAction,
): Promise<Record<string, unknown>> {
  const resolveContext = async (override?: { siteName?: string; siteType?: SiteType }): Promise<SiteContext> => {
    return ensureSite(buildSiteOptions(tenant, override));
  };

  switch (payload.action) {
    case 'ensure_site': {
      const context = await resolveContext({
        siteName: payload.siteName ?? tenant.sharepoint.siteDisplayName,
        siteType: payload.siteType,
      });
      await ensureLibraries(context.siteId);
      await ensurePermissions(context.siteId);
      return {
        tenantId: tenant.id,
        siteId: context.siteId,
        groupId: context.groupId,
        siteUrl: context.siteUrl,
        siteType: context.siteType,
      };
    }
    case 'ensure_libraries': {
      const context = await resolveContext();
      await ensureLibraries(context.siteId);
      return {
        tenantId: tenant.id,
        siteId: context.siteId,
        siteUrl: context.siteUrl,
        siteType: context.siteType,
      };
    }
    case 'ensure_groups_permissions': {
      const context = await resolveContext();
      await ensurePermissions(context.siteId);
      return {
        tenantId: tenant.id,
        siteId: context.siteId,
        siteUrl: context.siteUrl,
        siteType: context.siteType,
      };
    }
    case 'create_catalog_page': {
      const context = await resolveContext({ siteName: payload.siteName });
      const url = await ensureCatalogPage(context.siteId, 'Catalog', payload.catalogLinks ?? {});
      return {
        tenantId: tenant.id,
        catalogUrl: url,
        siteId: context.siteId,
        siteUrl: context.siteUrl,
      };
    }
    case 'share_deliverable': {
      const context = await resolveContext({ siteName: payload.siteName });
      const driveId = await resolveDriveId(context.siteId, 'Deliverables');
      const itemId = await resolveItemId(context.siteId, driveId, payload.driveItemPath);
      const webUrl = await createExpiringLink({
        driveId,
        itemId,
        type: payload.shareType ?? 'view',
        expiresAt: payload.expiresAt,
      });
      return {
        tenantId: tenant.id,
        link: webUrl,
        siteId: context.siteId,
        siteUrl: context.siteUrl,
      };
    }
    case 'list_folder': {
      const context = await resolveContext({ siteName: payload.siteName });
      const driveId = await resolveDriveId(context.siteId, payload.libraryName);
      const items = await listFolder(context.siteId, driveId, payload.driveItemPath ?? '/');
      return {
        tenantId: tenant.id,
        items,
        siteId: context.siteId,
        siteUrl: context.siteUrl,
      };
    }
    case 'link_repo_and_base44': {
      logger.info('Linking repo/Base44 metadata', {
        repoUrl: payload.repoUrl,
        base44Url: payload.base44Url,
        sharepointUrl: payload.sharepointUrl,
      });
      return {
        message: 'Update GitHub README and Base44 metadata with provided SharePoint URL',
        repoUrl: payload.repoUrl,
        base44Url: payload.base44Url,
        sharepointUrl: payload.sharepointUrl,
        tenantId: tenant.id,
      };
    }
    default:
      return { message: 'Unsupported action' };
  }
}
