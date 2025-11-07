import { graphFetch } from './auth.js';

export async function listFolder(siteId: string, driveId: string, path: string): Promise<any> {
  const targetPath = path.replace(/^\/+/, '');
  if (!targetPath || targetPath === '.') {
    const data = await graphFetch(`/sites/${siteId}/drives/${driveId}/root/children`);
    return data?.value ?? [];
  }
  const encoded = encodeURIComponent(targetPath);
  const data = await graphFetch(`/sites/${siteId}/drives/${driveId}/root:/${encoded}:/children`);
  return data?.value ?? [];
}

export async function resolveDriveId(siteId: string, libraryName: string): Promise<string> {
  const drives = await graphFetch(`/sites/${siteId}/drives`);
  const drive = (drives.value as Array<{ id: string; name: string }>).find((d) => d.name === libraryName);
  if (!drive) {
    throw new Error(`Drive ${libraryName} not found`);
  }
  return drive.id;
}

export async function resolveItemId(siteId: string, driveId: string, drivePath: string): Promise<string> {
  if (!drivePath || drivePath === '/' || drivePath === '.') {
    return 'root';
  }
  const encoded = encodeURIComponent(drivePath.replace(/^\/+/, ''));
  const item = await graphFetch(`/sites/${siteId}/drives/${driveId}/root:/${encoded}`);
  if (!item?.id) {
    throw new Error(`Item not found for path ${drivePath}`);
  }
  return item.id;
}
