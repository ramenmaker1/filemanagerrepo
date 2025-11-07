import { graphFetch } from './auth.js';
import { logger } from '../logger.js';

const REQUIRED_LIBRARIES = ['Projects', 'Assets', 'Data', 'Deliverables', 'Templates', 'Legal & Finance'];

export async function ensureDocumentLibrary(siteId: string, name: string): Promise<void> {
  const lists = await graphFetch(`/sites/${siteId}/lists?$filter=displayName eq '${name.replace(/'/g, "''")}'`);
  if (lists.value?.length) {
    logger.info('Library already exists', { siteId, name });
    return;
  }

  await graphFetch(`/sites/${siteId}/lists`, {
    method: 'POST',
    body: JSON.stringify({
      displayName: name,
      list: { template: 'documentLibrary' }
    })
  });

  logger.info('Created document library', { siteId, name });
}

export async function ensureLibraries(siteId: string): Promise<void> {
  for (const name of REQUIRED_LIBRARIES) {
    await ensureDocumentLibrary(siteId, name);
  }
}

export { REQUIRED_LIBRARIES };
