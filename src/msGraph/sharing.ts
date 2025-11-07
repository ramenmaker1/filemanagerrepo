import { graphFetch } from './auth.js';

export interface ShareLinkOptions {
  driveId: string;
  itemId: string;
  type?: 'view' | 'edit';
  expiresAt?: string;
}

export async function createExpiringLink({ driveId, itemId, type = 'view', expiresAt }: ShareLinkOptions): Promise<string> {
  const body: Record<string, unknown> = {
    type,
    scope: 'anonymous'
  };

  if (expiresAt) {
    body.expirationDateTime = expiresAt;
  }

  const response = await graphFetch(`/drives/${driveId}/items/${itemId}/createLink`, {
    method: 'POST',
    body: JSON.stringify(body)
  });

  return response?.link?.webUrl as string;
}
