import { Router } from 'express';
import { z } from 'zod';

import { dispatchGraphAction } from '../msGraph/actions.js';
import { logger } from '../logger.js';

const router = Router();

const ProvisionSchema = z.object({
  siteType: z.enum(['team', 'communication']).optional(),
  displayName: z.string().min(1).optional(),
});

router.post('/', async (req, res, next) => {
  try {
    const body = ProvisionSchema.parse(req.body ?? {});
    const tenant = req.tenant;
    if (!tenant) {
      throw new Error('Tenant context missing for provisioning');
    }

    const displayName = body.displayName ?? tenant.sharepoint.siteDisplayName;
    const siteType = body.siteType ?? tenant.sharepoint.siteType ?? 'team';
    const result = await dispatchGraphAction(tenant, {
      action: 'ensure_site',
      siteType,
      siteName: displayName,
    });
    res.json(result);
  } catch (error) {
    logger.error('Provision failed', { error });
    next(error);
  }
});

export default router;
