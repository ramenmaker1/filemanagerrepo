import { Router } from 'express';
import { z } from 'zod';

import { dispatchGraphAction } from '../msGraph/actions.js';
import { logger } from '../logger.js';

const router = Router();

const ShareSchema = z.object({
  path: z.string().min(1),
  type: z.enum(['view', 'edit']).default('view'),
  expiresAt: z.string().datetime().optional(),
  displayName: z.string().min(1).optional(),
});

router.post('/', async (req, res, next) => {
  try {
    const body = ShareSchema.parse(req.body ?? {});
    const tenant = req.tenant;
    if (!tenant) {
      throw new Error('Tenant context missing for share request');
    }

    const result = await dispatchGraphAction(tenant, {
      action: 'share_deliverable',
      driveItemPath: body.path,
      shareType: body.type,
      expiresAt: body.expiresAt,
      siteName: body.displayName,
    });
    res.json(result);
  } catch (error) {
    logger.error('Share link creation failed', { error });
    next(error);
  }
});

export default router;
