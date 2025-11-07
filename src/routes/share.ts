import { Router } from 'express';
import { z } from 'zod';

import { CONFIG } from '../config.js';
import { dispatchGraphAction } from '../msGraph/actions.js';
import { logger } from '../logger.js';

const router = Router();

const ShareSchema = z.object({
  path: z.string().min(1),
  type: z.enum(['view', 'edit']).default('view'),
  expiresAt: z.string().datetime().optional(),
  displayName: z.string().min(1).optional()
});

router.post('/', async (req, res, next) => {
  try {
    const body = ShareSchema.parse(req.body ?? {});
    const displayName = body.displayName ?? CONFIG.sharepoint.siteDisplayName;
    const result = await dispatchGraphAction(displayName, {
      action: 'share_deliverable',
      driveItemPath: body.path,
      shareType: body.type,
      expiresAt: body.expiresAt
    });
    res.json(result);
  } catch (error) {
    logger.error('Share link creation failed', { error });
    next(error);
  }
});

export default router;
