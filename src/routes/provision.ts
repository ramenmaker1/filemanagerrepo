import { Router } from 'express';
import { z } from 'zod';

import { CONFIG } from '../config.js';
import { dispatchGraphAction } from '../msGraph/actions.js';
import { logger } from '../logger.js';

const router = Router();

const ProvisionSchema = z.object({
  siteType: z.enum(['team', 'communication']).optional(),
  displayName: z.string().min(1).optional()
});

router.post('/', async (req, res, next) => {
  try {
    const body = ProvisionSchema.parse(req.body ?? {});
    const displayName = body.displayName ?? CONFIG.sharepoint.siteDisplayName;
    const siteType = body.siteType ?? CONFIG.sharepoint.siteType;
    const result = await dispatchGraphAction(displayName, {
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
