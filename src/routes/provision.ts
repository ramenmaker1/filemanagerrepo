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
    const result = await dispatchGraphAction(body.displayName ?? CONFIG.siteDisplayName, {
      action: 'ensure_site',
      siteType: body.siteType ?? CONFIG.siteType,
      siteName: body.displayName ?? CONFIG.siteDisplayName
    });
    res.json(result);
  } catch (error) {
    logger.error('Provision failed', { error });
    next(error);
  }
});

export default router;
