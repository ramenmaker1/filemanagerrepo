import { Router } from 'express';
import { z } from 'zod';

import { dispatchGraphAction } from '../msGraph/actions.js';
import { logger } from '../logger.js';

const router = Router();

const ListSchema = z.object({
  libraryName: z.string().min(1),
  path: z.string().default('/'),
  displayName: z.string().min(1).optional(),
});

router.get('/', async (req, res, next) => {
  try {
    const body = ListSchema.parse({
      libraryName: req.query.libraryName,
      path: req.query.path ?? '/',
      displayName: req.query.displayName,
    });

    const tenant = req.tenant;
    if (!tenant) {
      throw new Error('Tenant context missing for list request');
    }

    const result = await dispatchGraphAction(tenant, {
      action: 'list_folder',
      libraryName: body.libraryName,
      driveItemPath: body.path,
      siteName: body.displayName,
    });
    res.json(result);
  } catch (error) {
    logger.error('List folder failed', { error });
    next(error);
  }
});

export default router;
