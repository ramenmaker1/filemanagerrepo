import { Router } from 'express';
import { z } from 'zod';

import { CONFIG } from '../config.js';
import { dispatchGraphAction } from '../msGraph/actions.js';
import { logger } from '../logger.js';

const router = Router();

const ListSchema = z.object({
  libraryName: z.string().min(1),
  path: z.string().default('/'),
  displayName: z.string().min(1).optional()
});

router.get('/', async (req, res, next) => {
  try {
    const body = ListSchema.parse({
      libraryName: req.query.libraryName,
      path: req.query.path ?? '/',
      displayName: req.query.displayName
    });

    const displayName = body.displayName ?? CONFIG.sharepoint.siteDisplayName;
    const result = await dispatchGraphAction(displayName, {
      action: 'list_folder',
      libraryName: body.libraryName,
      driveItemPath: body.path
    });
    res.json(result);
  } catch (error) {
    logger.error('List folder failed', { error });
    next(error);
  }
});

export default router;
