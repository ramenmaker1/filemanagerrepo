import { Router } from 'express';
import { z } from 'zod';

import { dispatchGraphAction } from '../msGraph/actions.js';
import { logger } from '../logger.js';

const router = Router();

const CatalogSchema = z.object({
  title: z.string().min(1).default('Catalog'),
  repos: z.array(z.string().url()).default([]),
  base44: z.array(z.string().url()).default([]),
  dataBuckets: z.array(z.string().url()).default([]),
  displayName: z.string().min(1).optional(),
});

router.post('/ensure', async (req, res, next) => {
  try {
    const body = CatalogSchema.parse(req.body ?? {});
    const tenant = req.tenant;
    if (!tenant) {
      throw new Error('Tenant context missing for catalog ensure');
    }

    const result = await dispatchGraphAction(tenant, {
      action: 'create_catalog_page',
      catalogLinks: {
        repos: body.repos,
        base44: body.base44,
        dataBuckets: body.dataBuckets,
      },
      siteName: body.displayName,
    });
    res.json(result);
  } catch (error) {
    logger.error('Catalog ensure failed', { error });
    next(error);
  }
});

export default router;
