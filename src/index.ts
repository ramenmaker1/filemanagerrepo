import express, { type NextFunction, type Request, type Response } from 'express';

import { CONFIG, requireEnv } from './config.js';
import { logger } from './logger.js';
import provision from './routes/provision.js';
import share from './routes/share.js';
import catalog from './routes/catalog.js';
import list from './routes/list.js';
import agent from './routes/agent.js';

requireEnv();

const app = express();
app.use(express.json());

app.get('/health', (_req, res) => res.json({ ok: true }));
app.use('/provision', provision);
app.use('/share', share);
app.use('/catalog', catalog);
app.use('/list', list);
app.use('/agent', agent);

app.use((err: any, _req: Request, res: Response, _next: NextFunction) => {
  logger.error('Unhandled error', { error: err.message });
  res.status(500).json({ error: err.message ?? 'Internal Server Error' });
});

app.listen(CONFIG.port, () => {
  logger.info(`Elion SP Agent listening on :${CONFIG.port}`);
});
