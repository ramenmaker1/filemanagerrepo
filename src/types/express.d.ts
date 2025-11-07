import type { RequestApiKeyContext } from '../middleware/auth.js';
import type { TenantRecord } from '../tenants.js';

declare global {
  namespace Express {
    interface Request {
      tenant?: TenantRecord;
      apiKey?: RequestApiKeyContext;
    }
  }
}

export {};
