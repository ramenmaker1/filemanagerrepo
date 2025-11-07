import { z } from 'zod';

export const MsGraphActionSchema = z.discriminatedUnion('action', [
  z.object({
    action: z.literal('ensure_site'),
    siteType: z.enum(['team', 'communication']).optional(),
    siteName: z.string().min(1).optional()
  }),
  z.object({
    action: z.literal('ensure_libraries')
  }),
  z.object({
    action: z.literal('ensure_groups_permissions')
  }),
  z.object({
    action: z.literal('create_catalog_page'),
    catalogLinks: z.object({
      repos: z.array(z.string()),
      base44: z.array(z.string()),
      dataBuckets: z.array(z.string())
    })
  }),
  z.object({
    action: z.literal('share_deliverable'),
    driveItemPath: z.string().min(1),
    shareType: z.enum(['view', 'edit']).optional(),
    expiresAt: z.string().optional()
  }),
  z.object({
    action: z.literal('list_folder'),
    libraryName: z.string().min(1),
    driveItemPath: z.string().default('/')
  }),
  z.object({
    action: z.literal('link_repo_and_base44'),
    repoUrl: z.string().url(),
    base44Url: z.string().url(),
    sharepointUrl: z.string().url()
  })
]);

export type MsGraphAction = z.infer<typeof MsGraphActionSchema>;
