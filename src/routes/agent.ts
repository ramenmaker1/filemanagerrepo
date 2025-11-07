import { Router } from 'express';
import { z } from 'zod';

import { CONFIG } from '../config.js';
import { dispatchGraphAction } from '../msGraph/actions.js';
import { MsGraphActionSchema } from '../msGraph/schema.js';
import { openai } from '../openai.js';
import { logger } from '../logger.js';

const router = Router();

const ToolDefinition = {
  name: 'ms_graph_ops',
  description: 'Provision and manage SharePoint/OneDrive for Elion Studio.',
  parameters: {
    type: 'object',
    properties: {
      action: {
        type: 'string',
        enum: [
          'ensure_site',
          'ensure_libraries',
          'ensure_groups_permissions',
          'create_catalog_page',
          'share_deliverable',
          'list_folder',
          'link_repo_and_base44'
        ]
      },
      siteType: { type: 'string', enum: ['team', 'communication'] },
      siteName: { type: 'string' },
      libraryName: { type: 'string' },
      driveItemPath: { type: 'string' },
      shareType: { type: 'string', enum: ['view', 'edit'] },
      expiresAt: { type: 'string' },
      catalogLinks: {
        type: 'object',
        properties: {
          repos: { type: 'array', items: { type: 'string' } },
          base44: { type: 'array', items: { type: 'string' } },
          dataBuckets: { type: 'array', items: { type: 'string' } }
        }
      },
      repoUrl: { type: 'string' },
      base44Url: { type: 'string' },
      sharepointUrl: { type: 'string' }
    },
    required: ['action']
  }
} as const;

const ContentSchema = z.union([z.string(), z.array(z.any())]);

const MessageSchema = z.object({
  role: z.enum(['system', 'user', 'assistant', 'tool']),
  content: ContentSchema,
  tool_call_id: z.string().optional(),
  name: z.string().optional()
});

const AgentRequestSchema = z.object({
  messages: z.array(MessageSchema)
});

type Message = z.infer<typeof MessageSchema>;

function toResponseContent(content: Message['content']) {
  if (Array.isArray(content)) {
    return content;
  }
  return [{ type: 'text', text: content }];
}

function normalizeMessages(messages: Message[]) {
  return messages.map((message) => ({
    ...message,
    content: toResponseContent(message.content)
  }));
}

router.post('/complete', async (req, res, next) => {
  try {
    const { messages } = AgentRequestSchema.parse(req.body ?? {});
    const normalized = normalizeMessages(messages);

    const baseInput = [
      {
        role: 'system',
        content: [
          {
            type: 'text',
            text: 'You are the Elion SharePoint Agent. Keep the SharePoint structure in compliance: libraries (Projects, Assets, Data, Deliverables, Templates, Legal & Finance), expiring share links only from Deliverables, and a Catalog page listing repos, Base44 apps, and data buckets. When asked to share, create a time-boxed link. Never store .env or Git repos in cloud drives. If device settings are requested, return human steps (agent cannot flip OS switches). Use the tool to reconcile state idempotently.'
          }
        ]
      },
      ...normalized
    ];

    const first = await openai.responses.create({
      model: CONFIG.openai.model,
      input: baseInput,
      tools: [ToolDefinition]
    });

    const toolCall = first.output?.find(
      (item: any) => item.type === 'tool_call' && item.tool_name === 'ms_graph_ops'
    ) as any;

    if (!toolCall) {
      return res.json({ reply: first.output_text ?? '', response: first });
    }

    const args = JSON.parse(toolCall.arguments ?? '{}');
    const parsed = MsGraphActionSchema.parse(args);
    const tenant = req.tenant;
    if (!tenant) {
      throw new Error('Tenant context missing for agent request');
    }

    const result = await dispatchGraphAction(tenant, parsed);

    const second = await openai.responses.create({
      model: CONFIG.openai.model,
      input: [
        ...baseInput,
        {
          role: 'tool',
          tool_call_id: toolCall.id,
          content: [{ type: 'text', text: JSON.stringify(result) }]
        }
      ]
    });

    res.json({ reply: second.output_text ?? '', toolResult: result, response: second });
  } catch (error) {
    logger.error('Agent completion failed', { error });
    next(error);
  }
});

export default router;
