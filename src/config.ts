import 'dotenv/config';

const requiredVars = [
  'AZURE_TENANT_ID',
  'AZURE_CLIENT_ID',
  'AZURE_CLIENT_SECRET',
  'OPENAI_API_KEY'
] as const;

export function requireEnv(): void {
  const missing = requiredVars.filter((key) => !process.env[key]);
  if (missing.length) {
    throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
  }
}

export const CONFIG = {
  tenantId: process.env.AZURE_TENANT_ID ?? '',
  clientId: process.env.AZURE_CLIENT_ID ?? '',
  clientSecret: process.env.AZURE_CLIENT_SECRET ?? '',
  openAiApiKey: process.env.OPENAI_API_KEY ?? '',
  siteDisplayName: process.env.SITE_DISPLAY_NAME ?? 'Elion Studio',
  port: Number(process.env.PORT ?? 8080),
  openaiModel: process.env.OPENAI_MODEL ?? 'gpt-4.1'
};
