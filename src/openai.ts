import OpenAI from 'openai';

import { CONFIG } from './config.js';

export const openai = new OpenAI({ apiKey: CONFIG.openAiApiKey });
