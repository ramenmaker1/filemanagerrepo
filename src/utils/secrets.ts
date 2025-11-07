import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

export interface SecretSource {
  value?: string;
  env?: string;
  file?: string;
}

function stripTrailingNewline(value: string): string {
  return value.replace(/\r?\n$/, '');
}

export function readSecretFile(path: string): string {
  const absolute = resolve(path);
  const raw = readFileSync(absolute, 'utf8');
  return stripTrailingNewline(raw);
}

export function readSecretFromEnv(name: string): string | undefined {
  const direct = process.env[name];
  if (direct && direct.length > 0) {
    return direct;
  }

  const fileEnv = `${name}_FILE`;
  const filePath = process.env[fileEnv];
  if (filePath && filePath.length > 0) {
    return readSecretFile(filePath);
  }

  return undefined;
}

export function resolveSecretSource(label: string, source: SecretSource): string {
  const provided = [source.value, source.env, source.file].filter((item) => item !== undefined);

  if (provided.length === 0) {
    throw new Error(`${label} must define one of value, env, or file for secret material`);
  }

  if (provided.length > 1) {
    throw new Error(`${label} must not specify multiple secret sources simultaneously`);
  }

  if (source.value !== undefined) {
    return source.value;
  }

  if (source.env !== undefined) {
    const resolved = readSecretFromEnv(source.env);
    if (!resolved) {
      throw new Error(`Environment variable ${source.env} is not defined for ${label}`);
    }
    return resolved;
  }

  return readSecretFile(source.file as string);
}
