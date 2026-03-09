import * as fs from 'fs';
import * as path from 'path';

export function fileExists(filePath?: string): boolean {
  if (!filePath) {
    return false;
  }
  try {
    return fs.existsSync(filePath);
  } catch {
    return false;
  }
}

export function readJsonFile<T>(filePath: string, fallback: T): T {
  try {
    if (!fs.existsSync(filePath)) {
      return fallback;
    }
    const raw = fs.readFileSync(filePath, 'utf8');
    if (!raw.trim()) {
      return fallback;
    }
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

export function writeJsonFile(filePath: string, payload: unknown): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), 'utf8');
}
