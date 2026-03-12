import * as https from 'https';
import { URL } from 'url';
import * as zlib from 'zlib';

export interface OpenAiCompatibleOptions {
  apiKey: string;
  baseUrl: string;
  model: string;
  prompt: string;
  temperature: number;
  maxTokens: number;
  timeoutMs: number;
}

export interface OpenAiCompatibleResult {
  content: string;
  model?: string;
  usage?: {
    promptTokens?: number;
    completionTokens?: number;
    totalTokens?: number;
  };
}

export interface OpenAiCompatibleModelListOptions {
  apiKey: string;
  baseUrl: string;
  timeoutMs: number;
}

interface OpenAiLikeChoice {
  message?: {
    content?: unknown;
    reasoning_content?: unknown;
  };
  text?: unknown;
  delta?: {
    content?: unknown;
  };
}

interface OpenAiLikeResponse {
  model?: string;
  choices?: OpenAiLikeChoice[];
  status?: string | number;
  msg?: unknown;
  body?: unknown;
  output_text?: unknown;
  output?: unknown;
  data?: unknown;
  message?: {
    content?: unknown;
  };
  error?: {
    message?: string;
  };
  usage?: {
    prompt_tokens?: number;
    completion_tokens?: number;
    total_tokens?: number;
  };
}

const MODEL_NOT_SUPPORTED_MARKER = '[MODEL_NOT_SUPPORTED]';

function normalizeBusinessStatus(status: unknown): string {
  if (status === undefined || status === null) {
    return '';
  }
  return String(status).trim();
}

function isBusinessSuccessStatus(status: string): boolean {
  if (!status) {
    return true;
  }
  return status === '0' || status === '200' || /^2\d\d$/.test(status) || status.toLowerCase() === 'success';
}

function isModelNotSupportedMessage(message: string): boolean {
  const normalized = message.toLowerCase();
  return (
    normalized.includes('model not support')
    || normalized.includes('model not supported')
    || normalized.includes('模型不支持')
  );
}

export function isModelNotSupportedApiError(error: unknown): boolean {
  if (!(error instanceof Error)) {
    return false;
  }
  return error.message.includes(MODEL_NOT_SUPPORTED_MARKER) || isModelNotSupportedMessage(error.message);
}

function normalizeEndpoint(baseUrl: string): string {
  const trimmed = baseUrl.trim().replace(/\/+$/g, '');
  if (!trimmed) {
    throw new Error('API 地址为空');
  }
  if (trimmed.endsWith('/chat/completions')) {
    return trimmed;
  }
  return `${trimmed}/chat/completions`;
}

function normalizeModelsEndpoint(baseUrl: string): string {
  const trimmed = baseUrl.trim().replace(/\/+$/g, '');
  if (!trimmed) {
    throw new Error('API 地址为空');
  }
  if (trimmed.endsWith('/models')) {
    return trimmed;
  }
  if (trimmed.endsWith('/chat/completions')) {
    return `${trimmed.slice(0, -('/chat/completions'.length))}/models`;
  }
  return `${trimmed}/models`;
}

function collectTextParts(value: unknown, depth = 0): string[] {
  if (value == null || depth > 6) {
    return [];
  }

  if (typeof value === 'string') {
    const text = value.trim();
    return text ? [text] : [];
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    return [String(value)];
  }

  if (Array.isArray(value)) {
    return value.flatMap(item => collectTextParts(item, depth + 1));
  }

  if (typeof value !== 'object') {
    return [];
  }

  const objectValue = value as Record<string, unknown>;
  const preferredKeys = [
    'content',
    'text',
    'output_text',
    'reasoning_content',
    'result',
    'answer',
    'message',
    'delta'
  ];
  const preferredParts = preferredKeys.flatMap(key => collectTextParts(objectValue[key], depth + 1));
  if (preferredParts.length > 0) {
    return preferredParts;
  }

  return Object.keys(objectValue).flatMap(key => {
    if (key.includes('text') || key.includes('content')) {
      return collectTextParts(objectValue[key], depth + 1);
    }
    return [];
  });
}

function normalizeMessageContent(content: unknown): string {
  const parts = collectTextParts(content)
    .map(item => item.trim())
    .filter(Boolean);
  if (parts.length === 0) {
    return '';
  }
  return parts.join('\n').trim();
}

function parseResponseBody(body: string): OpenAiCompatibleResult {
  let parsed: OpenAiLikeResponse;
  try {
    parsed = JSON.parse(body) as OpenAiLikeResponse;
  } catch {
    throw new Error(`API 返回非 JSON：${body.slice(0, 200)}`);
  }

  const businessStatus = normalizeBusinessStatus(parsed.status);
  const businessMessage = normalizeMessageContent(parsed.msg);
  if (!isBusinessSuccessStatus(businessStatus)) {
    const detail = businessMessage || '未知错误';
    if (businessStatus === '435' || isModelNotSupportedMessage(detail)) {
      throw new Error(`${MODEL_NOT_SUPPORTED_MARKER} API 模型不支持（status=${businessStatus}）：${detail}`);
    }
    throw new Error(`API 业务错误（status=${businessStatus}）：${detail}`);
  }

  if (parsed.error?.message) {
    throw new Error(`API 返回错误：${parsed.error.message}`);
  }

  const choice = parsed.choices?.[0];
  const candidates: unknown[] = [
    choice?.message?.content,
    choice?.message?.reasoning_content,
    choice?.delta?.content,
    choice?.text,
    parsed.output_text,
    parsed.message?.content,
    parsed.body,
    parsed.output,
    parsed.data
  ];
  let content = '';
  for (const candidate of candidates) {
    content = normalizeMessageContent(candidate);
    if (content) {
      break;
    }
  }

  if (!content && choice) {
    content = normalizeMessageContent(choice);
  }

  if (!content) {
    throw new Error(`API 返回为空，未获取到可写入内容。响应片段：${body.slice(0, 200)}`);
  }

  return {
    content,
    model: parsed.model,
    usage: parsed.usage
      ? {
        promptTokens: parsed.usage.prompt_tokens,
        completionTokens: parsed.usage.completion_tokens,
        totalTokens: parsed.usage.total_tokens
      }
      : undefined
  };
}

function isOpenRouterEndpoint(url: URL): boolean {
  return /(^|\.)openrouter\.ai$/i.test(url.hostname);
}

function decodeResponseBody(payload: Buffer, contentEncoding: string | undefined): string {
  const normalized = (contentEncoding || '').toLowerCase();
  try {
    if (normalized.includes('gzip')) {
      return zlib.gunzipSync(payload).toString('utf8');
    }
    if (normalized.includes('br')) {
      return zlib.brotliDecompressSync(payload).toString('utf8');
    }
    if (normalized.includes('deflate')) {
      return zlib.inflateSync(payload).toString('utf8');
    }
  } catch {
    // 解压失败时回退到 utf8 直接解析，避免丢失原始错误信息
  }
  return payload.toString('utf8');
}

function extractModelName(value: unknown): string | null {
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed || null;
  }
  if (!value || typeof value !== 'object') {
    return null;
  }
  const objectValue = value as Record<string, unknown>;
  const fromId = typeof objectValue.id === 'string' ? objectValue.id.trim() : '';
  if (fromId) {
    return fromId;
  }
  const fromModel = typeof objectValue.model === 'string' ? objectValue.model.trim() : '';
  if (fromModel) {
    return fromModel;
  }
  const fromName = typeof objectValue.name === 'string' ? objectValue.name.trim() : '';
  if (fromName) {
    return fromName;
  }
  return null;
}

function collectModelNames(value: unknown, depth = 0): string[] {
  if (value == null || depth > 6) {
    return [];
  }
  if (Array.isArray(value)) {
    return value.flatMap(item => {
      const name = extractModelName(item);
      if (name) {
        return [name];
      }
      return collectModelNames(item, depth + 1);
    });
  }
  if (typeof value !== 'object') {
    return [];
  }

  const objectValue = value as Record<string, unknown>;
  const direct = extractModelName(objectValue);
  if (direct) {
    return [direct];
  }
  const preferredKeys = ['data', 'models', 'items', 'result', 'results', 'list'];
  const preferred = preferredKeys.flatMap(key => collectModelNames(objectValue[key], depth + 1));
  if (preferred.length > 0) {
    return preferred;
  }
  return Object.values(objectValue).flatMap(item => collectModelNames(item, depth + 1));
}

function uniqueModelNames(models: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const model of models) {
    const trimmed = model.trim();
    if (!trimmed || seen.has(trimmed)) {
      continue;
    }
    seen.add(trimmed);
    result.push(trimmed);
  }
  return result;
}

export async function generateWithOpenAiCompatibleApi(
  options: OpenAiCompatibleOptions
): Promise<OpenAiCompatibleResult> {
  const endpoint = normalizeEndpoint(options.baseUrl);
  const url = new URL(endpoint);
  const body = JSON.stringify({
    model: options.model,
    messages: [
      {
        role: 'user',
        content: options.prompt
      }
    ],
    temperature: options.temperature,
    max_tokens: options.maxTokens,
    stream: false
  });

  return new Promise<OpenAiCompatibleResult>((resolve, reject) => {
    const headers: Record<string, string | number> = {
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(body),
      Authorization: `Bearer ${options.apiKey}`,
      Accept: 'application/json',
      'Accept-Encoding': 'identity',
      'User-Agent': 'writing-agent/0.0.2'
    };
    if (isOpenRouterEndpoint(url)) {
      headers['HTTP-Referer'] = process.env.WRITING_AGENT_OPENROUTER_REFERER || 'https://github.com/daydayup2026/writing-agent';
      headers['X-Title'] = process.env.WRITING_AGENT_OPENROUTER_TITLE || 'writing-agent';
    }

    const request = https.request(
      {
        protocol: url.protocol,
        hostname: url.hostname,
        port: url.port || undefined,
        path: `${url.pathname}${url.search}`,
        method: 'POST',
        headers,
        timeout: options.timeoutMs
      },
      response => {
        const chunks: Buffer[] = [];
        response.on('data', chunk => {
          if (typeof chunk === 'string') {
            chunks.push(Buffer.from(chunk, 'utf8'));
            return;
          }
          chunks.push(chunk);
        });

        response.on('end', () => {
          const statusCode = response.statusCode || 0;
          const payload = decodeResponseBody(Buffer.concat(chunks), response.headers['content-encoding']);
          if (statusCode >= 400) {
            reject(new Error(`API 请求失败（HTTP ${statusCode}）：${payload.slice(0, 300)}`));
            return;
          }

          try {
            resolve(parseResponseBody(payload));
          } catch (error) {
            reject(error);
          }
        });
      }
    );

    request.on('timeout', () => {
      request.destroy(new Error(`API 请求超时（>${options.timeoutMs}ms）`));
    });

    request.on('error', error => {
      reject(error);
    });

    request.write(body);
    request.end();
  });
}

export async function listOpenAiCompatibleModels(
  options: OpenAiCompatibleModelListOptions
): Promise<string[]> {
  const endpoint = normalizeModelsEndpoint(options.baseUrl);
  const url = new URL(endpoint);

  return new Promise<string[]>((resolve, reject) => {
    const headers: Record<string, string | number> = {
      Authorization: `Bearer ${options.apiKey}`,
      Accept: 'application/json',
      'Accept-Encoding': 'identity',
      'User-Agent': 'writing-agent/0.0.2'
    };
    if (isOpenRouterEndpoint(url)) {
      headers['HTTP-Referer'] = process.env.WRITING_AGENT_OPENROUTER_REFERER || 'https://github.com/daydayup2026/writing-agent';
      headers['X-Title'] = process.env.WRITING_AGENT_OPENROUTER_TITLE || 'writing-agent';
    }

    const request = https.request(
      {
        protocol: url.protocol,
        hostname: url.hostname,
        port: url.port || undefined,
        path: `${url.pathname}${url.search}`,
        method: 'GET',
        headers,
        timeout: options.timeoutMs
      },
      response => {
        const chunks: Buffer[] = [];
        response.on('data', chunk => {
          if (typeof chunk === 'string') {
            chunks.push(Buffer.from(chunk, 'utf8'));
            return;
          }
          chunks.push(chunk);
        });

        response.on('end', () => {
          const statusCode = response.statusCode || 0;
          const payload = decodeResponseBody(Buffer.concat(chunks), response.headers['content-encoding']);
          if (statusCode >= 400) {
            reject(new Error(`模型列表请求失败（HTTP ${statusCode}）：${payload.slice(0, 300)}`));
            return;
          }

          let parsed: Record<string, unknown>;
          try {
            parsed = JSON.parse(payload) as Record<string, unknown>;
          } catch {
            reject(new Error(`模型列表返回非 JSON：${payload.slice(0, 200)}`));
            return;
          }

          const errorNode = parsed.error as { message?: unknown } | undefined;
          const errorMessage = typeof errorNode?.message === 'string' ? errorNode.message.trim() : '';
          if (errorMessage) {
            reject(new Error(`模型列表请求失败：${errorMessage}`));
            return;
          }

          const models = uniqueModelNames(
            collectModelNames(parsed.data).concat(collectModelNames(parsed.models))
          );
          resolve(models);
        });
      }
    );

    request.on('timeout', () => {
      request.destroy(new Error(`模型列表请求超时（>${options.timeoutMs}ms）`));
    });

    request.on('error', error => {
      reject(error);
    });

    request.end();
  });
}
