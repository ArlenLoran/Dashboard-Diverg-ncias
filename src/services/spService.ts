/**
 * SharePoint REST API helper (classic _spPageContextInfo)
 * - Same-origin fetch
 * - Uses request digest for write operations
 * - OData verbose
 */

declare global {
  interface Window {
    _spPageContextInfo?: any;
  }
}

export type SpResult<T> =
  | { status: true; data: T; message?: string }
  | { status: false; message: string; error?: any };

function sanitizeInput(input: string): string {
  // Remove < and > to avoid basic injection in list/folder names
  return String(input ?? '').replace(/[<>]/g, '').trim();
}

function getContext() {
  const ctx = window._spPageContextInfo;
  if (!ctx) {
    throw new Error(
      'SharePoint context (_spPageContextInfo) não encontrado. Este app deve rodar dentro de uma página SharePoint.'
    );
  }
  return ctx;
}

export function hasSpContext(): boolean {
  return !!window._spPageContextInfo;
}

export function normalizeSharePointEmail(value: unknown): string {
  const raw = String(value ?? '').trim().toLowerCase();
  if (!raw) return '';

  const lastPipeValue = raw.includes('|') ? (raw.split('|').pop() || raw) : raw;
  const match = lastPipeValue.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
  return String(match?.[0] || lastPipeValue).trim().toLowerCase();
}

export function getCurrentSharePointUserEmail(): string {
  const ctx = window._spPageContextInfo;
  if (!ctx) return '';

  const candidates = [
    ctx.userEmail,
    ctx.userLoginName,
    ctx.userPrincipalName,
    ctx.accountName,
    ctx.systemUserKey,
    ctx.userDisplayName
  ];

  for (const candidate of candidates) {
    const email = normalizeSharePointEmail(candidate);
    if (email && email.includes('@')) return email;
  }

  return '';
}

export function isSharePointUserAllowedByEnv(userEmail: string, allowedEmailsEnv: string): boolean {
  const normalizedUserEmail = normalizeSharePointEmail(userEmail);
  if (!normalizedUserEmail) return false;

  const allowedEmails = String(allowedEmailsEnv || '')
    .split(/[\s,;]+/)
    .map(normalizeSharePointEmail)
    .filter(Boolean);

  return allowedEmails.includes(normalizedUserEmail);
}

export function spSiteUrl(): string {
  return getContext().siteAbsoluteUrl;
}

export function spWebRelUrl(): string {
  return getContext().webServerRelativeUrl;
}

function digest(): string {
  return getContext().formDigestValue;
}

// ============================================================================
// Request Digest (FormDigestValue) refresh
// ============================================================================

let _digestCacheValue: string | null = null;
let _digestCacheExpiresAt = 0; // epoch ms

function isDigestValid(): boolean {
  return !!_digestCacheValue && Date.now() < _digestCacheExpiresAt;
}

async function refreshDigest(force: boolean = false): Promise<string> {
  if (!force && isDigestValid()) return _digestCacheValue as string;

  const ctx = getContext();
  if (!force && !_digestCacheValue && ctx?.formDigestValue) {
    _digestCacheValue = String(ctx.formDigestValue);
    _digestCacheExpiresAt = Date.now() + 4 * 60 * 1000;
    return _digestCacheValue;
  }

  const url = `${spSiteUrl()}/_api/contextinfo`;
  const resp = await fetch(url, {
    method: 'POST',
    headers: {
      Accept: 'application/json; odata=verbose',
      'Content-Type': 'application/json; odata=verbose'
    },
    credentials: 'same-origin'
  });
  if (!resp.ok) throw new Error(await parseSpError(resp));
  const data = await resp.json();
  const info = data?.d?.GetContextWebInformation;
  const value = String(info?.FormDigestValue || '').trim();
  const timeoutSec = Number(info?.FormDigestTimeoutSeconds || 0);
  if (!value) throw new Error('Não foi possível obter FormDigestValue (contextinfo).');

  _digestCacheValue = value;
  const safetyMs = 30 * 1000;
  const ttlMs = timeoutSec > 0 ? timeoutSec * 1000 : 4 * 60 * 1000;
  _digestCacheExpiresAt = Date.now() + Math.max(30 * 1000, ttlMs - safetyMs);

  try {
    (window as any)._spPageContextInfo.formDigestValue = value;
  } catch {
    // ignore
  }

  return value;
}

function looksLikeDigestInvalidMessage(msg: string): boolean {
  const m = String(msg || '').toLowerCase();
  return (
    m.includes('security validation for this page is invalid') ||
    m.includes('the security validation for this page is invalid') ||
    m.includes('x-requestdigest') ||
    m.includes('formdigest')
  );
}

async function fetchWithDigestRetry(url: string, init: RequestInit, retryOnce: boolean = true): Promise<Response> {
  const d = await refreshDigest(false);
  const headers: any = { ...(init.headers || {}) };
  headers['X-RequestDigest'] = d;

  const resp = await fetch(url, { ...init, headers, credentials: 'same-origin' });
  if (resp.ok) return resp;

  const msg = await parseSpError(resp);
  if (retryOnce && looksLikeDigestInvalidMessage(msg)) {
    const d2 = await refreshDigest(true);
    const headers2: any = { ...(init.headers || {}) };
    headers2['X-RequestDigest'] = d2;
    return fetch(url, { ...init, headers: headers2, credentials: 'same-origin' });
  }

  return resp;
}

function headersJson(): HeadersInit {
  return {
    Accept: 'application/json; odata=verbose',
    'Content-Type': 'application/json; odata=verbose'
  };
}

async function parseSpError(resp: Response): Promise<string> {
  try {
    const data = await resp.json();
    return (
      data?.error?.message?.value ||
      data?.odata?.error?.message?.value ||
      resp.statusText ||
      'Erro SharePoint'
    );
  } catch {
    return resp.statusText || 'Erro SharePoint';
  }
}

function escapeODataString(str: string): string {
  return String(str ?? '').replace(/'/g, "''");
}

const META_CACHE_TTL_MS = 10 * 60 * 1000; 
type CacheEntry<T> = { value: T; expiresAt: number };
const entityTypeCache = new Map<string, CacheEntry<string>>();

async function listItemEntityTypeFullName(listName: string): Promise<string> {
  const ln = sanitizeInput(listName);
  if (!ln) throw new Error('listName inválido.');
  const key = ln.toLowerCase();
  const cached = entityTypeCache.get(key);
  if (cached && Date.now() < cached.expiresAt) return cached.value;

  const url = `${spSiteUrl()}/_api/web/lists/getByTitle('${escapeODataString(ln)}')?$select=ListItemEntityTypeFullName`;
  const resp = await fetch(url, { method: 'GET', headers: { Accept: 'application/json; odata=verbose' }, credentials: 'same-origin' });
  if (!resp.ok) throw new Error(await parseSpError(resp));
  const data = await resp.json();
  const t = data?.d?.ListItemEntityTypeFullName;
  if (!t) throw new Error('Não foi possível obter ListItemEntityTypeFullName.');
  entityTypeCache.set(key, { value: t, expiresAt: Date.now() + META_CACHE_TTL_MS });
  return t;
}

export async function spListExists(listName: string): Promise<boolean> {
  const ln = sanitizeInput(listName);
  if (!ln) return false;
  const url = `${spSiteUrl()}/_api/web/lists?$select=Id,Title&$filter=Title%20eq%20'${escapeODataString(ln)}'&$top=1`;
  try {
    const resp = await fetch(url, {
      method: 'GET',
      headers: { Accept: 'application/json; odata=verbose' },
      credentials: 'same-origin'
    });
    if (!resp.ok) return false;
    const data: any = await resp.json();
    const results = data?.d?.results;
    return Array.isArray(results) && results.length > 0;
  } catch {
    return false;
  }
}

export async function spCreateList(listName: string, template: number = 100): Promise<SpResult<true>> {
  try {
    const url = `${spSiteUrl()}/_api/web/lists`;
    const body = {
      __metadata: { type: 'SP.List' },
      AllowContentTypes: true,
      BaseTemplate: template,
      ContentTypesEnabled: true,
      Description: 'Lista de configuração do App',
      Title: sanitizeInput(listName)
    };
    const resp = await fetchWithDigestRetry(url, {
      method: 'POST',
      headers: headersJson(),
      body: JSON.stringify(body)
    });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    return { status: true, data: true };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}

export async function spListGetItems<T = any>(
  listName: string,
  options?: {
    select?: string[] | string;
    filter?: string;
    orderBy?: string;
    top?: number;
  }
): Promise<SpResult<T[]>> {
  const ln = sanitizeInput(listName);
  if (!ln) return { status: false, message: 'listName inválido.' };

  const selectArr: string[] = Array.isArray(options?.select)
    ? (options!.select as string[])
    : typeof options?.select === 'string'
      ? options.select.split(',').map((s) => s.trim()).filter(Boolean)
      : [];

  const select = selectArr.length ? `$select=${selectArr.join(',')}` : '';
  const filter = options?.filter ? `$filter=${options.filter}` : '';
  const orderBy = options?.orderBy ? `$orderby=${options.orderBy}` : '';
  const top = options?.top ? `$top=${options.top}` : '$top=5000';

  const qs = [select, filter, orderBy, top].filter(Boolean).join('&');
  const url = `${spSiteUrl()}/_api/web/lists/getByTitle('${escapeODataString(ln)}')/items${qs ? `?${qs}` : ''}`;

  try {
    const resp = await fetch(url, {
      method: 'GET',
      headers: { Accept: 'application/json; odata=verbose' },
      credentials: 'same-origin'
    });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    const data = await resp.json();
    return { status: true, data: data?.d?.results || [] };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}

export async function spListAddItem<T = any>(
  listName: string,
  fields: Record<string, any>
): Promise<SpResult<{ id: number; raw: T }>> {
  const ln = sanitizeInput(listName);
  try {
    const entityType = await listItemEntityTypeFullName(ln);
    const body: any = { ...fields, __metadata: { type: entityType } };
    const url = `${spSiteUrl()}/_api/web/lists/getByTitle('${escapeODataString(ln)}')/items`;
    const resp = await fetchWithDigestRetry(url, {
      method: 'POST',
      headers: headersJson(),
      body: JSON.stringify(body)
    });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    const data = await resp.json();
    return { status: true, data: { id: data?.d?.Id, raw: data?.d } };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}

export async function spListUpdateItem(
  listName: string,
  itemId: number,
  fields: Record<string, any>
): Promise<SpResult<true>> {
  const ln = sanitizeInput(listName);
  try {
    const entityType = await listItemEntityTypeFullName(ln);
    const body: any = { ...fields, __metadata: { type: entityType } };
    const url = `${spSiteUrl()}/_api/web/lists/getByTitle('${escapeODataString(ln)}')/items(${itemId})`;
    const resp = await fetchWithDigestRetry(url, {
      method: 'POST',
      headers: { ...headersJson(), 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
      body: JSON.stringify(body)
    });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    return { status: true, data: true };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}

export async function spListDeleteItem(
  listName: string,
  itemId: number
): Promise<SpResult<true>> {
  const ln = sanitizeInput(listName);
  try {
    const url = `${spSiteUrl()}/_api/web/lists/getByTitle('${escapeODataString(ln)}')/items(${itemId})`;
    const resp = await fetchWithDigestRetry(url, {
      method: 'POST',
      headers: { ...headersJson(), 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' }
    });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    return { status: true, data: true };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}

async function spFieldExists(listName: string, fieldTitle: string): Promise<boolean> {
  const ln = sanitizeInput(listName);
  const ft = sanitizeInput(fieldTitle);
  const url = `${spSiteUrl()}/_api/web/lists/getByTitle('${escapeODataString(ln)}')/fields/getByInternalNameOrTitle('${escapeODataString(ft)}')`;
  const resp = await fetch(url, { method: 'GET', headers: { Accept: 'application/json; odata=verbose' }, credentials: 'same-origin' });
  return resp.ok;
}

export async function spListEnsureTextField(listName: string, fieldTitle: string): Promise<SpResult<true>> {
  try {
    if (await spFieldExists(listName, fieldTitle)) return { status: true, data: true };
    const ln = sanitizeInput(listName);
    const ft = sanitizeInput(fieldTitle);
    const url = `${spSiteUrl()}/_api/web/lists/getByTitle('${escapeODataString(ln)}')/fields`;
    const body = { __metadata: { type: 'SP.Field' }, Title: ft, FieldTypeKind: 2 };
    const resp = await fetchWithDigestRetry(url, { method: 'POST', headers: headersJson(), body: JSON.stringify(body) });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    return { status: true, data: true };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}

export async function spListEnsureMultiLineTextField(listName: string, fieldTitle: string): Promise<SpResult<true>> {
  try {
    if (await spFieldExists(listName, fieldTitle)) return { status: true, data: true };
    const ln = sanitizeInput(listName);
    const ft = sanitizeInput(fieldTitle);
    const listUrl = `${spSiteUrl()}/_api/web/lists/getByTitle('${escapeODataString(ln)}')`;
    const safeName = ft.replace(/\s+/g, '');
    const schemaXml = `<Field DisplayName="${ft}" Name="${safeName}" StaticName="${safeName}" Type="Note" RichText="false" />`;
    const url = `${listUrl}/fields/createFieldAsXml`;
    const body = {
      parameters: {
        __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
        SchemaXml: schemaXml,
        Options: 0
      }
    };
    const resp = await fetchWithDigestRetry(url, { method: 'POST', headers: headersJson(), body: JSON.stringify(body) });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    return { status: true, data: true };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}

export async function spListEnsureNumberField(listName: string, fieldTitle: string): Promise<SpResult<true>> {
  try {
    if (await spFieldExists(listName, fieldTitle)) return { status: true, data: true };
    const ln = sanitizeInput(listName);
    const ft = sanitizeInput(fieldTitle);
    const url = `${spSiteUrl()}/_api/web/lists/getByTitle('${escapeODataString(ln)}')/fields`;
    const body = { __metadata: { type: 'SP.Field' }, Title: ft, FieldTypeKind: 9 };
    const resp = await fetchWithDigestRetry(url, { method: 'POST', headers: headersJson(), body: JSON.stringify(body) });
    if (!resp.ok) return { status: false, message: await parseSpError(resp) };
    return { status: true, data: true };
  } catch (error: any) {
    return { status: false, message: error.message };
  }
}
