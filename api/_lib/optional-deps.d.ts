/**
 * Ambient declarations for the OPTIONAL storage SDKs.
 *
 * `@vercel/blob` and `@upstash/redis` are provisioned by the PO at deploy time
 * (they are NOT in package.json for the build/local-test phase — this wave never
 * touches live Vercel storage). These minimal declarations let the real storage
 * adapter (api/_lib/storage.ts) `import()` them and type-check WITHOUT the
 * packages installed. On Vercel the real packages supply the full types; these
 * stubs are shadowed there and only cover the members we call.
 */
declare module '@vercel/blob' {
  export function put(
    path: string,
    body: Uint8Array | ArrayBuffer | Blob | string,
    opts: {
      access: 'public';
      contentType?: string;
      cacheControlMaxAge?: number;
      addRandomSuffix?: boolean;
      token?: string;
    },
  ): Promise<{ url: string; pathname: string }>;
  export function del(url: string | string[], opts?: { token?: string }): Promise<void>;
}

declare module '@upstash/redis' {
  export class Redis {
    static fromEnv(): Redis;
    get<T>(key: string): Promise<T | null>;
    set(key: string, value: unknown, opts?: { ex?: number }): Promise<unknown>;
    del(...keys: string[]): Promise<number>;
    incr(key: string): Promise<number>;
    expire(key: string, seconds: number): Promise<unknown>;
    sadd(key: string, ...members: string[]): Promise<number>;
    srem(key: string, ...members: string[]): Promise<number>;
    scard(key: string): Promise<number>;
    smembers(key: string): Promise<string[]>;
  }
}
