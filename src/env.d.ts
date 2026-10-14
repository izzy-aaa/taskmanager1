/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_SHEETS_API_URL?: string;
}

// @canva-ct/genai is optional and private; it's used loosely through meetings.ts.
declare module '@canva-ct/genai';
