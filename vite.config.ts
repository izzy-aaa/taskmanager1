import { defineConfig } from 'vite';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);

function installed(name: string): boolean {
  try {
    require.resolve(name);
    return true;
  } catch {
    try {
      require.resolve(`${name}/package.json`);
      return true;
    } catch {
      return false;
    }
  }
}

// @canva-ct/genai is optional. When it isn't installed, point the import at a
// stub so the app still builds; meetings.ts then reports that AI isn't set up.
const genaiAlias = installed('@canva-ct/genai')
  ? {}
  : { '@canva-ct/genai': fileURLToPath(new URL('./src/genai-stub.ts', import.meta.url)) };

export default defineConfig({
  base: './',
  resolve: { alias: genaiAlias },
});
