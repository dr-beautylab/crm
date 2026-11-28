// Bundles @anthropic-ai/sdk into one file (src/vendor/anthropic-sdk.js).
// Cloudflare Pages here has no build step, so packages are not installed there.
// Node built-ins are only used by credential-file / local-tool helpers we never call
// (we always pass apiKey), so they are replaced with empty stand-ins.
// Re-run after upgrading the SDK: node tools/vendor-sdk.mjs
import { build } from 'esbuild';
const stub = { name: 'no-node', setup(b) {
  b.onResolve({ filter: /^node:/ }, (a) => ({ path: a.path, namespace: 'no-node' }));
  b.onLoad({ filter: /.*/, namespace: 'no-node' }, () => ({ contents: 'export default {}; export const once = undefined, Worker = undefined;' }));
} };
await build({
  stdin: { contents: "export { default } from '@anthropic-ai/sdk';", resolveDir: process.cwd() },
  bundle: true, format: 'esm', platform: 'neutral', mainFields: ['module', 'main'],
  conditions: ['worker', 'browser'], minify: true, legalComments: 'inline',
  outfile: 'src/vendor/anthropic-sdk.js', plugins: [stub],
});
console.log('ok');
