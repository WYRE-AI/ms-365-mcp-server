import { pathToFileURL } from 'node:url';
import { createAppOnlyServer } from './server.js';
import { guardAppOnly } from './guard.js';
import { numeric } from './errors.js';
export async function start(): Promise<void> {
  guardAppOnly(process.env, process.argv.slice(2));
  const { server, host } = createAppOnlyServer();
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(numeric(process.env, 'MCP_HTTP_PORT', 8080), host, resolve);
  });
}
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  start().catch((error) => {
    console.error(error instanceof Error ? error.message : 'App-only startup failed');
    process.exitCode = 1;
  });
}
