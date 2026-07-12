import { cp, mkdir, rm } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';

await rm('dist', { recursive: true, force: true });
await mkdir('dist', { recursive: true });
await cp('index.html', 'dist/index.html');
await cp('styles.css', 'dist/styles.css');
await cp('CNAME', 'dist/CNAME');

const result = spawnSync('webpack', ['--config', 'webpack.config.cjs', '--mode', 'production'], {
  stdio: 'inherit',
});
if (result.status !== 0) process.exit(result.status || 1);
