import { mkdir, copyFile } from 'node:fs/promises';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const output = resolve(repo, 'gateway/dist');
for (const relative of ['index.html', 'src/portfolio-integrity.js', 'assets/multsoft-mb-ima.png']) {
  const target = resolve(output, relative);
  await mkdir(dirname(target), { recursive: true });
  await copyFile(resolve(repo, relative), target);
}
