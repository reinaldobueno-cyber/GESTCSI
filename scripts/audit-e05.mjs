import { readFile } from 'node:fs/promises';

const method = JSON.parse(await readFile(new URL('../governance/e05-http-method-contract.json', import.meta.url), 'utf8'));
const source = await readFile(new URL('../apps_script/ClickUpSync.gs', import.meta.url), 'utf8');
const router = source.slice(source.indexOf('function doGet(e)'), source.indexOf('function legacyActionPolicy_'));
const mutableGet = method.post_required.filter((action) => router.includes(`action === '${action}'`));
const result = {
  stage: 'E05',
  ready: mutableGet.length === 0 && method.conditional_effect_get.length === 0 && method.read_candidate.length === 0,
  actions_reviewed: method.read_verified.length + method.read_candidate.length + method.conditional_effect_get.length + method.post_required.length,
  mutable_get_remaining: mutableGet.length,
  conditional_get_remaining: method.conditional_effect_get.length,
  read_candidates_unverified: method.read_candidate.length,
  mutable_get_examples: mutableGet.slice(0, 8),
  conditional_get_actions: method.conditional_effect_get
};
process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
