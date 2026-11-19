import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { Context } from '@deepseek-ai/cordis';
import Skills from '@deepseek-ai/dsh-skill';
import SystemPrompt from '@deepseek-ai/dsh-system-prompt';
import Tools from '@deepseek-ai/dsh-tools';
import { executeToolBatch, ToolExecutionPipeline } from '@deepseek-ai/dsh-tools/pipeline';
import { apply } from '../index.js';

test('Harness accepts both the compact Office status and the default menu', async t => {
  const dataDir = await mkdtemp(join(tmpdir(), 'office-status-'));
  t.after(() => rm(dataDir, { recursive: true, force: true }));
  const ctx = new Context();
  t.after(() => ctx.fiber.dispose());
  new SystemPrompt(ctx, {});
  new Tools(ctx);
  new Skills(ctx);
  await apply(ctx, { dataDir });
  const calls = [{ action: 'status' }, {}].map((args, index) => ({
    type: 'tool_call', id: String(index), name: 'office_tools', arguments: JSON.stringify(args),
  }));
  const { results } = await executeToolBatch(
    { calls, signal: AbortSignal.timeout(5000) }, new ToolExecutionPipeline(ctx), () => 1,
  );
  const [status, menu] = results.map(({ result }) => {
    assert.equal(result.isError, false, result.content[0].text);
    return JSON.parse(result.content[0].text);
  });
  assert.equal(status.ready, false);
  assert.ok(status.next);
  assert.equal(Object.hasOwn(status, 'tools'), false);
  assert.ok(menu.tools.includes('office_excel_create_workbook'));
});
