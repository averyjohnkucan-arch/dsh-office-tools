import test from 'node:test';
import assert from 'node:assert/strict';
import { Context } from '@deepseek-ai/cordis';
import SystemPrompt from '@deepseek-ai/dsh-system-prompt';
import Tools from '@deepseek-ai/dsh-tools';
import { replaceStockOffice, withoutStockOffice } from '../lib/replace-stock.js';

const stock = { id: 'tool-office', name: '@deepseek-ai/dsh-tool-office' };
const present = { id: 'present', name: '@deepseek-ai/dsh-tool-present' };

test('only the stock Office provider is removed from a preset', () => {
  const original = { id: 'standard', plugins: [stock, present] };
  assert.deepEqual(withoutStockOffice(original), { id: 'standard', plugins: [present] });
  assert.deepEqual(original.plugins, [stock, present]);
});

test('RC4 assembles chat prompts while the Office preset replacement is active', async t => {
  const ctx = new Context();
  new SystemPrompt(ctx, {});
  new Tools(ctx);
  t.after(() => ctx.fiber.dispose());
  ctx.tools.setExecutor({ describe: definitions => ({ schemas: definitions }) });
  const updates = [];
  const raw = { id: 'standard', plugins: [stock, present] };
  const preset = {
    entry: { options: { name: '@deepseek-ai/dsh-agent-preset' } },
    config: raw,
    uid: 1,
    restart: async () => {
      updates.push(ctx.waterfall('internal/config', preset, () => raw));
    },
  };
  const activeContext = {
    on: ctx.on.bind(ctx),
    tools: ctx.tools,
    registry: new Map([['preset', { fibers: [preset] }]]),
  };
  const restore = await replaceStockOffice(activeContext);
  assert.deepEqual(updates[0].plugins, [present]);
  const newer = { ...preset, config: { id: 'ptc', plugins: [stock, present] }, uid: null };
  assert.deepEqual(ctx.waterfall('internal/config', newer, () => newer.config).plugins, [present]);
  assert.match(ctx.tools.checkGuards({ name: 'office_render' }).reason, /disabled/);
  assert.equal(ctx.tools.checkGuards({ name: 'office_excel_create_workbook' }), undefined);
  assert.deepEqual((await ctx.systemPrompt.assemble()).tools, []);
  for (const name of ['office_excel_create_workbook', 'present']) {
    ctx.tools.register({
      name, description: name, parameters: { type: 'object', properties: {} },
      output: { schema: { type: 'object' }, render: () => [] }, execute: async () => ({}),
    });
  }
  const assembly = await ctx.systemPrompt.assemble();
  assert.deepEqual(assembly.tools.map(tool => tool.name).sort(), ['office_excel_create_workbook', 'present']);
  await restore();
  assert.deepEqual(updates[1].plugins, [stock, present]);
  assert.equal(ctx.tools.checkGuards({ name: 'office_render' }), undefined);
});
