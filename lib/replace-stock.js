const stockProvider = '@deepseek-ai/dsh-tool-office';
const presetProvider = '@deepseek-ai/dsh-agent-preset';
const stockTools = new Set(['office_render', 'office_convert', 'office_recalculate']);

export function withoutStockOffice(config) {
  if (!Array.isArray(config?.plugins)) return config;
  const plugins = config.plugins.filter(plugin => plugin?.name !== stockProvider);
  return plugins.length === config.plugins.length ? config : { ...config, plugins };
}

function isPreset(fiber) {
  return fiber.entry?.options.name === presetProvider;
}

export async function replaceStockOffice(ctx) {
  const presets = new Set();
  const remove = ctx.on('internal/config', (fiber, next) => {
    const config = next();
    if (!isPreset(fiber)) return config;
    const replacement = withoutStockOffice(config);
    if (replacement !== config) presets.add(fiber);
    return replacement;
  }, { global: true });

  for (const runtime of ctx.registry.values()) {
    for (const fiber of runtime.fibers) {
      if (isPreset(fiber) && withoutStockOffice(fiber.config) !== fiber.config) presets.add(fiber);
    }
  }
  const restart = async () => {
    for (const fiber of presets) if (fiber.uid !== null) await fiber.restart();
  };
  // Existing conversations can retain a revision created before the preset changed.
  const unguard = ctx.tools.guard(({ name }) => stockTools.has(name)
    ? 'Stock Office tools are disabled by DSH Office Tools. Use its Office editing and export tools.'
    : undefined);
  const unfilter = ctx.on('system-prompt/assemble', async (_assembly, _context, next) => {
    const assembly = await next();
    return { ...assembly, tools: assembly.tools.filter(tool => !stockTools.has(tool.name)) };
  });

  const dispose = async () => {
    remove();
    unguard();
    unfilter();
    await restart();
  };
  try {
    await restart();
  } catch (error) {
    await dispose();
    throw error;
  }
  return dispose;
}
