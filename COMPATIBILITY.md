# DSH compatibility — 0.1.5

The Office status action now returns a JSON object accepted by the Harness tool dispatcher. The menu also returns the editing tool names. Both forms run through the RC4 tool execution pipeline in the regression test.

## 0.1.4

Fixed a first-turn chat failure on DeepSeek Harness 0.2.2-rc.4. Harness attaches tool declarations after the prompt presentation hook returns. Office Tools now relies on the preset to replace the stock provider and on its dispatch guard for sessions pinned to older presets. The regression test calls the Harness prompt builder with both an empty and a populated tool registry.

## 0.1.3

Verified on DeepSeek Harness **0.2.2-rc.4** with Node.js 24. A clean Web boot verified that Standard and PTC omit the stock Office tool provider, the stock Office skill is disabled, and the plugin's editing tools are registered. Tests exercise the preset filter, the dispatch guard for older conversation revisions, restoration on disposal, and real Office-MCP workbook editing and attachment delivery.

# DeepSeek Harness RC4

This build targets DSH 0.2.2-rc.4 and Node.js 24.9 or newer. The following notes describe validation of the previous release.

# DSH compatibility — 0.1.2

Reviewed on November 14, 2026 against DSH **0.2.2-rc.1**, Cordis **4.1.0** and Schemastery **3.19.0**. Dependency declarations and lockfiles target the current host instead of pulling in the older DSH/Cordis API generation.

Updated host constraints and made the real workbook integration test accept an explicitly isolated Python executable. Verified workbook creation, formulas, formatting, charts, attachment output and workspace path restrictions.

DSH 0.2.2 declares Node 24.9 or newer. Compatibility checks cover the existing Node 22.22 installation as well as an isolated Node 24 runtime; the user’s system runtime was not changed. CI uses Node 24. Model servers, CUDA kernels and optional large model downloads are outside this source/API compatibility check; request serialization and scheduler tests use controlled local fixtures where appropriate.

The GitHub update does not replace an already-installed plugin automatically. Update the plugin through DSH to use this release, preserving its saved settings and model profiles.
