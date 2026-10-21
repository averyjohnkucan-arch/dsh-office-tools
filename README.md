# DSH Office Tools

Standalone plugin for DeepSeek Harness 0.2.1-rc.2. Install from this GitHub repository using DSH Plugins, or `dsh plugin --profile web add <repository-url> --config.ignore-scripts=true`.

The plugin bundles the unchanged [gawirable/office-mcp](https://github.com/gawirable/office-mcp) server at commit `cf210cb6581f33cf3ceb57cf1defff4fc58d050b`. It exposes 47 real MCP editing tools through the public DSH tools registry, adds an `office-excel` workflow through the public skill registry, and provides `office_present` for downloadable attachments. No Harness source modifications, old installation paths or external tool plugins are required.

## Setup

Run `node scripts/setup.js` from the installed plugin directory, or select `office_tools(action="setup")` in DSH. Setup creates an isolated Python environment under `$DSH_HOME/dsh-office-tools/venv`. It requires Python 3.10+ (`python3` on Linux, `python` on Windows). If Python lacks pip, setup uses the official PyPA bootstrap inside the isolated environment. `--python /path/to/python` and `--data-dir /path/to/state` override setup locations.

DSH plugin configuration supports `dataDir`, `pythonExecutable` (an already provisioned Python containing the dependencies) and `sofficeExecutable`. Excel, Word and PowerPoint editing works without LibreOffice. PDF/HTML conversion requires a separate LibreOffice installation. Python packages are installed on the target computer; they are not committed into this repository. Bundled source and tool schemas ship with the plugin.

## Model workflow

Ask DSH to create an Excel workbook. Tool discovery can find `office_tools`, the `office-excel` skill and the `office_excel_*` tools by their descriptions. The model creates and edits a workspace file, verifies it, then calls `office_present` so the user gets the editable XLSX attachment. Formula strings are preserved; formula caches are not proof of recalculation. Existing external files must first be copied into the conversation workspace. Paths and symlink targets are restricted to that workspace.

Word and PowerPoint tools use the same workflow with `office_word_*` and `office_pptx_*`. MCP failures are surfaced rather than presented as successful edits. Connections close when the plugin unloads.

## Validation

`node --test test/*.test.js` exercises real MCP workbook creation, formula/data writes, formatting, chart creation, readback and the DSH tool/skill registry and attachment output. Tests require setup first. Windows path/runtime branches are included, but native Windows execution has not been tested on the Linux development host.
