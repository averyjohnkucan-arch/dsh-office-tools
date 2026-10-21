import {readFile,stat} from 'node:fs/promises';
import {basename} from 'node:path';
import Schema from '@deepseek-ai/schemastery';
import {defineTool} from '@deepseek-ai/dsh-tools';
import {approveEscalation} from '@deepseek-ai/dsh-sandbox';
import {OfficeRuntime,workspaceOf,projectPath} from './lib/runtime.js';
import {setup} from './scripts/setup.js';
export const inject=['tools','skills','attachments','sandboxPolicy'];
export const Config=Schema.object({dataDir:Schema.string().default(''),pythonExecutable:Schema.string().default(''),sofficeExecutable:Schema.string().default('')});
const output={schema:{type:'object',additionalProperties:true},render:(_a,v)=>[{type:'text',text:JSON.stringify(v)}]};
export async function apply(ctx,config={}){
 const runtime=new OfficeRuntime(config),tools=JSON.parse(await readFile(new URL('./data/tools.json',import.meta.url),'utf8'));
 ctx.effect(()=>()=>runtime.close());
 for(const tool of tools)ctx.effect(()=>ctx.tools.register({name:'office_'+tool.name,description:tool.description+' Local Office-MCP. Paths stay in the conversation workspace. Use office_present for the final downloadable file.',parameters:tool.inputSchema,output,timeoutMs:150000,isConcurrencySafe:()=>false,execute:(args,exec)=>runtime.call(tool.name,args,exec)}));
 ctx.effect(()=>ctx.skills.register({name:'office-excel',description:'Create and edit Excel XLSX workbooks locally using Office-MCP: sheets, cells, formulas, formatting and charts, with a downloadable spreadsheet file.',provider:'runtime',source:'runtime',content:'For an Excel task, discover the office_excel_* tools. office_excel_create_workbook creates the workbook; office_excel_write_range writes a rectangular data array; values starting with = are formulas. Inspect the exact tool schemas before calling them. Use office_excel_format_cells and office_excel_add_chart as needed. Include the header row when selecting a chart data_range, for example A1:B3 for two products and their units. Keep all paths workspace-relative. Inspect the finished workbook with office_excel_get_info and office_excel_read_sheet, verify expected values and formulas, and finish with office_present(paths=[...]) so the user receives the XLSX attachment. Excel evaluates formulas when opened; creating a formula does not prove its cached result was recalculated. PDF/HTML exports require LibreOffice. If tools say the runtime is missing, run office_tools(action="setup") under the normal DSH approval policy. Never claim a file was created until the actual tool succeeds.'}));
 ctx.effect(()=>ctx.tools.register(defineTool({name:'office_tools',description:'Local Office document editing setup and discovery. Real Excel XLSX, Word DOCX and PowerPoint PPTX tools are named office_excel_*, office_word_* and office_pptx_*. Use menu to see names, status to check the runtime, setup to install isolated dependencies. Finish with office_present.',parameters:{action:{type:'string',enum:['menu','status','setup'],default:'menu'}},output,timeoutMs:660000,isConcurrencySafe:()=>false,execute:async(args,exec)=>{
  if(args.action==='setup'){
   const policy=ctx.sandboxPolicy.resolve(exec.agent?{session:exec.agent.session}:{});
   await approveEscalation({requestedMode:'danger-full-access',effectiveMode:policy.mode,subject:'Office runtime setup',justification:'Install isolated Python dependencies for the bundled Office-MCP server.'},{approver:ctx.get('approval'),agent:exec.agent,callId:exec.callId,toolName:'office_tools',signal:exec.signal});
   return setup({dataDir:config.dataDir,signal:exec.signal});
  }
  return {...await runtime.status(),tools:args.action==='status'?undefined:tools.map(t=>'office_'+t.name),skill:'office-excel',present:'office_present',source:'https://github.com/gawirable/office-mcp',upstreamCommit:'cf210cb6581f33cf3ceb57cf1defff4fc58d050b'};
 }})));
 ctx.effect(()=>ctx.tools.register(defineTool({name:'office_present',description:'Attach finished Excel, Word, PowerPoint, PDF or CSV files from the conversation workspace for the user to download. Call after validating the document.',parameters:{paths:{type:'array',items:{type:'string'},required:true}},timeoutMs:30000,isConcurrencySafe:()=>true,output:{schema:{type:'object',additionalProperties:true},render:(_a,value)=>[{type:'text',text:value.files.map(f=>`[${f.name}](<${f.path}>)`).join('\n')},...value.files.map(f=>({type:'file',attachment:f.attachment}))]},execute:async({paths},exec)=>{
  if(!Array.isArray(paths)||!paths.length||paths.length>12)throw Error('Present 1–12 workspace files.');const workspace=await workspaceOf(exec),files=[];
  for(const path of paths){const file=await projectPath(workspace,path),info=await stat(file);if(!info.isFile()||info.size>256*1024**2)throw Error('File is not regular or exceeds 256 MiB.');files.push({name:basename(file),path:file,bytes:info.size,attachment:await ctx.attachments.saveFile({data:await readFile(file),name:basename(file)})});}
  return {files};
 }})));
}
