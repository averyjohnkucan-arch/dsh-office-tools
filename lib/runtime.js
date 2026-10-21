import {Client} from '@modelcontextprotocol/client';
import {StdioClientTransport} from '@modelcontextprotocol/client/stdio';
import {realpath,stat,access,mkdir} from 'node:fs/promises';
import {resolve,join,relative,dirname,isAbsolute} from 'node:path';
import {homedir} from 'node:os';
import {fileURLToPath} from 'node:url';
export const pkg=fileURLToPath(new URL('../',import.meta.url));
export const rootOf=config=>resolve(config.dataDir||join(process.env.DSH_HOME||join(homedir(),'.dsh'),'dsh-office-tools'));
export const pythonOf=config=>config.pythonExecutable||join(rootOf(config),'venv',process.platform==='win32'?'Scripts':'bin',process.platform==='win32'?'python.exe':'python');
export async function workspaceOf(exec){const cwd=exec.workspace||exec.agent?.session.header?.cwd||exec.agent?.session.meta?.cwd;if(!cwd)throw Error('Office tools need a conversation workspace.');return realpath(cwd);}
export async function projectPath(workspace,path){
 if(typeof path!=='string'||!path.trim())throw Error('A workspace-relative file path is required.');
 const dest=resolve(workspace,path),within=p=>{const r=relative(workspace,p);return !r||r!=='..'&&!r.startsWith('..'+(process.platform==='win32'?'\\':'/'))&&!isAbsolute(r)};
 if(!within(dest))throw Error('Office paths must remain inside the conversation workspace.');
 let parent=dest;for(;;){try{const actual=await realpath(parent);if(!within(actual))throw Error('Office path follows a symlink outside the workspace.');break}catch(error){if(error.code!=='ENOENT')throw error;const next=dirname(parent);if(next===parent)throw error;parent=next}}
 return dest;
}
export class OfficeRuntime{
 constructor(config={}){this.config=config;this.connections=new Map();}
 async connect(workspace,signal){
  if(this.connections.has(workspace))return this.connections.get(workspace);
  const pending=(async()=>{
   const python=pythonOf(this.config);try{await access(python)}catch{throw Error('Office runtime is not installed. Run office_tools(action="setup") or the plugin setup script first.')}
   await mkdir(join(workspace,'.office'),{recursive:true});
   const env={...process.env,PYTHONNOUSERSITE:'1',PYTHONDONTWRITEBYTECODE:'1',PYTHONUTF8:'1',OFFICE_MCP_DEFAULT_FOLDER:workspace,...(this.config.sofficeExecutable?{OFFICE_MCP_SOFFICE:this.config.sofficeExecutable}:{})};
   const transport=new StdioClientTransport({command:python,args:['-B',join(pkg,'vendor/office-mcp/server.py')],cwd:workspace,env,stderr:'pipe'});
   let stderr='';transport.stderr?.on('data',chunk=>{stderr=(stderr+chunk).slice(-4000)});
   const client=new Client({name:'dsh-office-tools',version:'0.1.0'});
   try{await client.connect(transport,{signal,timeout:30000});return {client,transport}}catch(error){await client.close().catch(()=>{});throw Error('Office-MCP startup failed: '+error.message+' '+stderr)}
  })();this.connections.set(workspace,pending);
  try{return await pending}catch(error){this.connections.delete(workspace);throw error}
 }
 async call(name,args,exec){
  const workspace=await workspaceOf(exec),connection=await this.connect(workspace,exec.signal),safe={...args};
  // All bundled upstream file parameters use these names. The folder argument
  // always belongs to the conversation, regardless of model-supplied input.
  for(const key of ['path','output','image_path','output_dir'])if(safe[key])safe[key]=await projectPath(workspace,safe[key]);
  if('folder' in safe)safe.folder=workspace;
  const result=await connection.client.callTool({name,arguments:safe},{signal:exec.signal,timeout:120000});
  if(result.isError)throw Error(result.content.filter(x=>x.type==='text').map(x=>x.text).join('\n'));
  return result;
 }
 async status(){try{await access(pythonOf(this.config));return {ready:true,python:pythonOf(this.config)}}catch{return {ready:false,next:'Run office_tools(action="setup") to install this plugin’s isolated Python runtime.'}}}
 async close(){const pending=[...this.connections.values()];this.connections.clear();await Promise.allSettled(pending.map(async p=>(await p).client.close()))}
}
