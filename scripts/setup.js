import {spawn} from 'node:child_process';
import {mkdir,writeFile,access} from 'node:fs/promises';
import {join,resolve} from 'node:path';
import {homedir} from 'node:os';
import {fileURLToPath} from 'node:url';
const pkg=fileURLToPath(new URL('../',import.meta.url));
export function run(command,args,{signal,cwd}={}){return new Promise((yes,no)=>{const p=spawn(command,args,{cwd,signal,stdio:['ignore','pipe','pipe']});let stdout='',stderr='';p.stdout.on('data',b=>stdout=(stdout+b).slice(-20000));p.stderr.on('data',b=>stderr=(stderr+b).slice(-20000));p.once('error',no);p.once('close',code=>code===0?yes(stdout):no(Error(`${command} exited ${code}: ${stderr}`)));})}
export async function setup({dataDir,pythonExecutable,signal}={}){
 const root=resolve(dataDir||join(process.env.DSH_HOME||join(homedir(),'.dsh'),'dsh-office-tools')),venv=join(root,'venv'),python=join(venv,process.platform==='win32'?'Scripts':'bin',process.platform==='win32'?'python.exe':'python');
 await mkdir(root,{recursive:true});const sourcePython=pythonExecutable||process.env.PYTHON||(process.platform==='win32'?'python':'python3');
 try{await access(python)}catch{await run(sourcePython,['-m','venv','--without-pip',venv],{signal})}
 try{await run(python,['-m','pip','--version'],{signal})}catch{
  const r=await fetch('https://bootstrap.pypa.io/get-pip.py',{signal:AbortSignal.any([...(signal?[signal]:[]),AbortSignal.timeout(30000)])});if(!r.ok)throw Error('Could not download the official PyPA pip bootstrap.');const target=join(root,'get-pip.py');await writeFile(target,await r.text());await run(python,[target,'--disable-pip-version-check'],{signal});
 }
 await run(python,['-m','pip','install','--disable-pip-version-check','--no-input','-r',join(pkg,'data/python-requirements.txt'),join(pkg,'vendor/office-mcp')],{signal});
 return {ready:true,python,note:'Excel/DOCX/PPTX editing is ready. PDF/HTML conversion additionally needs LibreOffice on PATH or sofficeExecutable configured.'};
}
if(process.argv[1]&&resolve(process.argv[1])===fileURLToPath(import.meta.url)){const args=process.argv.slice(2);console.log(JSON.stringify(await setup({dataDir:args[args.indexOf('--data-dir')+1]&&args.includes('--data-dir')?args[args.indexOf('--data-dir')+1]:undefined,pythonExecutable:args.includes('--python')?args[args.indexOf('--python')+1]:undefined,signal:AbortSignal.timeout(600000)})));}
