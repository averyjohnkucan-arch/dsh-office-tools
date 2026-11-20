import {randomBytes,timingSafeEqual} from 'node:crypto';
import {setup,run} from '../scripts/setup.js';

export function attachOfficeSetup(ctx,runtime,config) {
  const path='/plugins/dsh-office-tools/setup';
  const token=randomBytes(32).toString('hex');
  const connection={path,token};
  let controller=null, pending=null, status='idle', detail='';
  ctx.on('webserver/index-inject',items=>items.push({kind:'global',name:'__DSH_OFFICE_TOOLS__',value:connection}));
  ctx.inject(['connection'],scope=>scope.effect(()=>ctx.webServer.register({
    kind:'exact',path:path+'/connection',handler(req,res){
      if(req.method!=='GET'||req.headers.origin&&req.headers.origin!==ctx.webServer.getOrigin()){res.writeHead(403);res.end();return;}
      if(!scope.connection.authorizeIndex(req,res))return;
      res.setHeader('Cache-Control','no-store');res.setHeader('Content-Type','application/json');res.end(JSON.stringify(connection));
    },
  })));
  ctx.effect(()=>ctx.webServer.register({
    kind:'exact',path,async handler(req,res) {
      res.setHeader('Cache-Control','no-store');res.setHeader('Content-Type','application/json');
      const actual=Buffer.from((req.headers.authorization||'').replace(/^Bearer /,''));
      const expected=Buffer.from(token);
      if(actual.length!==expected.length||!timingSafeEqual(actual,expected)){res.writeHead(401);res.end('{}');return;}
      if(req.method!=='POST'||!String(req.headers['content-type']).startsWith('application/json')||req.headers.origin&&req.headers.origin!==ctx.webServer.getOrigin()){res.writeHead(403);res.end('{}');return;}
      try {
        let body='';
        for await(const chunk of req){body+=chunk;if(body.length>8192)throw Error('Request too large.');}
        const {action}=JSON.parse(body);
        if(action==='setup'){
          if(pending)throw Error('Office setup is already running.');
          if(config.pythonExecutable)throw Error('This installation uses a custom Python. Check that runtime or clear the override to install the managed one.');
          controller=new AbortController();status='running';detail='Installing the private Office MCP runtime.';
          pending=setup({dataDir:config.dataDir,signal:AbortSignal.any([controller.signal,AbortSignal.timeout(600000)])})
            .then(()=>{status='complete';detail='Office runtime installed.';})
            .catch(error=>{status=controller.signal.aborted?'cancelled':'failed';detail=error.message;})
            .finally(()=>{controller=null;pending=null;});
        } else if(action==='cancel')controller?.abort();
        else if(action!=='status')throw Error('Unknown Office setup action.');
        const readiness=await runtime.status();
        if(readiness.ready && !pending) {
          try {await run(readiness.python,['-B','-c','import docx, openpyxl, pptx, mcp'],{signal:AbortSignal.timeout(10000)});}
          catch(error){readiness.ready=false;readiness.next=error.message;}
        }
        res.end(JSON.stringify({status,detail,readiness}));
      } catch(error){res.writeHead(400);res.end(JSON.stringify({error:error.message}));}
    },
  }));
  ctx.effect(()=>async()=>{controller?.abort();await pending;});
}
