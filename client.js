window.__ModuleLoader__.load({id:'dsh-office-tools',factory(require){
  const {createElement:h,useEffect,useState}=require('react');
  let connection=window.__DSH_OFFICE_TOOLS__;
  async function request(action,retry=true){
    if(!connection){
      const response=await fetch('/plugins/dsh-office-tools/setup/connection');
      if(!response.ok)throw Error('Office Tools settings are unavailable.');
      connection=await response.json();
    }
    const response=await fetch(connection.path,{
      method:'POST',headers:{'Content-Type':'application/json',Authorization:'Bearer '+connection.token},body:JSON.stringify({action}),
    });
    if(response.status===401&&retry){connection=null;return request(action,false);}
    const result=await response.json();
    if(!response.ok)throw Error(result.error||'Office settings request failed.');
    return result;
  }
  function Page(){
    const [value,setValue]=useState(null),[error,setError]=useState('');
    const running=value?.status==='running';
    useEffect(()=>{
      let active=true;
      const refresh=()=>request('status').then(v=>active&&setValue(v)).catch(e=>active&&setError(e.message));
      void refresh();
      if(!running)return ()=>{active=false;};
      const timer=setInterval(refresh,1800);
      return ()=>{active=false;clearInterval(timer);};
    },[running]);
    const select=action=>{setError('');void request(action).then(setValue).catch(e=>setError(e.message));};
    return h('section',{style:{display:'grid',gap:12}},
      h('h2',null,'Office Tools'),
      h('p',null,'Install the private Python runtime for Excel, Word and PowerPoint editing. PDF and HTML export also use LibreOffice.'),
      value&&h('p',{role:'status'},running?value.detail:value.readiness.ready?'Office runtime ready.':value.detail||'Office runtime setup needed.'),
      error&&h('p',{role:'alert'},error),
      h('div',null,
        h('button',{disabled:running,onClick:()=>select('setup')},'Install Office runtime'),
        h('button',{disabled:running,onClick:()=>select('status')},'Check runtime'),
        running&&h('button',{onClick:()=>select('cancel')},'Cancel')));
  }
  return {inject:['slots'],apply(ctx){
    ctx.slots.inject('settings.section',()=>ctx.slots.register({name:'settings.section',id:'office-tools',order:73,label:()=> 'Office Tools'},Page));
  }};
}});
