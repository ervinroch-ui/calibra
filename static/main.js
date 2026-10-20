(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const state = {file:null, inspect:null, current:null, currentId:null, tableLimit:60, library:[], selected:new Set(), libraryFilter:""};
  const fmt = (v,d=2) => v == null || !Number.isFinite(Number(v)) ? '—' : new Intl.NumberFormat('es-NI',{maximumFractionDigits:d}).format(Number(v));
  const nf = v => new Intl.NumberFormat('es-NI').format(v);
  const el = (tag,cls,text) => { const node=document.createElement(tag); if(cls)node.className=cls; if(text!==undefined)node.textContent=text; return node; };
  const info = (msg,type='danger') => { const n=$('notice');n.textContent=msg;n.className='notice '+type;n.classList.remove('is-hidden'); };
  const clearInfo=()=>{$('notice').classList.add('is-hidden');$('notice').textContent='';};
  let toastTimer;
  const toast=(message,error=false)=>{const t=$('toast');t.textContent=message;t.className='toast'+(error?' error':'');clearTimeout(toastTimer);toastTimer=setTimeout(()=>t.classList.add('is-hidden'),4300);};
  const renderOptions=(select,values,placeholder)=>{select.replaceChildren();if(placeholder){const opt=new Option(placeholder,'');select.add(opt);}values.forEach(v=>select.add(new Option(v.label??v, v.value??v)));};
  const prettyBytes=b=>b<1048576 ? `${fmt(b/1024,1)} KB` : `${fmt(b/1048576,2)} MB`;
  const updateFile = file => {
    clearInfo(); state.file=file;state.inspect=null;$('configuration').classList.add('is-hidden');$('sheetSelect').replaceChildren();$('downloadBatchPpt').classList.add('is-hidden');$('downloadBatchPpt').removeAttribute('href');
    if(!file){$('fileInput').value='';$('fileInfo').classList.add('is-hidden');return;}
    $('fileInfo').classList.remove('is-hidden');$('fileName').textContent=file.name;$('fileSize').textContent=prettyBytes(file.size);
    inspectFile();
  };
  const request = async(url,data)=>{
    const response=await fetch(url,{method:'POST',body:data,headers:{'X-Requested-With':'XMLHttpRequest'}});
    let body;
    try{body=await response.json();}catch{throw Error('No se recibió una respuesta válida del servidor.');}
    if(!response.ok)throw Error(body.error||`Error ${response.status}`);
    return body;
  };
  const formData = withSettings => {
    const fd=new FormData();fd.append('file',state.file);
    if($('sheetSelect').value)fd.append('sheet',$('sheetSelect').value);
    fd.append('date_from',$('dateFrom').value);
    fd.append('date_to',$('dateTo').value);
    fd.append('date_column',$('dateColumnSelect').value);
    if(withSettings){fd.append('metric',$('metricSelect').value);fd.append('label_column',$('labelSelect').value);fd.append('lsl',$('lslInput').value);fd.append('usl',$('uslInput').value);}
    return fd;
  };
  const inspectFile = async()=>{
    if(!state.file)return;
    const btn=$('runAnalysis');btn.disabled=true;
    try{
      if(state.file.size>12*1024*1024)throw Error('El archivo supera los 12 MB permitidos.');
      const data=await request('/api/inspect',formData(false));state.inspect=data;
      const currentSheet=data.selected_sheet;
      renderOptions($('sheetSelect'),data.sheets);
      $('sheetSelect').value=currentSheet;
      const numeric=data.columns.filter(c=>c.numeric);
      renderOptions($('metricSelect'),numeric.filter(c=>c.name.toLocaleLowerCase('es')!=='fecha').map(c=>c.name));
      renderOptions($('labelSelect'),data.columns.map(c=>c.name),'Usar número de fila');
      const previousDateColumn=$('dateColumnSelect').value;
      renderOptions($('dateColumnSelect'),data.columns.map(c=>c.name),'Detectar Fecha automáticamente');
      $('dateColumnSelect').value=data.columns.some(c=>c.name===previousDateColumn) ? previousDateColumn :
        (data.columns.find(c=>c.name.toLocaleLowerCase('es')==='fecha')?.name??'');
      $('rowsFound').textContent=`${nf(data.row_count)} filas`;
      const head=$('previewTable').querySelector('thead');const body=$('previewTable').querySelector('tbody');head.replaceChildren();body.replaceChildren();
      const colNames=data.columns.slice(0,6).map(c=>c.name);
      const tr=el('tr');colNames.forEach(name=>tr.append(el('th',null,name)));head.append(tr);
      data.preview.forEach(row=>{const tr=el('tr');colNames.forEach(name=>tr.append(el('td',null,String(row[name]??''))));body.append(tr);});
      $('configuration').classList.remove('is-hidden');
      if(!numeric.length)info('No encontré columnas con al menos dos valores numéricos en la muestra. Revisá los encabezados y el formato del archivo.');
      else {clearInfo();btn.disabled=false;}
    }catch(e){info(e.message);$('configuration').classList.add('is-hidden');toast(e.message,true);}
  };
  const validatePeriod=()=>{
    const a=$('dateFrom').value, b=$('dateTo').value;
    if(a&&b&&a>b){info('La fecha inicial no puede ser posterior a la final.');return false;}
    return true;
  };
  const run = async()=>{
    if(!state.file)return;
    if(!$('metricSelect').value){info('Seleccioná una columna numérica.');return;}
    if(!validatePeriod())return;
    const lsl=$('lslInput').value,usl=$('uslInput').value;
    if(lsl!=='' && usl!=='' && Number(lsl)>=Number(usl)){info('El límite superior debe ser mayor que el inferior.');return;}
    const b=$('runAnalysis');b.disabled=true;b.textContent='Analizando las mediciones…';clearInfo();
    try{
      const data=await request('/api/analyze',formData(true));
      showResult(data.result,data.id,data.filename);await Promise.all([loadHistory(),loadLibrary()]);toast('Análisis archivado en la Biblioteca de Excel.');
      $('resultados').scrollIntoView({behavior:'smooth',block:'start'});
    }catch(e){info(e.message);toast(e.message,true);}finally{b.disabled=false;b.innerHTML='Generar cartas de control <span>↗</span>';}
  };
  const runBatch=async()=>{
    if(!state.file||!validatePeriod())return;
    const b=$('runBatch');b.disabled=true;b.textContent='Procesando ensayos y preparando PowerPoint…';clearInfo();
    try{
      const data=await request('/api/analyze-all',formData(false));
      await Promise.all([loadHistory(),loadLibrary()]);
      showResult(data.first.result,data.first.id,data.first.filename);
      const link=$('downloadBatchPpt');link.href=data.pptx_url;link.classList.remove('is-hidden');
      if(data.skipped.length)info(`${data.count} ensayos procesados. No incluidos: ${data.skipped.join(' | ')}`,'warning');
      toast(`${data.count} ensayos guardados. Se está descargando el informe PowerPoint.`);
      link.click();
      $('resultados').scrollIntoView({behavior:'smooth',block:'start'});
    }catch(e){info(e.message);toast(e.message,true);}finally{b.disabled=false;b.textContent='▣ Generar PowerPoint completo (G16)';}
  };
  const setText=(id,text)=>$(id).textContent=text;
  const showResult=(r,id,filename)=>{
    state.current=r;state.currentId=id;state.tableLimit=60;
    $('resultados').classList.remove('is-hidden');
    const period=r.settings.date_filter;
    const dates=period ? ` · Período: ${period.start||'inicio'} a ${period.end||'fin'}` : '';
    setText('resultSource',`${filename} · Hoja: ${r.settings.sheet} · ${r.settings.metric}${dates}`);
    $('downloadReport').href=`/api/analysis/${id}/csv`;
    $('downloadPpt').href=`/api/analysis/${id}/pptx`;
    let tone='neutral',symbol='!',tag='REQUIERE ATENCIÓN',title='',subtitle='';
    if(r.verdict==='CUMPLE'){tone='';symbol='✓';tag='DENTRO DE ESPECIFICACIONES';title='Todas las mediciones válidas cumplen los límites.';subtitle='No se detectaron incumplimientos de los rangos definidos. Revisá aparte la estabilidad estadística.';}
    else if(r.verdict==='FUERA_DE_RANGO'){tone='bad';symbol='!';tag='FUERA DE ESPECIFICACIÓN';title='Se detectaron mediciones fuera de rango.';subtitle=`${nf(r.summary.below+r.summary.above)} medición(es) incumplen los límites configurados. Investigá los valores señalados.`;}
    else if(r.verdict==='REVISAR_OMISIONES'){title='Los valores evaluados cumplen; revisá los datos omitidos.';subtitle='Hay filas vacías o no numéricas que no se pudieron verificar.';}
    else {title='Definí límites para verificar el cumplimiento.';subtitle='Las cartas de control estadístico están disponibles, pero no es posible afirmar si el producto cumple especificaciones.';}
    $('verdictBanner').className='verdict-banner '+tone;
    setText('verdictSymbol',symbol);setText('verdictTag',tag);setText('verdictTitle',title);setText('verdictText',subtitle);
    setText('kpiTotal',nf(r.summary.total));setText('kpiInside',r.spec_defined?nf(r.summary.within):'—');
    setText('kpiPercent',r.spec_defined?`${fmt(r.summary.percent)}% de las lecturas válidas`:'Faltan límites especificados');
    setText('kpiOutside',r.spec_defined?nf(r.summary.below+r.summary.above):'—');setText('kpiSpc',nf(r.summary.spc_violations));
    const limitData=[['Promedio',r.limits.i_cl],['Control inferior (I)',r.limits.i_lcl],['Control superior (I)',r.limits.i_ucl],['Media rango móvil',r.limits.mr_cl],['Control superior (MR)',r.limits.mr_ucl],['Mínimo especificado',r.settings.lsl],['Máximo especificado',r.settings.usl],['Medición mínima',r.summary.minimum],['Medición máxima',r.summary.maximum]];
    const limitList=$('limitsList');limitList.replaceChildren();limitData.forEach(([label,v])=>{const item=el('div','limit-item');item.append(el('span',null,label),el('strong',null,fmt(v,4)));limitList.append(item);});
    const warnings=[...r.warnings];
    if(r.summary.spc_violations||r.summary.mr_violations)warnings.unshift(`Posibles causas especiales: ${r.summary.spc_violations} punto(s) fuera de límites en I y ${r.summary.mr_violations} salto(s) fuera de límite en MR.`);
    if(!r.spec_defined)warnings.push('Sin límites de especificación: no es posible evaluar la aceptación del lote.');
    if(!warnings.length)warnings.push('No se detectaron incidencias con las reglas implementadas. Revisá igualmente el contexto del proceso.');
    const warningList=$('warningList');warningList.replaceChildren();warnings.forEach(w=>{const div=el('div','alert-item'+(warnings.length===1&&w.startsWith('No se detectaron')?' good':''));div.append(el('span',null,'✳'),el('div',null,w));warningList.append(div);});
    drawCharts(r);renderRows();
  };
  const renderRows=()=>{
    const r=state.current;if(!r)return;
    const filter=$('rowFilter').value;
    const all=r.points.filter(p=>filter==='all'||(filter==='issues'&&(!p.in_specs||!p.in_control))||(filter==='spec'&&!p.in_specs)||(filter==='control'&&!p.in_control));
    const body=$('dataRows');body.replaceChildren();
    if(!all.length){const tr=el('tr'),td=el('td',null,filter==='issues'?'Sin incidencias: las mediciones cumplen las condiciones de este filtro.':'No hay registros para este filtro.');td.colSpan=5;tr.append(td);body.append(tr);}
    all.slice(0,state.tableLimit).forEach(p=>{
      const tr=el('tr');tr.append(el('td',null,String(p.index)),el('td',null,p.label),el('td',null,fmt(p.value,5)));
      let status=r.spec_defined?(p.in_specs?'Cumple':p.status):'Sin límites';
      tr.append(makeCellBadge(status,r.spec_defined?(p.in_specs?'':'bad'):'warn'));
      tr.append(makeCellBadge(p.in_control?'En control':'Fuera de control',p.in_control?'':'warn'));body.append(tr);
    });
    setText('tableCount',`Mostrando ${nf(Math.min(state.tableLimit,all.length))} de ${nf(all.length)} registros · ${nf(r.summary.omitted)} fila(s) sin lectura válida`);
    $('loadMore').classList.toggle('is-hidden',all.length<=state.tableLimit);
  };
  const makeCellBadge=(text,type)=>{const td=el('td');td.append(el('span','pill '+type,text));return td;};
  const svgEl=(tag,attributes={},text)=>{const node=document.createElementNS('http://www.w3.org/2000/svg',tag);Object.entries(attributes).forEach(([k,v])=>node.setAttribute(k,String(v)));if(text!==undefined)node.textContent=text;return node;};
  function chart(container,readings,guides,alerts,title){
    container.replaceChildren();
    if(readings.length===0)return;
    const width=750,height=325,left=56,right=15,top=20,bottom=42,plotW=width-left-right,plotH=height-top-bottom;
    const raw=[...readings.map(p=>p.v),...guides.filter(g=>g.v!==null&&Number.isFinite(g.v)).map(g=>g.v)];
    let lo=Math.min(...raw),hi=Math.max(...raw);const margin=(hi-lo||Math.abs(hi)*.03||1)*.15;lo-=margin;hi+=margin;
    const y=v=>top+(hi-v)/(hi-lo)*plotH;
    const x=i=>left+(readings.length===1?plotW/2:i*plotW/(readings.length-1));
    const svg=svgEl('svg',{viewBox:`0 0 ${width} ${height}`,role:'img','aria-label':title});
    svg.append(svgEl('title',{},title));
    for(let i=0;i<=5;i++){
      const value=hi-(hi-lo)*i/5,yp=y(value);
      svg.append(svgEl('line',{x1:left,y1:yp,x2:width-right,y2:yp,stroke:'#eaf0f4','stroke-width':1}));
      svg.append(svgEl('text',{x:left-9,y:yp+4,'text-anchor':'end',fill:'#9aabb9','font-size':11,'font-family':'DM Sans, sans-serif'},fmt(value,2)));
    }
    const stamps=new Set([0,readings.length-1]);const step=Math.max(1,Math.ceil((readings.length-1)/6));for(let i=0;i<readings.length;i+=step)stamps.add(i);
    [...stamps].sort((a,b)=>a-b).forEach(i=>svg.append(svgEl('text',{x:x(i),y:height-15,'text-anchor':'middle',fill:'#9aabb9','font-size':10,'font-family':'DM Sans, sans-serif'},String(i+1))));
    guides.forEach(g=>{
      if(g.v==null||!Number.isFinite(g.v))return;
      const attrs={x1:left,y1:y(g.v),x2:width-right,y2:y(g.v),stroke:g.color,'stroke-width':g.width??1.7,'stroke-dasharray':g.dash??'6 5',opacity:.8};svg.append(svgEl('line',attrs));
      if(g.text)svg.append(svgEl('text',{x:width-right-5,y:y(g.v)-6,'text-anchor':'end',fill:g.color,'font-size':10,'font-weight':700},g.text));
    });
    // Muestreo visual sólo para el trazo: los indicadores y el CSV cubren TODAS las mediciones.
    // Conservamos, además, todos los puntos anómalos de la serie.
    const indexes=[];const maxPoints=1400;
    if(readings.length<=maxPoints)readings.forEach((_,i)=>indexes.push(i));
    else {const choices=new Set([0,readings.length-1]);const bucket=Math.ceil(readings.length/(maxPoints/2));for(let from=0;from<readings.length;from+=bucket){let minI=from,maxI=from;for(let j=from+1;j<Math.min(readings.length,from+bucket);j++){if(readings[j].v<readings[minI].v)minI=j;if(readings[j].v>readings[maxI].v)maxI=j;}choices.add(minI);choices.add(maxI);}readings.forEach((p,i)=>{if(alerts(p,i))choices.add(i);});indexes.push(...[...choices].sort((a,b)=>a-b));}
    const path=indexes.map((i,k)=>`${k?'L':'M'}${x(i).toFixed(2)} ${y(readings[i].v).toFixed(2)}`).join(' ');
    svg.append(svgEl('path',{d:path,fill:'none',stroke:'#7465db','stroke-width':2.2,'stroke-linejoin':'round','stroke-linecap':'round'}));
    const outlierIndexes=readings.map((p,i)=>alerts(p,i)?i:-1).filter(i=>i>=0);
    outlierIndexes.slice(0,160).forEach(i=>{const c=svgEl('circle',{cx:x(i),cy:y(readings[i].v),r:4.1,fill:'#e25264',stroke:'white','stroke-width':1.5});c.append(svgEl('title',{},`${readings[i].label}\n${fmt(readings[i].v,5)} — fuera de límite`));svg.append(c);});
    const cross=svgEl('line',{y1:top,y2:height-bottom,stroke:'#94a3b8','stroke-dasharray':'3 5',opacity:0});const dot=svgEl('circle',{r:4.5,fill:'#7465db',stroke:'white','stroke-width':2,opacity:0});svg.append(cross,dot);
    const hover=svgEl('rect',{x:left,y:top,width:plotW,height:plotH,fill:'transparent'});
    const tip=el('div','chart-tooltip');Object.assign(tip.style,{position:'absolute',pointerEvents:'none',display:'none',background:'#17243b',color:'#fff',borderRadius:'8px',padding:'8px 11px',fontSize:'10px',boxShadow:'0 7px 22px #0003',zIndex:'2',maxWidth:'180px'});
    container.style.position='relative';container.append(svg,tip);
    hover.addEventListener('pointermove',e=>{
      const bounds=svg.getBoundingClientRect(),sx=(e.clientX-bounds.left)/bounds.width*width;
      let i=Math.round((sx-left)/plotW*(readings.length-1));i=Math.max(0,Math.min(readings.length-1,i));const p=readings[i];
      const xp=x(i),yp=y(p.v);cross.setAttribute('x1',xp);cross.setAttribute('x2',xp);cross.setAttribute('opacity','1');dot.setAttribute('cx',xp);dot.setAttribute('cy',yp);dot.setAttribute('opacity','1');
      tip.textContent=`${p.label}  ·  ${fmt(p.v,5)}`;tip.style.display='block';tip.style.left=`${Math.min(bounds.width-170,Math.max(5,xp/width*bounds.width+10))}px`;tip.style.top=`${Math.max(0,yp/height*bounds.height-28)}px`;
    });
    hover.addEventListener('pointerleave',()=>{cross.setAttribute('opacity','0');dot.setAttribute('opacity','0');tip.style.display='none';});
    svg.append(hover);
  }
  const drawCharts=r=>{
    const l=r.limits,s=r.settings,pts=r.points;
    const v=pts.map(p=>({v:p.value,label:p.label}));
    chart($('chartIndividuals'),v,[
      {v:l.i_cl,color:'#17aaa3',dash:'',width:2},{v:l.i_lcl,color:'#e6a143'},{v:l.i_ucl,color:'#e6a143'},
      {v:s.lsl,color:'#e35d68',dash:'5 7'},{v:s.usl,color:'#e35d68',dash:'5 7'}],
    (_,i)=>!pts[i].in_control||!pts[i].in_specs,'Carta de individuales (I)');
    const m=pts.slice(1).map((p,i)=>({v:p.mr,label:`${pts[i].label} → ${p.label}`}));
    chart($('chartMR'),m,[{v:l.mr_cl,color:'#17aaa3',dash:'',width:2},{v:l.mr_ucl,color:'#e6a143'}],p=>p.v>l.mr_ucl,'Carta de rango móvil (MR)');
  };
  const statusClass=code=>code==='VERDE'?'green':code==='AMARILLO'?'yellow':code==='ROJO'?'red':'gray';
  const libraryDate=date=>new Date(date).toLocaleString('es-NI',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
  const downloadGroup=id=>{window.location.href=`/api/library/${id}/pptx`;};
  const openSavedAnalysis=async id=>{
    const response=await fetch(`/api/analysis/${id}`);
    if(!response.ok)throw Error('No pude recuperar las cartas de control guardadas.');
    const data=await response.json();showResult(data.result,data.id,data.filename);
    $('resultados').scrollIntoView({behavior:'smooth',block:'start'});
  };
  const updateSelection=()=>{
    const total=state.selected.size;
    $('selectedCount').textContent=String(total);
    $('compilePpt').disabled=total===0||total>10;
    $('compilePpt').title=total>10?'El compilado admite máximo 10 archivos.':'';
  };
  const renderLibrary=()=>{
    const area=$('libraryContainer');area.replaceChildren();
    const query=state.libraryFilter.toLocaleLowerCase('es').trim();
    const shown=state.library.filter(f=>!query||(`${f.filename} ${f.kind} ${f.created_at}`).toLocaleLowerCase('es').includes(query));
    if(!shown.length){const node=el('div','empty-state');node.append(el('span',null,'▤'),el('strong',null,'No hay archivos en esta búsqueda'),el('p',null,'Probá con otro nombre o analizá un nuevo Excel.'));area.append(node);return;}
    shown.forEach(file=>{
      const wrapper=el('article','library-file');
      const checkbox=el('input','library-checkbox');checkbox.type='checkbox';checkbox.checked=state.selected.has(file.id);
      checkbox.setAttribute('aria-label',`Seleccionar ${file.filename} para compilado`);
      checkbox.addEventListener('change',()=>{if(checkbox.checked)state.selected.add(file.id);else state.selected.delete(file.id);updateSelection();});
      const icon=el('span','library-file-icon','▤');
      const description=el('div','library-file-main');
      const title=el('strong','library-file-title',file.filename);
      const filter=file.analyses?.find(item=>item.period)?.period;
      const periodText=filter ? ` · Período: ${filter.start||'inicio'} — ${filter.end||'fin'}` : '';
      const info=el('div','library-file-details',`${libraryDate(file.created_at)}  ·  ${prettyBytes(file.size)}  ·  ${file.kind}  ·  ${nf(file.count)} ensayo(s)  ·  ${nf(file.sample_count)} mediciones${periodText}`);
      const badges=el('div','library-badges');
      for(const [status,count] of Object.entries(file.statuses)){if(count){const pill=el('span','library-status '+statusClass(status),`${status} ${count}`);badges.append(pill);}}
      description.append(title,info,badges);
      const actions=el('div','library-actions');
      const download=el('a','library-download','▣ PowerPoint');download.href=`/api/library/${file.id}/pptx`;download.setAttribute('download','');
      const excel=el('a','library-aux','↓ Excel original');excel.href=`/api/library/${file.id}/excel`;excel.setAttribute('download','');
      const expand=el('button','library-aux','Ver ensayos ▾');expand.type='button';
      const remove=el('button','library-delete','Eliminar');remove.type='button';remove.title='Eliminar copia archivada y resultados de ese análisis';
      const details=el('div','library-detail is-hidden');
      file.analyses.forEach(item=>{
        const line=el('div','library-analysis');
        const name=el('div','library-analysis-name');name.append(el('strong',null,item.metric),el('small',null,`${item.sheet} · ${nf(item.total)} lecturas · ${nf(item.outside)} fuera`));
        const btn=el('button','library-aux','Abrir carta ↗');btn.type='button';btn.addEventListener('click',async()=>{btn.disabled=true;try{await openSavedAnalysis(item.id);}catch(error){toast(error.message,true);}finally{btn.disabled=false;}});
        line.append(name,btn);details.append(line);
      });
      if(file.skipped.length){const caution=el('p','library-skipped',`Ensayos omitidos: ${file.skipped.join(' | ')}`);details.append(caution);}
      expand.addEventListener('click',()=>{details.classList.toggle('is-hidden');expand.textContent=details.classList.contains('is-hidden')?'Ver ensayos ▾':'Ocultar ensayos ▴';});
      remove.addEventListener('click',async()=>{
        if(!confirm(`¿Eliminar de forma definitiva el registro «${file.filename}» y sus ${file.count} análisis? Esta acción no borra tu Excel de la computadora.`))return;
        remove.disabled=true;
        try{
          const response=await fetch(`/api/library/${file.id}`,{method:'DELETE',headers:{'X-Requested-With':'XMLHttpRequest'}});
          const body=await response.json();if(!response.ok)throw Error(body.error||'No se pudo eliminar el archivo.');
          state.selected.delete(file.id);await Promise.all([loadLibrary(),loadHistory()]);toast('Registro archivado eliminado.');
        }catch(error){toast(error.message,true);remove.disabled=false;}
      });
      actions.append(download,excel,expand,remove);
      const top=el('div','library-file-top');top.append(checkbox,icon,description,actions);
      wrapper.append(top,details);area.append(wrapper);
    });
  };
  const loadLibrary=async()=>{
    try{
      const response=await fetch('/api/library');if(!response.ok)throw Error('No pude cargar la biblioteca.');
      const data=await response.json();state.library=data.items;
      const current=new Set(data.items.map(f=>f.id));
      for(const id of state.selected){if(!current.has(id))state.selected.delete(id);}
      setText('fileCount',String(data.items.length));setText('libraryFiles',nf(data.items.length));
      setText('libraryAnalyses',nf(data.items.reduce((acc,f)=>acc+f.count,0)));
      setText('libraryMeasurements',nf(data.items.reduce((acc,f)=>acc+f.sample_count,0)));
      renderLibrary();updateSelection();
    }catch(error){toast(error.message,true);}
  };
  const compileSelected=()=>{
    const ids=[...state.selected];
    if(!ids.length){toast('Seleccioná uno o varios archivos.',true);return;}
    if(ids.length>10){toast('Elegí como máximo 10 archivos para el compilado.',true);return;}
    const total=state.library.filter(x=>state.selected.has(x.id)).reduce((n,x)=>n+x.count,0);
    if(total>60){toast('Seleccioná menos archivos: máximo 60 análisis por compilado.',true);return;}
    window.location.href=`/api/library/compile.pptx?ids=${ids.join(',')}`;
    toast(`Generando compilado de ${ids.length} archivo(s).`);
  };
  const loadHistory=async()=>{
    try{
      const response=await fetch('/api/history');if(!response.ok)throw Error('No se pudo cargar el historial.');
      const data=await response.json();setText('historyCount',String(data.items.length));$('downloadAllPpt').classList.toggle('is-hidden',!data.items.length);const area=$('historyContainer');area.replaceChildren();
      if(!data.items.length){const node=el('div','empty-state');node.append(el('span',null,'◷'),el('strong',null,'Todavía no hay análisis'),el('p',null,'Tu primer resultado aparecerá aquí.'));area.append(node);return;}
      data.items.forEach(item=>{
        const btn=el('button','history-row');btn.type='button';
        const icon=el('span','history-icon'+(item.verdict==='FUERA_DE_RANGO'?' bad':''),item.verdict==='CUMPLE'?'✓':item.verdict==='FUERA_DE_RANGO'?'!':'⌁');
        const content=el('span','history-content');content.append(el('strong',null,item.filename),el('small',null,`${item.metric} · ${item.sheet}`));
        const meta=el('span','history-meta');meta.append(el('b',null,`${nf(item.total)} lecturas · ${nf(item.outside)} fuera`),document.createTextNode(new Date(item.created_at).toLocaleString('es-NI',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})));
        btn.append(icon,content,meta,el('span','history-arrow','↗'));btn.addEventListener('click',async()=>{
          btn.disabled=true;try{const response=await fetch(`/api/analysis/${item.id}`);if(!response.ok)throw Error('No se pudo abrir el análisis.');const data=await response.json();showResult(data.result,data.id,data.filename);$('resultados').scrollIntoView({behavior:'smooth'});}catch(e){toast(e.message,true);}finally{btn.disabled=false;}
        });area.append(btn);
      });
    }catch(e){toast(e.message,true);}
  };
  document.addEventListener('DOMContentLoaded',()=>{
    setText('copyrightYear',String(new Date().getFullYear()));setText('todayDate',new Date().toLocaleDateString('es-NI',{day:'numeric',month:'long',year:'numeric'}));
    const drop=$('dropZone'),input=$('fileInput');
    drop.addEventListener('click',()=>input.click());drop.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();input.click();}});
    drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('dragging');});drop.addEventListener('dragleave',()=>drop.classList.remove('dragging'));
    drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('dragging');if(e.dataTransfer.files.length)updateFile(e.dataTransfer.files[0]);});
    input.addEventListener('change',()=>{if(input.files.length)updateFile(input.files[0]);});
    $('clearFile').addEventListener('click',()=>updateFile(null));
    $('sheetSelect').addEventListener('change',inspectFile);
    $('runAnalysis').addEventListener('click',run);
    $('october2026').addEventListener('click',()=>{$('dateFrom').value='2026-10-01';$('dateTo').value='2026-10-31';toast('Filtro listo: 1 al 31 de octubre de 2026.');});
    $('clearPeriod').addEventListener('click',()=>{$('dateFrom').value='';$('dateTo').value='';toast('Se analizará el archivo completo.');});
    $('runBatch').addEventListener('click',runBatch);
    $('rowFilter').addEventListener('change',()=>{state.tableLimit=60;renderRows();});
    $('loadMore').addEventListener('click',()=>{state.tableLimit+=120;renderRows();});
    $('refreshHistory').addEventListener('click',loadHistory);
    $('refreshLibrary').addEventListener('click',loadLibrary);
    $('librarySearch').addEventListener('input',e=>{state.libraryFilter=e.target.value;renderLibrary();});
    $('compilePpt').addEventListener('click',compileSelected);
    $('selectLibrary').addEventListener('click',()=>{
      const visible=state.library.filter(f=>f.filename.toLocaleLowerCase('es').includes(state.libraryFilter.toLocaleLowerCase('es')));
      const ids=visible.map(f=>f.id);
      const allChecked=ids.length&&ids.every(id=>state.selected.has(id));
      ids.forEach(id=>{if(allChecked)state.selected.delete(id);else if(state.selected.size<10)state.selected.add(id);});
      renderLibrary();updateSelection();
    });
    document.querySelectorAll('.nav-link').forEach(link=>link.addEventListener('click',()=>{document.querySelectorAll('.nav-link').forEach(a=>a.classList.remove('is-active'));link.classList.add('is-active');}));
    loadHistory();loadLibrary();
  });
})();
