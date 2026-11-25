/* Simulation Lab v0.1 — in-workbook grid exploration. No external calculation service. */
'use strict';
const state={x:[{id:1,cell:'',title:'',start:'0',end:'20',step:'2'}],y:[{id:2,cell:'',title:''}],z:[],next:3,running:false,cancel:false};
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=n=>Number.isSafeInteger(n)?n.toLocaleString('en-US'):'Too large';
function points(v){const a=Number(v.start),b=Number(v.end),s=Number(v.step);if(!Number.isFinite(a)||!Number.isFinite(b)||!Number.isFinite(s)||s<=0||b<a)return 0;return Math.max(0,Math.floor((b-a)/s+1e-10)+1);}
function metrics(){const factors=state.x.map(points),n=factors.reduce((p,v)=>p*v,1);return {factors,n};}
function countRender(){const {factors,n}=metrics();$('total').textContent=fmt(n);$('factor').textContent=factors.map(fmt).join(' × ')+' = '+fmt(n);$('gridWarning').textContent=n>5000?'Large experiment: recalculation may be slow.':n===0?'Correct the start/end/step values.':'';$('gridWarning').className=n>5000||n===0?'warning':'';}
function itemHtml(kind,v,i){const isX=kind==='x',nm=kind.toUpperCase()+(i+1),title=v.title||v.cell||nm;
 return `<div class="item" data-kind="${kind}" data-id="${v.id}"><div class="item-head"><span>${esc(nm)} · ${esc(title)}</span><button class="remove" data-remove="${kind}:${v.id}" title="Remove">✕</button></div><div class="fields"><div class="field"><label>Excel cell</label><div class="cellrow"><input data-field="cell" placeholder="Sheet1!B1" value="${esc(v.cell)}"><button data-pick="${kind}:${v.id}" title="Use the currently selected Excel cell">Use selection</button></div></div><div class="field"><label>Display title (optional)</label><input data-field="title" placeholder="Default: cell address" value="${esc(v.title)}"></div></div>${isX?`<div class="grid3"><div class="field"><label>Start</label><input data-field="start" type="number" step="any" value="${esc(v.start)}"></div><div class="field"><label>End</label><input data-field="end" type="number" step="any" value="${esc(v.end)}"></div><div class="field"><label>Step</label><input data-field="step" type="number" step="any" value="${esc(v.step)}"></div></div><div class="item-foot">Grid: <b>${fmt(points(v))}</b> values</div>`:''}</div>`;}
function render(){for(const kind of ['x','y','z'])$(kind+'List').innerHTML=state[kind].map((v,i)=>itemHtml(kind,v,i)).join('');countRender();}
function add(kind){state[kind].push(kind==='x'?{id:state.next++,cell:'',title:'',start:'0',end:'10',step:'1'}:{id:state.next++,cell:'',title:''});render();}
function message(str,type=''){$('message').className='message '+type;$('message').textContent=str;}
function progress(done,total,feasible){$('progressBar').style.width=(total?100*done/total:0)+'%';$('progressText').textContent=`${fmt(done)} / ${fmt(total)} evaluated · ${fmt(feasible)} feasible`;}
function setRunning(yes){state.running=yes;$('runBtn').disabled=yes;$('cancelBtn').disabled=!yes;for(const b of document.querySelectorAll('[data-remove],[data-pick],#addX,#addY,#addZ'))b.disabled=yes;}
// An Excel cell reference can include a quoted worksheet name, e.g. 'Sales 2026'!$B$4.
function parseCell(input,defaultSheet){let str=input.trim();if(!str)throw new Error('A required cell reference is blank.');let sheet=defaultSheet,address=str;let m=str.match(/^'((?:[^']|'')+)'!\$?([A-Z]{1,3})\$?([1-9][0-9]*)$/i);if(m){sheet=m[1].replace(/''/g,"'");address=m[2]+m[3];}else{m=str.match(/^([^!]+)!\$?([A-Z]{1,3})\$?([1-9][0-9]*)$/i);if(m){sheet=m[1];address=m[2]+m[3];}else{m=str.match(/^\$?([A-Z]{1,3})\$?([1-9][0-9]*)$/i);if(m)address=m[1]+m[2];else throw new Error('Invalid cell reference: '+str);}}
 const row=Number(address.match(/\d+$/)[0]);if(row>1048576)throw new Error('Row is outside Excel limits: '+str);return {sheet,address:address.toUpperCase(),key:sheet.toLowerCase()+'!'+address.toUpperCase()};}
function fullname(sheet,address){return (sheet.includes(' ')||sheet.includes("'"))?"'"+sheet.replace(/'/g,"''")+"'!"+address:sheet+'!'+address;}
function validate(){const {n}=metrics();if(!state.x.length||!state.y.length)throw new Error('Add at least one X variable and one Y observation.');if(n===0||!Number.isSafeInteger(n))throw new Error('Invalid grid size.');if(n>20000)throw new Error('Prototype safety limit: 20,000 evaluations. Reduce the grids.');if(n>1048575)throw new Error('More results than fit on an Excel sheet.');for(const k of ['x','y','z'])for(const v of state[k])if(!v.cell.trim())throw new Error('Every '+k.toUpperCase()+' row needs a cell reference.');return n;}
async function getSelectedCell(){return Excel.run(async ctx=>{const r=ctx.workbook.getSelectedRange();r.load(['address','rowCount','columnCount']);r.worksheet.load('name');await ctx.sync();if(r.rowCount!==1||r.columnCount!==1)throw Error('Select exactly one Excel cell.');return fullname(r.worksheet.name,r.address.split('!').pop().replace(/\$/g,''));});}
function getCellObj(ctx,ref){return ctx.workbook.worksheets.getItem(ref.sheet).getRange(ref.address);}
function mixedIndex(idx,factors){const out=Array(factors.length);for(let j=factors.length-1;j>=0;j--){out[j]=idx%factors[j];idx=Math.floor(idx/factors[j]);}return out;}
function feasibleValue(value){return value===true;}
function pauseUI(){return new Promise(resolve=>setTimeout(resolve,0));}
async function run(){if(state.running)return;let total;try{total=validate();}catch(e){message(e.message,'error');return;}
 if(!window.Excel||!window.Office){message('Open this add-in inside Excel. Standalone browser preview cannot run experiments.','error');return;}
 state.cancel=false;setRunning(true);message('Preparing workbook…');progress(0,total,0);
 let done=0,ok=0,skipped=0,results=[],error=null,outputName='';
 try{
 await Excel.run(async ctx=>{
   const book=ctx.workbook;const app=ctx.application;
   const sheet=book.worksheets.getActiveWorksheet();sheet.load('name');app.load('calculationMode');await ctx.sync();
   const defaultSheet=sheet.name,originalMode=app.calculationMode;
   const xr=state.x.map(v=>({v,ref:parseCell(v.cell,defaultSheet)}));
   const yr=state.y.map(v=>({v,ref:parseCell(v.cell,defaultSheet)}));
   const zr=state.z.map(v=>({v,ref:parseCell(v.cell,defaultSheet)}));
   const all=[...xr,...yr,...zr];const uniqueX=new Set();
   for(const x of xr){if(uniqueX.has(x.ref.key))throw Error('Duplicate X input cell: '+x.v.cell);uniqueX.add(x.ref.key);}
   for(const x of xr){if([...yr,...zr].some(v=>v.ref.key===x.ref.key))throw Error('Input cell cannot also be Y or Z: '+x.v.cell);}
   const xRanges=xr.map(o=>getCellObj(ctx,o.ref));const yRanges=yr.map(o=>getCellObj(ctx,o.ref));const zRanges=zr.map(o=>getCellObj(ctx,o.ref));
   for(const r of xRanges)r.load(['values','formulas']);await ctx.sync();
   const originals=xRanges.map(r=>({value:r.values[0][0],formula:r.formulas[0][0]}));
   const factors=xr.map(o=>points(o.v));
   try{
     // Manual calculation avoids premature recalculation on every assignment.
     app.calculationMode='Manual';await ctx.sync();
     for(let i=0;i<total;i++){
       if(state.cancel)break;
       const digits=mixedIndex(i,factors);
       const values=xr.map((o,j)=>{const v=Number(o.v.start)+digits[j]*Number(o.v.step);return Number(v.toPrecision(14));});
       xRanges.forEach((range,j)=>{range.values=[[values[j]]];});
       await ctx.sync();
       app.calculate($('fullCalc').checked?'Full':'Recalculate');
       await ctx.sync();
       for(const r of [...zRanges,...yRanges])r.load('values');
       await ctx.sync();
       const checks=zRanges.map(r=>r.values[0][0]);
       if(checks.every(feasibleValue)){
         const outcomes=yRanges.map(r=>r.values[0][0]);
         results.push([...values,...outcomes]);ok++;
       }else skipped++;
       done=i+1;
       if(done===1||done%10===0||done===total){progress(done,total,ok);await pauseUI();}
     }
   }finally{
     // Restore original X-cell contents (formulas or constants) and calculation mode.
     xRanges.forEach((range,j)=>{
       const o=originals[j];
       if(typeof o.formula==='string' && o.formula.startsWith('=')){
         range.formulas=[[o.formula]];
       }else{
         range.values=[[o.value]];
       }
     });
     app.calculationMode=originalMode;
     await ctx.sync();
   }
   // Write results only after the model has been restored.
   const out=book.worksheets.add('Lab '+new Date().toISOString().replace(/[-:TZ.]/g,'').slice(0,12));out.load('name');await ctx.sync();outputName=out.name;
   const headers=[...xr.map(o=>o.v.title.trim()||fullname(o.ref.sheet,o.ref.address)),...yr.map(o=>o.v.title.trim()||fullname(o.ref.sheet,o.ref.address))];
   const ncols=headers.length,colsX=xr.length;
   const head=out.getRangeByIndexes(0,0,1,ncols);head.values=[headers];head.format.font.bold=true;head.format.rowHeight=30;
   if(colsX){const h=out.getRangeByIndexes(0,0,1,colsX);h.format.fill.color='#D9EAFB';}
   if(yr.length){const h=out.getRangeByIndexes(0,colsX,1,yr.length);h.format.fill.color='#FFF0B8';}
   for(let k=0;k<results.length;k+=500){const chunk=results.slice(k,k+500);const block=out.getRangeByIndexes(k+1,0,chunk.length,ncols);block.values=chunk;if(colsX)out.getRangeByIndexes(k+1,0,chunk.length,colsX).format.fill.color='#EFF6FF';if(yr.length)out.getRangeByIndexes(k+1,colsX,chunk.length,yr.length).format.fill.color='#FFFBEB';await ctx.sync();}
   out.getRangeByIndexes(0,0,Math.max(1,results.length+1),ncols).format.autofitColumns();
   out.getRange('A1').select();await ctx.sync();
 });
 progress(done,total,ok);
 message(`${state.cancel?'Cancelled':'Completed'}: ${fmt(done)} evaluations, ${fmt(ok)} feasible, ${fmt(skipped)} infeasible.\nResults saved in worksheet: ${outputName}.`,'ok');
 }catch(e){error=e;message('Experiment failed: '+(e&&e.message?e.message:String(e))+'\nThe add-in attempted to restore input cells and calculation mode. Verify the workbook before saving.','error');}
 finally{setRunning(false);state.cancel=false;}
}
document.addEventListener('input',event=>{const input=event.target.closest('[data-field]');if(!input||state.running)return;const node=input.closest('[data-kind]');const v=state[node.dataset.kind].find(a=>a.id===Number(node.dataset.id));if(!v)return;v[input.dataset.field]=input.value;if(input.dataset.field==='title'){node.querySelector('.item-head span').textContent=node.dataset.kind.toUpperCase()+(state[node.dataset.kind].indexOf(v)+1)+' · '+(v.title||v.cell||'');}if(['start','end','step'].includes(input.dataset.field)){node.querySelector('.item-foot b').textContent=fmt(points(v));countRender();}});
document.addEventListener('click',async event=>{if(state.running)return;const remove=event.target.closest('[data-remove]');if(remove){const [kind,id]=remove.dataset.remove.split(':');state[kind]=state[kind].filter(v=>v.id!==Number(id));render();return;}const pick=event.target.closest('[data-pick]');if(pick){const [kind,id]=pick.dataset.pick.split(':');try{const cell=await getSelectedCell();const v=state[kind].find(v=>v.id===Number(id));if(v){v.cell=cell;render();message('Selected '+cell);}}catch(e){message(e.message,'error');}}});
$('addX').onclick=()=>add('x');$('addY').onclick=()=>add('y');$('addZ').onclick=()=>add('z');$('runBtn').onclick=run;$('cancelBtn').onclick=()=>{state.cancel=true;message('Cancellation requested; finishing the current recalculation…');};
render();
if(window.Office){Office.onReady(info=>{if(info.host===Office.HostType.Excel){$('officeStatus').textContent='Connected to Microsoft Excel';$('officeStatus').style.background='#def6eb';$('officeStatus').style.color='#166c51';}else{$('officeStatus').textContent='Open this tool inside Excel';}}).catch(e=>message(e.message,'error'));}
else $('officeStatus').textContent='Design preview — Excel connection unavailable';
