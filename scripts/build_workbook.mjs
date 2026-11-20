// Author with the installed @oai/artifact-tool runtime. No external services.
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {pathToFileURL, fileURLToPath} from 'node:url';

const ROOT=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
if(path.basename(ROOT)!=='NickelOps-Ore-Delivery-Reconciliation-and-Scenario-Analysis'||path.resolve(process.cwd())!==ROOT) throw Error('Run from authorized project root');
function local(relative){const p=path.resolve(ROOT,relative);if(!p.startsWith(ROOT+path.sep))throw Error('Boundary violation');return p;}
process.env.TEMP=process.env.TMP=local('tmp');
const dep=process.env.NICKELOPS_NODE_MODULES||path.join(process.env.USERPROFILE,'.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules');
const req=createRequire(path.join(dep,'loader.cjs'));
const {Workbook,SpreadsheetFile}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const spec=JSON.parse(await fs.readFile(local('tmp/workbook_spec.json'),'utf8'));
const workbook=Workbook.create();
for(const s of spec.sheets)workbook.worksheets.add(s.name);
const sheets=Object.fromEntries(spec.sheets.map(s=>[s.name,workbook.worksheets.getItem(s.name)]));
const navy='#243C55',ink='#23354B',amber='#FFF0BD',red='#FCE3DC';
function endCell(anchor,rows){const match=anchor.match(/^([A-Z]+)(\d+)$/);let c=0;for(const ch of match[1])c=c*26+ch.charCodeAt(0)-64;let end=c+rows[0].length-1,letters='';while(end){end--;letters=String.fromCharCode(65+end%26)+letters;end=Math.floor(end/26);}return `${letters}${Number(match[2])+rows.length-1}`;}
function coalesce(blocks){const result=[];for(const b of blocks){const last=result.at(-1),m=b.cell.match(/^([A-Z]+)(\d+)$/),p=last?.cell.match(/^([A-Z]+)(\d+)$/);if(last&&p[1]===m[1]&&Number(p[2])+last.data.length===Number(m[2])&&last.data[0].length===b.data[0].length)last.data.push(...b.data);else result.push({cell:b.cell,data:b.data.map(r=>[...r])});}return result;}
for(const def of spec.sheets){
  const s=sheets[def.name];s.showGridLines=false;
  for(const block of coalesce(def.values))s.getRange(`${block.cell}:${endCell(block.cell,block.data)}`).values=block.data.map(row=>row.map(v=>v===''?null:v));
  for(const block of coalesce(def.formulas))s.getRange(`${block.cell}:${endCell(block.cell,block.data)}`).formulas=block.data;
  const used=s.getUsedRange();used.format.font={name:'Arial',size:10,color:ink};used.format.rowHeight=21;used.format.verticalAlignment='center';
  for(const [c,width] of Object.entries(def.widths))s.getRange(`${c}:${c}`).format.columnWidth=width;
  s.getRange('A2:K2').format.font={name:'Arial',size:16,bold:true,color:navy};s.getRange('A2:K2').format.rowHeight=29;
  s.getRange('A3:K3').format.font={name:'Arial',size:10,italic:true,color:'#5C6674'};
  if(def.name!=='START HERE')s.getRange('A4').formulas=[['=HYPERLINK("#\'START HERE\'!B6","Back to START HERE")']];
  for(const tab of def.tables){const t=s.tables.add(tab.range,true,tab.name);t.style='TableStyleMedium2';const header=t.getHeaderRowRange();header.format.fill=navy;header.format.font={name:'Arial',size:10,bold:true,color:'#FFFFFF'};header.format.rowHeight=40;header.format.wrapText=true;header.format.horizontalAlignment='center';}
  for(const f of def.formats)s.getRange(f.range).setNumberFormat(f.format);
  for(const r of def.inputs){s.getRange(r).format.fill=amber;s.getRange(r).format.font.color='#6C4400';}
  for(const warning of def.warnings)s.getRange(warning.range).conditionalFormats.addCustom(warning.formula,{fill:red,font:{color:'#9C3025'}});
  if(def.freezeRows)s.freezePanes.freezeRows(def.freezeRows);
  if(def.freezeColumns)s.freezePanes.freezeColumns(def.freezeColumns);
  if(['DASHBOARD','SUMMARY'].includes(def.name))s.tabColor=navy;
  if(def.name==='SCENARIOS')s.tabColor='#627C95';
  console.log('Authored '+def.name);
}
for(const v of spec.validations)sheets[v.sheet].getRange(v.range).dataValidation={rule:v.rule};
const home=sheets['START HERE'];
for(let i=1;i<spec.sheets.length;i++)home.getRange(`B${11+i}`).formulas=[[`=HYPERLINK("#'${spec.sheets[i].name}'!A2","${spec.sheets[i].name}")`]];
home.getRange('B6:H6').format.font={name:'Arial',size:14,bold:true,color:navy};home.getRange('B22:H39').format.rowHeight=25;
for(const def of spec.sheets)for(const ch of def.charts){
 const s=sheets[def.name];const ranges=ch.ranges.map(x=>s.getRange(x));const chart=s.charts.add(ch.type,ranges.length===1?ranges[0]:ranges);
 chart.title=ch.title;chart.titleTextStyle.typeface='Arial';chart.titleTextStyle.fontSize=13;chart.setPosition(...ch.position);chart.hasLegend=false;
 chart.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:10}};
 chart.yAxis={min:0,numberFormatCode:ch.yFormat,numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:10}};
 if(ch.type==='line')chart.series.items[0].line={fill:'#356B9A',style:'solid',width:2};else chart.series.items[0].fill='#356B9A';
}
const dashboard=sheets.DASHBOARD;
for(const r of ['A10:D10','A21:D21','A33:D33']){dashboard.getRange(r).format.fill=navy;dashboard.getRange(r).format.font={name:'Arial',size:10,bold:true,color:'#FFFFFF'};}
dashboard.getRange('A11:C17').format.rowHeight=24;dashboard.getRange('D11:D17').format.font={name:'Arial',size:12,bold:true,color:navy};
// Ordinary headers without native Table objects.
for(const [sheetName,ranges] of Object.entries({'SCENARIOS':['A11:D11','A18:D18','G18:J18'],'DASHBOARD':['A22:D22']}))for(const r of ranges){const x=sheets[sheetName].getRange(r);x.format.fill=navy;x.format.font={name:'Arial',size:10,bold:true,color:'#FFFFFF'};x.format.wrapText=true;x.format.rowHeight=34;}
await workbook.recalculate();
function value(sheet,cell){return sheets[sheet].getRange(cell).values[0][0];}
function write(sheet,cell,v){sheets[sheet].getRange(cell).values=[[v]];}
const tests=[];
function test(name,expected,actual,tol=1e-7){const passed=typeof expected==='number'?typeof actual==='number'&&Math.abs(expected-actual)<=tol:JSON.stringify(expected)===JSON.stringify(actual);tests.push({name,expected,actual,tolerance:tol,result:passed?'PASS':'FAIL',engine:'@oai/artifact-tool'});if(!passed)console.log('FAIL '+name+' '+JSON.stringify({expected,actual}));}
function scenVals(){return {wmt:value('SCENARIOS','C12'),dmt:value('SCENARIOS','C13'),ni:value('SCENARIOS','C14')};}
const acceptedSnapshot=sheets['CLEAN DELIVERIES'].getRange(`A7:T${spec.end}`).values;
const reviewSnapshot=sheets['DATA REVIEW'].getRange(`A7:AK${spec.end}`).values;
const baseline={wmt:value('SCENARIOS','B12'),dmt:value('SCENARIOS','B13'),ni:value('SCENARIOS','B14')};
const defaultScenario=scenVals();
for(const [metric,k] of [['wmt','wmt'],['dmt','dmt'],['ni','contained_ni_t']])test('Initial Python control '+metric,spec.initial_controls[k],baseline[metric]);
function assumptions(v,m,g){write('SCENARIOS','B6',v);write('SCENARIOS','E6',m);write('SCENARIOS','H6',g);}
assumptions(0,0,0);await workbook.recalculate();for(const k of Object.keys(baseline))test('Zero-change scenario '+k,baseline[k],scenVals()[k]);
assumptions(0,2,0);await workbook.recalculate();let moistureOnly=scenVals();test('Higher moisture reduces DMT',true,moistureOnly.dmt<baseline.dmt);test('Higher moisture reduces Ni',true,moistureOnly.ni<baseline.ni);test('Higher moisture leaves WMT',baseline.wmt,moistureOnly.wmt);
assumptions(0,0,.1);await workbook.recalculate();let gradeOnly=scenVals();test('Higher grade leaves WMT',baseline.wmt,gradeOnly.wmt);test('Higher grade leaves DMT',baseline.dmt,gradeOnly.dmt);test('Grade pp uses division by 100 once',baseline.ni+baseline.dmt*.001,gradeOnly.ni);
assumptions(.1,0,0);await workbook.recalculate();for(const k of Object.keys(baseline))test('10% volume scales '+k,baseline[k]*1.1,scenVals()[k]);
assumptions(-1,0,0);await workbook.recalculate();test('Minus 100% volume valid',0,scenVals().wmt);test('Zero DMT grade unavailable','Unavailable',value('SCENARIOS','C15'));
assumptions(-1.01,0,0);await workbook.recalculate();test('Volume below -100% blocked',true,String(value('SCENARIOS','B9')).startsWith('BLOCKED'));test('Invalid scenario total unavailable','Unavailable',value('SCENARIOS','C14'));
assumptions(0,100,0);await workbook.recalculate();test('Out-of-range moisture blocked',true,String(value('SCENARIOS','B9')).startsWith('BLOCKED'));
assumptions(0,0,-10);await workbook.recalculate();test('Out-of-range grade blocked',true,String(value('SCENARIOS','B9')).startsWith('BLOCKED'));
write('SCENARIOS','B6','');await workbook.recalculate();test('Blank assumption blocked',true,String(value('SCENARIOS','B9')).startsWith('BLOCKED'));
assumptions(.1,2,.1);await workbook.recalculate();
// Filter through the authoritative dashboard controls, then inspect the affected chart source.
write('DASHBOARD','F6','S02');write('DASHBOARD','H6','P01');await workbook.recalculate();
const filtered={count:value('DASHBOARD','D11'),wmt:value('DASHBOARD','D12'),dmt:value('DASHBOARD','D13'),ni:value('DASHBOARD','D15')};
const independentlyFiltered=acceptedSnapshot.filter(r=>r[7]==='Accepted'&&r[3]==='S02'&&r[4]==='P01');
test('Supplier/stockpile selector count',independentlyFiltered.length,filtered.count,0);test('Supplier/stockpile selector WMT',independentlyFiltered.reduce((a,r)=>a+r[8],0),filtered.wmt);
test('Filtered chart source reconciles',filtered.wmt,sheets.DASHBOARD.getRange('B50:B109').values.reduce((a,r)=>a+(typeof r[0]==='number'?r[0]:0),0));
test('Filtered Ni chart source updates',filtered.ni,value('DASHBOARD','K50'));
const filteredPreview=await workbook.render({sheetName:'DASHBOARD',range:'A1:M43',scale:1,format:'png'});await fs.writeFile(local('previews/dashboard_filtered.png'),new Uint8Array(await filteredPreview.arrayBuffer()));
write('DASHBOARD','F6','All');write('DASHBOARD','H6','All');write('DASHBOARD','B6',46300);write('DASHBOARD','D6',46200);await workbook.recalculate();test('Reversed dates blocked',true,String(value('DASHBOARD','B8')).startsWith('BLOCKED'));test('Invalid scope KPI unavailable','Unavailable',value('DASHBOARD','D12'));
write('DASHBOARD','B6',46266);write('DASHBOARD','D6',46325); // replace below with original typed source dates to avoid assumed serials
const origDash=spec.sheets.find(s=>s.name==='DASHBOARD');for(const b of origDash.values)if(['B6','D6'].includes(b.cell))dashboard.getRange(b.cell).values=b.data;
await workbook.recalculate();
// Keyed correction: missing date becomes accepted, neighboring row must remain unchanged.
const neighborBefore=reviewSnapshot.find(r=>r[0]==='R000032')[1];
write('DATA REVIEW','AQ11','R000031');write('DATA REVIEW','AR11','Approve');write('DATA REVIEW','AS11','Engine test: synthetic confirmed date.');write('DATA REVIEW','AT11',46266);
// Use actual source period serial, not a runtime date.
write('DATA REVIEW','AT11',origDash.values.find(b=>b.cell==='B6').data[0][0]+6);
await workbook.recalculate();
const reviewNow=()=>sheets['DATA REVIEW'].getRange(`A7:B${spec.end}`).values;
test('Keyed date correction accepts intended record','Accepted',reviewNow().find(r=>r[0]==='R000031')[1]);test('Keyed correction increases count by one',spec.initial_controls.count+1,value('DASHBOARD','D11'),0);test('Neighboring row unchanged',neighborBefore,reviewNow().find(r=>r[0]==='R000032')[1]);
write('DATA REVIEW','AS11','');await workbook.recalculate();test('Correction without reason blocked','Unresolved',reviewNow().find(r=>r[0]==='R000031')[1]);write('DATA REVIEW','AS11','Engine test: reason restored.');write('DATA REVIEW','AQ12','R000031');await workbook.recalculate();test('Duplicate override keys block record','Unresolved',reviewNow().find(r=>r[0]==='R000031')[1]);
sheets['DATA REVIEW'].getRange('AQ11:BA12').clear({applyTo:'contents'});await workbook.recalculate();
write('DATA REVIEW','AS9','');await workbook.recalculate();test('Undocumented rejection keeps competing record blocked','Unresolved',reviewNow().find(r=>r[0]==='R000081')[1]);write('DATA REVIEW','AS9',spec.sheets.find(s=>s.name==='DATA REVIEW').values.find(b=>b.cell==='AQ7').data[2][2]);await workbook.recalculate();
const first=sheets['DATA REVIEW'].getRange('AQ7:BA7').values,second=sheets['DATA REVIEW'].getRange('AQ8:BA8').values;
sheets['DATA REVIEW'].getRange('AQ7:BA7').values=second;sheets['DATA REVIEW'].getRange('AQ8:BA8').values=first;await workbook.recalculate();test('Reordered correction register retains accepted count',spec.initial_controls.count,value('DASHBOARD','D11'),0);test('Reordered correction register retains DMT',baseline.dmt,value('DASHBOARD','D13'));
sheets['DATA REVIEW'].getRange('AQ7:BA7').values=first;sheets['DATA REVIEW'].getRange('AQ8:BA8').values=second;
write('DATA REVIEW','AQ11','R000001');write('DATA REVIEW','AR11','Hold');await workbook.recalculate();test('Keyed Hold decreases accepted population',spec.initial_controls.count-1,value('DASHBOARD','D11'),0);sheets['DATA REVIEW'].getRange('AQ11:BA11').clear({applyTo:'contents'});
await workbook.recalculate();test('All inputs restored',baseline.wmt,value('DASHBOARD','D12'));
const errorScan=await workbook.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:100},maxChars:7000,summary:'Final unexpected error scan'});
await fs.writeFile(local('reports/formula_error_scan.ndjson'),errorScan.ndjson);
const snapshot={engine:'@oai/artifact-tool',baseline,defaultScenario,filtered,moistureOnly,gradeOnly,tests,review:reviewSnapshot,clean:acceptedSnapshot,qa:sheets['QA CHECKS'].getRange('A7:F19').values,summary:sheets.SUMMARY.getRange('A6:H83').values};
await fs.writeFile(local('reports/engine_tests.json'),JSON.stringify(snapshot,null,2));
const previewRanges={'START HERE':'A1:H39','RAW DELIVERIES':'A1:J18','REFERENCES':'A1:H29','DATA REVIEW':'A1:J19','CLEAN DELIVERIES':'A1:M19','SUMMARY':'A1:H18','SCENARIOS':'A1:J27','DASHBOARD':'A1:M43','QA CHECKS':'A1:F25'};
for(const [name,range] of Object.entries(previewRanges)){try{const png=await workbook.render({sheetName:name,range,scale:1,format:'png'});await fs.writeFile(local(`previews/${name.toLowerCase().replaceAll(' ','_')}.png`),new Uint8Array(await png.arrayBuffer()));console.log('Rendered '+name);}catch(e){console.log('Render gap '+name+': '+e.message);}}
const ovpreview=await workbook.render({sheetName:'DATA REVIEW',range:'AQ2:BA12',scale:1,format:'png'});await fs.writeFile(local('previews/keyed_review_inputs.png'),new Uint8Array(await ovpreview.arrayBuffer()));
const output=await SpreadsheetFile.exportXlsx(workbook);await output.save(local('NickelOps.xlsx'));
console.log(JSON.stringify({tests:tests.length,failures:tests.filter(t=>t.result==='FAIL').length,output:'NickelOps.xlsx'}));
