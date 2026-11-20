// Final saved-artifact visual check. Leaves the delivered XLSX unchanged.
import fs from 'node:fs/promises';import path from 'node:path';import {fileURLToPath,pathToFileURL} from 'node:url';import {createRequire} from 'node:module';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');if(path.resolve(process.cwd())!==root||path.basename(root)!=='NickelOps-Ore-Delivery-Reconciliation-and-Scenario-Analysis')throw Error('Boundary violation');
const local=p=>{const d=path.resolve(root,p);if(!d.startsWith(root+path.sep))throw Error('Boundary violation');return d;};
process.env.TEMP=process.env.TMP=local('tmp');const dep=process.env.NICKELOPS_NODE_MODULES||path.join(process.env.USERPROFILE,'.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules');
const req=createRequire(path.join(dep,'loader.cjs'));const {Workbook,SpreadsheetFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(local('NickelOps.xlsx')));const s=w.worksheets.getItem('DASHBOARD');
if(process.argv.includes('--review-only')){
 const review=w.worksheets.getItem('DATA REVIEW');review.getRange('A2:A3').values=[['Review'],['SYNTHETIC DATA']];
 const preview=await w.render({sheetName:'DATA REVIEW',range:'A1:J19',scale:1,format:'png'});await fs.writeFile(local('previews/data_review.png'),new Uint8Array(await preview.arrayBuffer()));console.log('Final review pane rendered.');process.exit(0);
}
// Explicit zero origin for both value axes, matching the native saved chart metadata.
for(const chart of s.charts.items)chart.yAxis={min:0};
const serial=v=>typeof v==='number'?v:((v instanceof Date?v.getTime():Date.parse(v))-Date.UTC(1899,11,30))/86400000;
const start=serial(s.getRange('B6').values[0][0]),end=serial(s.getRange('D6').values[0][0]);const clean=w.worksheets.getItem('CLEAN DELIVERIES').getRange('A7:T316').values;
const tests=[];
for(const c of [{name:'dashboard',supplier:'All',pile:'All',end},{name:'dashboard_date_filtered',supplier:'All',pile:'All',end:start+29},{name:'dashboard_filtered',supplier:'S02',pile:'P01',end}]){
 s.getRange('F6').values=[[c.supplier]];s.getRange('H6').values=[[c.pile]];s.getRange('D6').values=[[c.end]];await w.recalculate();
 const expected=clean.filter(r=>r[7]==='Accepted'&&serial(r[2])>=start&&serial(r[2])<=c.end&&(c.supplier==='All'||r[3]===c.supplier)&&(c.pile==='All'||r[4]===c.pile));
 for(const [name,cell,actualExpected] of [['count','D11',expected.length],['WMT','D12',expected.reduce((a,r)=>a+r[8],0)]]){
  const actual=s.getRange(cell).values[0][0];const passed=typeof actual==='number'&&Math.abs(actual-actualExpected)<=1e-7;tests.push({name:'Saved XLSX import '+c.name+' '+name,expected:actualExpected,actual,tolerance:1e-7,result:passed?'PASS':'FAIL',engine:'@oai/artifact-tool XLSX import/recalculation'});if(!passed)throw Error('Saved artifact scope mismatch '+c.name+' '+name);
 }
 const png=await w.render({sheetName:'DASHBOARD',range:'A1:M43',scale:1,format:'png'});await fs.writeFile(local(`previews/${c.name}.png`),new Uint8Array(await png.arrayBuffer()));
 console.log('Final saved dashboard rendered: '+c.name);
}
await fs.writeFile(local('reports/saved_artifact_tests.json'),JSON.stringify({tests},null,2));
const qaPreview=await w.render({sheetName:'QA CHECKS',range:'A1:F25',scale:1,format:'png'});await fs.writeFile(local('previews/qa_checks.png'),new Uint8Array(await qaPreview.arrayBuffer()));
const review=w.worksheets.getItem('DATA REVIEW');review.getRange('A2:A3').values=[['Review'],['SYNTHETIC DATA']];
const reviewPreview=await w.render({sheetName:'DATA REVIEW',range:'A1:J19',scale:1,format:'png'});await fs.writeFile(local('previews/data_review.png'),new Uint8Array(await reviewPreview.arrayBuffer()));
