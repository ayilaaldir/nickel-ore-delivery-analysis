"""Verify the delivered XLSX, independent issue manifest and Python calculations; export snapshots."""
import csv
import json
import math
import platform
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, date, timedelta
from project_paths import local
from analysis_model import analyze, aggregate, scenario, csv_rows, number

TOL=1e-7
TESTS=[]
def check(name, expected, actual, tolerance=0, engine='Python 3 stdlib / openpyxl read-only'):
    if isinstance(expected,(int,float)) and not isinstance(expected,bool):
        ok=isinstance(actual,(int,float)) and not isinstance(actual,bool) and math.isfinite(actual) and abs(actual-expected)<=tolerance
    else:ok=expected==actual
    TESTS.append(dict(name=name,expected=expected,actual=actual,tolerance=tolerance,engine=engine,result='PASS' if ok else 'FAIL'))
    return ok

def write_csv(path, rows):
    with local(path).open('w',newline='',encoding='utf-8') as f:
        w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)

def main():
    import openpyxl
    book=local('NickelOps.xlsx')
    wb=openpyxl.load_workbook(book,data_only=True,read_only=False)
    formulas=openpyxl.load_workbook(book,data_only=False,read_only=False)
    raw=csv_rows('data/raw_deliveries.csv');r=analyze();a=[x for x in r if x['status']=='Accepted'];tot=aggregate(a)
    keyed={x['source_row_id']:x for x in r};counts=Counter(x['status'] for x in r)
    check('Raw rows partition mutually exclusively',len(raw),sum(counts.values()))
    check('Initial raw row count',310,len(raw));check('Ground-truth delivery count',300,len(csv_rows('evaluation/clean_ground_truth.csv')))
    check('Every source_row_id is unique and stable',len(raw),len({x['source_row_id'] for x in raw}))
    check('Accepted business identifiers unique',len(a),len({x['delivery_id'] for x in a}))
    saved_clean={row[0]:list(row) for row in wb['CLEAN DELIVERIES'].iter_rows(min_row=7,max_row=len(raw)+6,values_only=True)}
    saved_review={row[0]:list(row) for row in wb['DATA REVIEW'].iter_rows(min_row=7,max_row=len(raw)+6,values_only=True)}
    mismatch=[];metric_mismatch=[];blanks=[];metadata_mismatch=[]
    for x in r:
        row=saved_clean[x['source_row_id']]
        if row[7]!=x['status']:mismatch.append(x['source_row_id'])
        if x['status']=='Accepted':
            for index,key in [(1,'delivery_id'),(3,'supplier_id'),(4,'stockpile_id')]:
                if row[index]!=x[key]:metadata_mismatch.append(x['source_row_id']+':'+key)
            if not isinstance(row[2],datetime) or row[2].date().isoformat()!=x['delivery_date']:metadata_mismatch.append(x['source_row_id']+':delivery_date')
        for index,key in [(8,'weighbridge_net_wmt'),(9,'dmt'),(12,'contained_ni_t')]:
            actual=row[index];expected=x[key] if x['status']=='Accepted' else None
            if expected is None:
                if actual not in (None,''):blanks.append((x['source_row_id'],key,actual))
            elif not isinstance(actual,(float,int)) or abs(actual-expected)>TOL:metric_mismatch.append((x['source_row_id'],key,expected,actual))
    check('Saved workbook statuses agree for every source ID',[],mismatch)
    check('Accepted identifiers, reference IDs and dates agree with Python',[],metadata_mismatch)
    check('Saved workbook per-record WMT/DMT/Ni agrees with independent Python',[],metric_mismatch,TOL)
    check('All non-accepted measurements remain blank, not zero',[],blanks)
    saved_raw=[]
    for row in wb['RAW DELIVERIES'].iter_rows(min_row=7,max_row=len(raw)+6,values_only=True):
        row=list(row[:10]);row[2]=row[2].date().isoformat() if isinstance(row[2],datetime) else (row[2] or '')
        saved_raw.append(row)
    raw_mismatch=[]
    for src,saved in zip(raw,saved_raw):
        for i,(k,v) in enumerate(src.items()):
            if i>=10:continue
            actual=saved[i] if saved[i] is not None else ''
            if k in ['supplier_reported_wmt','weighbridge_net_wmt','moisture_pct','ni_grade_pct_dry']:
                if v!='' and (not isinstance(actual,(float,int)) or abs(float(v)-actual)>1e-10):raw_mismatch.append(src['source_row_id']+':'+k)
                elif v=='' and actual!='':raw_mismatch.append(src['source_row_id']+':'+k)
            elif v!=actual:raw_mismatch.append(src['source_row_id']+':'+k)
    check('Saved RAW DELIVERIES preserves every original field',[],raw_mismatch)
    for field,cell in [('count','D11'),('wmt','D12'),('dmt','D13'),('weighted_grade','D14'),('contained_ni_t','D15'),('weighted_moisture','D16'),('weight_exceptions','D17')]:
        check('Saved dashboard independent control: '+field,tot[field],wb['DASHBOARD'][cell].value,TOL)
    for field,col in [('wmt','C'),('dmt','D'),('contained_ni_t','E')]:
        for dim,rows in [('supplier',range(7,11)),('stockpile',range(15,18)),('date',range(24,84))]:
            check(f'Saved {dim} summary reconciles {field}',tot[field],sum(wb['SUMMARY'][f'{col}{i}'].value for i in rows),TOL)
    for dimension,field,groups,first in [('supplier','supplier_id',[f'S{i:02}' for i in range(1,5)],7),('stockpile','stockpile_id',[f'P{i:02}' for i in range(1,4)],15)]:
        for index,group in enumerate(groups,first):
            group_totals=aggregate([x for x in a if x[field]==group])
            for metric,c in [('count','B'),('wmt','C'),('dmt','D'),('contained_ni_t','E'),('weighted_grade','F'),('weighted_moisture','G')]:
                check(f'{dimension} {group} independently agrees: {metric}',group_totals[metric],wb['SUMMARY'][f'{c}{index}'].value,TOL)
    check('Visible QA checks all pass',13,sum(wb['QA CHECKS'][f'E{i}'].value=='PASS' for i in range(7,20)))
    manifest=csv_rows('evaluation/issue_manifest.csv');rawkey={x['source_row_id']:x for x in raw}
    signatures=Counter(tuple(x[k] for k in raw[0] if k!='source_row_id') for x in raw)
    business=Counter(x['delivery_id'].strip().upper() for x in raw)
    manifest_mismatch=[]
    for m in manifest:
        x=rawkey[m['source_row_id']];kind=m['injected_issue_type'];p=keyed[m['source_row_id']]
        detect={
          'IDENTIFIER_NORMALIZATION':any(x[c]!=x[c].strip().upper() for c in ['delivery_id','supplier_id','stockpile_id']),
          'UNKNOWN_SUPPLIER':x['supplier_id'].strip().upper() not in ['S01','S02','S03','S04'],
          'MISSING_DATE':x['delivery_date']=='','MISSING_MOISTURE':x['moisture_pct']=='','MISSING_GRADE':x['ni_grade_pct_dry']=='',
          'INVALID_MOISTURE_FRACTION':number(x['moisture_pct']) is not None and not 0<=float(x['moisture_pct'])<=1,
          'INVALID_GRADE_FRACTION':number(x['ni_grade_pct_dry']) is not None and not 0<=float(x['ni_grade_pct_dry'])<=1,
          'NEGATIVE_REPORTED_WEIGHT':float(x['supplier_reported_wmt'])<0,'NEGATIVE_WEIGHBRIDGE_WEIGHT':float(x['weighbridge_net_wmt'])<0,
          'EXACT_RETRANSMISSION':signatures[tuple(x[k] for k in raw[0] if k!='source_row_id')]>1,
          'CONFLICTING_DUPLICATE':business[x['delivery_id'].strip().upper()]>1 and signatures[tuple(x[k] for k in raw[0] if k!='source_row_id')]==1,
          'MATERIAL_WEIGHT_DIFFERENCE':abs((float(x['weighbridge_net_wmt'])-float(x['supplier_reported_wmt']))/float(x['supplier_reported_wmt']))>.03}[kind]
        exp='Accepted' if m['expected_handling'].startswith('Accepted') else ('Rejected' if m['expected_handling'].startswith('Rejected') else m['expected_handling'])
        if not detect or p['status']!=exp or saved_review[m['source_row_id']][1]!=exp:manifest_mismatch.append(m['source_row_id']+':'+kind)
    check('All injected manifest entries detected and handled as expected',[],manifest_mismatch)
    check('Independent manifest entry count',64,len(manifest))
    truth={x['delivery_id']:x for x in csv_rows('evaluation/clean_ground_truth.csv')}
    truth_mismatch=[]
    for x in a:
        original=truth[x['delivery_id']]
        for field in ['delivery_date','supplier_id','stockpile_id']:
            if x[field]!=original[field]:truth_mismatch.append(x['source_row_id']+':'+field)
        for field in ['supplier_reported_wmt','weighbridge_net_wmt','moisture_pct','ni_grade_pct_dry']:
            if abs(x[field]-float(original[field]))>TOL:truth_mismatch.append(x['source_row_id']+':'+field)
    check('Accepted normalized/corrected measurements agree with evaluation truth',[],truth_mismatch,TOL)
    if local('reports/reproducibility.json').exists():
        reproduction=json.loads(local('reports/reproducibility.json').read_text(encoding='utf-8'))
        check('Fixed-seed generation reproduces all data/evaluation bytes',True,reproduction['identical'])
    # Independent scenario mathematics and a hand-computable percentage-point example.
    zero=scenario(a);v=scenario(a,.1);mo=scenario(a,moisture_pp=2);gr=scenario(a,grade_pp=.1);combined=scenario(a,.1,2,.1)
    for key in ['wmt','dmt','contained_ni_t']:
        check('Python zero scenario '+key,tot[key],zero[key],TOL)
        check('Python +10% volume '+key,tot[key]*1.1,v[key],TOL)
    check('Python moisture raises WMT-independent losses',True,mo['dmt']<tot['dmt'] and mo['contained_ni_t']<tot['contained_ni_t'] and mo['wmt']==tot['wmt'])
    check('Python grade leaves WMT and DMT unchanged',True,gr['wmt']==tot['wmt'] and gr['dmt']==tot['dmt'] and gr['contained_ni_t']>tot['contained_ni_t'])
    check('Hand example scenario DMT: 110 x (1 - .32)',74.8,100*1.1*(1-(.30+2/100)),1e-12)
    check('Hand example scenario theoretical Ni: 74.8 x .016',1.1968,100*1.1*(1-(.30+2/100))*(.015+.1/100),1e-12)
    for label,kwargs in [('moisture',dict(moisture_pp=100)),('grade',dict(grade_pp=-10)),('volume',dict(volume=-1.01))]:
        blocked=False
        try:scenario(a,**kwargs)
        except ValueError:blocked=True
        check('Python invalid scenario blocked '+label,True,blocked)
    for key,cell in [('wmt','C12'),('dmt','C13'),('contained_ni_t','C14')]:check('Saved compound scenario '+key,combined[key],wb['SCENARIOS'][cell].value,TOL)
    for row,delta in enumerate([-4,-2,0,2,4],19):
        mo_case=scenario(a,moisture_pp=delta);gr_case=scenario(a,grade_pp=delta/10)
        for metric,c in [('dmt','B'),('contained_ni_t','C')]:check(f'Moisture one-way {delta:+} pp: {metric}',mo_case[metric],wb['SCENARIOS'][f'{c}{row}'].value,TOL)
        for metric,c in [('dmt','H'),('contained_ni_t','I')]:check(f'Grade one-way {delta/10:+} pp: {metric}',gr_case[metric],wb['SCENARIOS'][f'{c}{row}'].value,TOL)
    engine=json.loads(local('reports/engine_tests.json').read_text(encoding='utf-8'))
    scope=aggregate([x for x in a if x['supplier_id']=='S02' and x['stockpile_id']=='P01'])
    for key,field in [('count','count'),('wmt','wmt'),('dmt','dmt'),('ni','contained_ni_t')]:check('Mutated selected scope independently agrees: '+key,scope[field],engine['filtered'][key],TOL)
    TESTS.extend(engine['tests'])
    if local('reports/qa_compatibility_test.json').exists():TESTS.extend(json.loads(local('reports/qa_compatibility_test.json').read_text(encoding='utf-8'))['tests'])
    if 'qa' in engine and any(row[4]!='PASS' for row in engine['qa']):
        engine['qa_before_numeric_blank_check_fix']=engine.pop('qa')
        engine['qa_after_finalization']=[[wb['QA CHECKS'].cell(i,j).value for j in range(1,7)] for i in range(7,20)]
        local('reports/engine_tests.json').write_text(json.dumps(engine,indent=2),encoding='utf-8')
    if local('reports/saved_artifact_tests.json').exists():TESTS.extend(json.loads(local('reports/saved_artifact_tests.json').read_text(encoding='utf-8'))['tests'])
    check('Runtime mutation tests all pass',0,sum(t['result']!='PASS' for t in engine['tests']))
    frows=[];errors=[]
    error_tokens={'#REF!','#DIV/0!','#VALUE!','#NAME?','#N/A','#NUM!','#NULL!','#SPILL!','#CALC!'}
    for sheet in formulas:
        for row in sheet:
            for c in row:
                if c.data_type=='f':frows.append((sheet.title,c.coordinate,c.value if isinstance(c.value,str) else c.value.text))
    for sheet in wb:
        for row in sheet:
            for c in row:
                if c.data_type=='e' or (isinstance(c.value,str) and c.value in error_tokens):errors.append((sheet.title,c.coordinate,c.value))
    check('Delivered workbook contains no unexpected cached formula errors',[],errors)
    check('Production formulas never reference evaluation files',[],[(s,c) for s,c,f in frows if re.search('ground_truth|issue_manifest|evaluation/',f,re.I)])
    for function in ['VLOOKUP','SUMIFS','IF']:
        check('Required native formula present: '+function,True,any(re.search(r'\b'+function+r'\(',f,re.I) for _,_,f in frows))
    with zipfile.ZipFile(book) as z:
        entries=z.namelist();check('No external workbook links or macros',[],[x for x in entries if x.startswith('xl/externalLinks/') or 'vbaProject' in x])
        check('No genuine PivotTables are falsely claimed',0,sum(x.startswith('xl/pivotTables/pivotTable') and x.endswith('.xml') for x in entries))
        tables=sum(x.startswith('xl/tables/table') and x.endswith('.xml') for x in entries)
        charts=sum('/charts/chart' in x and x.endswith('.xml') for x in entries)
        check('Two native formula-backed charts installed',2,charts)
        chart_ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
        chart_sources=[];zero_axes=[]
        for entry in entries:
            if '/charts/chart' in entry and entry.endswith('.xml'):
                root=ET.fromstring(z.read(entry));chart_sources.extend(e.text for e in root.findall('.//c:f',chart_ns))
                zero_axes.extend(e.get('val')=='0' for e in root.findall('.//c:valAx/c:scaling/c:min',chart_ns))
        check('Both native chart value axes start at zero',True,len(zero_axes)==2 and all(zero_axes))
        check('Native chart sources point to formula cells in DASHBOARD',True,len(chart_sources)>=4 and all('DASHBOARD' in x.upper() for x in chart_sources))
        check('Native Excel Tables installed',True,tables>=10)
        check('RAW DELIVERIES protection safeguard installed',True,b'sheetProtection' in z.read('xl/worksheets/sheet2.xml'))
        check('Raw-sheet protection metadata has one valid element',1,len(ET.fromstring(z.read('xl/worksheets/sheet2.xml')).findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetProtection')))
        check('Named validation lists installed',True,all(x.encode() in z.read('xl/workbook.xml') for x in ['SupplierFilterChoices','StockpileFilterChoices','SourceRowIDs']))
    exports=[]
    for x in a:
        exports.append(dict(synthetic_notice='SYNTHETIC EDUCATIONAL DATA',source_row_id=x['source_row_id'],delivery_id=x['delivery_id'],delivery_date=x['delivery_date'],
          supplier_id=x['supplier_id'],stockpile_id=x['stockpile_id'],supplier_reported_wmt=x['supplier_reported_wmt'],weighbridge_net_wmt=x['weighbridge_net_wmt'],
          moisture_fraction=x['moisture_pct'],grade_dry_fraction=x['ni_grade_pct_dry'],dmt=x['dmt'],theoretical_contained_ni_t=x['contained_ni_t'],signed_variance_wmt=x['signed_variance_wmt'],variance_fraction=x['variance_fraction']))
    write_csv('exports/accepted_deliveries.csv',exports)
    sums=[]
    for dim,key,keys in [('total',None,['All accepted']),('supplier','supplier_id',[f'S{i:02}' for i in range(1,5)]),('stockpile','stockpile_id',[f'P{i:02}' for i in range(1,4)]),('date','delivery_date',[(date(2026,9,1)+timedelta(days=i)).isoformat() for i in range(60)])]:
        for value in keys:sums.append(dict(synthetic_notice='SYNTHETIC EDUCATIONAL DATA',dimension=dim,group=value,**aggregate(a if key is None else [x for x in a if x[key]==value])))
    write_csv('exports/summary.csv',sums)
    results=dict(synthetic=True,seed=20260901,python=platform.python_version(),openpyxl=openpyxl.__version__,counts=dict(counts),raw_count=len(raw),manifest_entries=len(manifest),
      totals=tot,scenario=combined,moisture_only=mo,grade_only=gr,volume_only=v,tests=TESTS,native_tables=tables,native_charts=charts,formula_cells=len(frows))
    local('reports/verification_results.json').write_text(json.dumps(results,indent=2),encoding='utf-8')
    failed=[t for t in TESTS if t['result']=='FAIL'];print(json.dumps(dict(checks=len(TESTS),failures=len(failed),counts=dict(counts),totals=tot),indent=2))
    for t in failed:print(json.dumps(t))
    wb.close();formulas.close()
    if failed:raise SystemExit(1)

if __name__=='__main__':main()
