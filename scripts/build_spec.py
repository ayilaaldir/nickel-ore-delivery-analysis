"""Workbook layout and native Excel formulas; production inputs only, no evaluation reads."""
import json
from datetime import date, timedelta
from project_paths import local
from analysis_model import analyze, aggregate

def col(n):
    s = ""
    while n:
        n, rem = divmod(n-1, 26); s = chr(65+rem)+s
    return s

def serial(v):
    return (date.fromisoformat(v)-date(1899,12,30)).days if v else ""

def main():
    data=json.loads(local("data/workbook_input.json").read_text(encoding="utf-8"))
    n=len(data['raw']); start=7; end=n+6
    names=['START HERE','RAW DELIVERIES','REFERENCES','DATA REVIEW','CLEAN DELIVERIES','SUMMARY','SCENARIOS','DASHBOARD','QA CHECKS']
    spec=dict(sheets=[], n=n, start=start, end=end, synthetic=True)
    def sheet(name, title, note, widths):
        s=dict(name=name,values=[dict(cell='A2',data=[[title]]),dict(cell='A3',data=[[note]]),dict(cell='A4',data=[['Back to START HERE']])],
               formulas=[],tables=[],widths=widths,formats=[],inputs=[],warnings=[],charts=[],freezeRows=6,freezeColumns=1,printRange='A1:L40')
        spec['sheets'].append(s);return s
    def val(s,cell,rows):s['values'].append(dict(cell=cell,data=rows))
    def form(s,cell,rows):s['formulas'].append(dict(cell=cell,data=rows))
    def fmt(s,ran,number):s['formats'].append(dict(range=ran,format=number))
    def table(s,ran,name):s['tables'].append(dict(range=ran,name=name))
    rawrange=f"'RAW DELIVERIES'!$A$7:$J${end}"
    rawids=f"'RAW DELIVERIES'!$A$7:$A${end}"
    reviewrange=f"'DATA REVIEW'!$A$7:$AK${end}"
    def rv(c):return f"'DATA REVIEW'!${c}$7:${c}${end}"
    def cl(c):return f"'CLEAN DELIVERIES'!${c}$7:${c}${end}"
    ov="'DATA REVIEW'!$AQ$7:$BA$46"; ovkeys="'DATA REVIEW'!$AQ$7:$AQ$46"
    def ovr(c):return f"'DATA REVIEW'!${c}$7:${c}$46"
    s=sheet(names[0], 'NickelOps', 'SYNTHETIC EDUCATIONAL PORTFOLIO | Ridla A. Gani',{'A':4,'B':28,'C':24,'D':24,'E':24,'F':24,'G':18,'H':18})
    s['freezeRows']=0;s['freezeColumns']=0;s['printRange']='A1:H39'
    val(s,'B6', [['Ore Delivery Reconciliation and Scenario Analysis'],['Independent portfolio project for operational data analysis.'],
         ['All companies, records and operational assumptions are fictional.'],['This is not paid work, mine operations, an engineering-approved model or mining experience.']])
    val(s,'B12',[[x] for x in names[1:]])
    val(s,'D12',[[x] for x in ['Preserved source snapshot, 310 rows','Reference lookups and illustrative thresholds','Review flags and keyed correction register at AQ6','Accepted metrics; filter status to Accepted','SUMIFS summaries; manual PivotTable instructions','Delivery-level assumptions and one-way sensitivities','Date, supplier and stockpile selectors','Population and metric reconciliations']])
    val(s,'B22',[[x] for x in [
       'Workflow: review RAW DELIVERIES, then filter DATA REVIEW by status or issue.',
       'Open DATA REVIEW AQ6. Enter a source_row_id, decision and reason before changing a field.',
       'Approve requires a reason. Corrections require Approve. Reject removes a competing submission.',
       'Exact retransmissions: retain the smallest source_row_id and exclude all later identical payloads.',
       'A conflicting delivery_id stays blocked until only one non-rejected, non-retransmitted submission remains.',
       'Auto accepts valid rows without review flags. Hold blocks a row. Invalid mathematics always blocks approval.',
       'WMT = wet metric tonnes. DMT = WMT x (1 - moisture fraction).',
       'Theoretical Ni tonnes = DMT x dry-basis nickel fraction. This is not recoverable or payable nickel, revenue or profit.',
       'Weighted grade = SUM(Ni tonnes) / SUM(DMT). Aggregate moisture = 1 - SUM(DMT) / SUM(WMT).',
       'Fractions are stored once: 1.5% = 0.015. A +2 percentage-point moisture change is entered as 2.',
       'Signed weight variance = weighbridge minus reported; percent uses positive reported WMT as denominator.',
       'Thresholds are illustrative review assumptions, not industry standards. Missing measurements remain unavailable.',
       'Fixed capacity: 310 raw/review/clean rows; 40 keyed decisions. Regenerate to change capacity; do not append rows.',
       'Formula summaries and charts recalculate. CSV exports and case study are dated delivery snapshots; rerun scripts to refresh.',
       'No genuine PivotTable is installed. SUMMARY includes exact manual Excel creation and refresh steps.',
       'Tested with artifact-tool and Python. Excel UI, native refresh and protection behavior were not tested.',
       'Legend: yellow = editable input; dark text = source/formula; status text and amber/red fill = warning.',
       'Offline Excel 365 workflow. No macros, external links, accounts or paid services required.']])
    raw=sheet(names[1],'Raw deliveries','SYNTHETIC | Preserved source snapshot. Do not edit. Missing values are blank.',{'A':17,'B':18,'C':15,'D':15,'E':15,'F':20,'G':22,'H':18,'I':20,'J':65})
    fields=list(data['raw'][0]);val(raw,'A6',[fields]); rawvals=[]
    for r in data['raw']:
        rawvals.append([serial(r[k]) if k=='delivery_date' else r[k] for k in fields])
    val(raw,'A7',rawvals); table(raw,f'A6:J{end}','RawDeliveries');fmt(raw,f'C7:C{end}','dd-mmm-yyyy');fmt(raw,f'F7:G{end}','#,##0.000');fmt(raw,f'H7:I{end}','0.00%');raw['printRange']='A1:J18'
    ref=sheet(names[2],'Reference data and review assumptions','SYNTHETIC | All thresholds below are illustrative assumptions, not industry standards.',{'A':28,'B':32,'C':22,'D':14,'E':75,'G':24,'H':24})
    val(ref,'A6',[list(data['suppliers'][0])]+[list(x.values()) for x in data['suppliers']]);table(ref,'A6:C10','Suppliers')
    val(ref,'A13',[list(data['stockpiles'][0])]+[list(x.values()) for x in data['stockpiles']]);table(ref,'A13:B16','Stockpiles')
    val(ref,'A19',[['review_assumption','value','unit','role','explanatory_note'],
        ['Absolute weight variance',.03,'fraction','Review','Review when ABS((net - reported) / reported) > threshold.'],
        ['Moisture review ceiling',.40,'fraction','Review','Mathematically valid values above this ceiling require documented approval.'],
        ['Dry-basis Ni review ceiling',.03,'fraction','Review','Mathematically valid values above this ceiling require documented approval.']]);table(ref,'A19:E22','ReviewThresholds');fmt(ref,'B20:B22','0.0%');ref['inputs']=['B20:B22']
    val(ref,'A25', [['Mathematical gates'],['Both WMT measurements must be numeric and positive. Moisture and grade must be numeric fractions in [0, 1].'],
         ['Date must be a numeric whole-day Excel date from 1 through 2958465. Identifiers must map to active references.'],
         ['Zero DMT is mathematically allowed at 100% moisture; weighted grade then reports Unavailable.'],
         ['Review thresholds do not make an invalid fraction valid. Approval cannot override mathematical gates.']])
    val(ref,'G6',[['supplier_selector'],['All']]+[[x['supplier_id']] for x in data['suppliers']]);val(ref,'H6',[['stockpile_selector'],['All']]+[[x['stockpile_id']] for x in data['stockpiles']]);ref['printRange']='A1:H29'
    rev=sheet(names[3],'Delivery review and keyed decisions','SYNTHETIC | Status in B. Original values in AB:AJ. Yellow keyed register at AQ6. Reasons required.',{'A':17,'B':22,'C':19,'D':16,'E':14,'F':14,'G':19,'H':19,'I':17,'J':17,'K':27,'L':22,'M':26,'N':22,'O':18,'P':14,'Q':65,'R':35,'S':40,'T':19,'U':17,'V':20,'W':20,'X':19,'Y':15,'Z':15,'AA':19,'AB':19,'AC':16,'AD':18,'AE':18,'AF':20,'AG':20,'AH':19,'AI':19,'AJ':65,'AK':45,'AP':4,'AQ':17,'AR':15,'AS':75,'AT':16,'AU':18,'AV':18,'AW':22,'AX':23,'AY':22,'AZ':22,'BA':50})
    heads=['source_row_id','inclusion_status','delivery_id','delivery_date','supplier_id','stockpile_id','reported_WMT','weighbridge_WMT','moisture_fraction','grade_dry_fraction','supplier_name','stockpile_name','normalization_or_override','exact_raw_signature','retransmission','review_decision','correction_reason','override_control','mathematical_status','illustrative_review','live_ID_count','signed_variance_WMT','variance_fraction','weight_exception','accepted_flag','excluded_duplicate_flag','unresolved_rejected_flag',
           'original_delivery_id','original_date','original_supplier_id','original_stockpile_id','original_reported_WMT','original_net_WMT','original_moisture_fraction','original_grade_fraction','original_record_note','reviewer_notes']
    rev['values'][0]['data']=[['Review']];rev['values'][1]['data']=[['SYNTHETIC DATA']]
    val(rev,'A6',[heads]); val(rev,'AQ6',[list(data['decisions'][0])]);val(rev,'AQ7',[[serial(x[k]) if k=='corrected_date' else x[k] for k in data['decisions'][0]] for x in data['decisions']]);table(rev,'AQ6:BA46','ReviewDecisions');rev['inputs']=['AQ7:BA46'];rev['freezeColumns']=2
    val(rev,'AQ2',[['KEYED REVIEW INPUTS']]);val(rev,'AQ3',[['Use source_row_id. Blank corrected fields keep original measurements. Zero remains a valid explicit correction.']]);val(rev,'AQ4',[['Decision: Auto / Approve / Reject / Hold. Correction requires Approve plus reason; duplicate keys block the row.']]);
    form(rev,'A5',[["=HYPERLINK(\"#'DATA REVIEW'!AQ6\",\"Open keyed review inputs\")"]])
    rowform=[]
    for i, row in enumerate(data['raw'],7):
        original={}
        for j, rawcol in enumerate(range(2,11),28):
            target=col(j); rr=f"'RAW DELIVERIES'!${col(rawcol)}$7:${col(rawcol)}${end}"
            original[target]=f'=IF(COUNTIFS({rawids},$A{i},{rr},"<>")=0,"",VLOOKUP($A{i},{rawrange},{rawcol},FALSE))'
        def corrected(orig, oc, idx, norm=False):
            expr=f'IF(COUNTIFS({ovkeys},$A{i},{ovr(oc)},"<>")=0,{orig}{i},VLOOKUP($A{i},{ov},{idx},FALSE))'
            return '='+('UPPER(TRIM('+expr+'))' if norm else expr)
        decision=f'=IF(COUNTIFS({ovkeys},$A{i},{ovr("AR")},"<>")=0,"Auto",VLOOKUP($A{i},{ov},2,FALSE))'
        reason=f'=IF(COUNTIFS({ovkeys},$A{i},{ovr("AS")},"<>")=0,"",VLOOKUP($A{i},{ov},3,FALSE))'
        override_count='+'.join(f'COUNTIFS({ovkeys},$A{i},{ovr(c)},"<>")' for c in ['AT','AU','AV','AW','AX','AY','AZ'])
        control=f'=IF(COUNTIFS({ovkeys},$A{i})>1,"BLOCKED: duplicate key",IF(AND(P{i}<>"Auto",P{i}<>"Approve",P{i}<>"Reject",P{i}<>"Hold"),"BLOCKED: decision",IF(AND(OR(P{i}="Approve",P{i}="Reject",({override_count})>0),LEN(TRIM(Q{i}))=0),"BLOCKED: reason required",IF(AND(({override_count})>0,P{i}<>"Approve"),"BLOCKED: correction needs Approve","OK"))))'
        checks=[(f'C{i}=""','Missing delivery ID'),(f'NOT(AND(ISNUMBER(D{i}),D{i}>=1,D{i}<=2958465,IFERROR(MOD(D{i},1),1)=0))','Missing/invalid date'),
                (f'COUNTIFS(REFERENCES!$A$7:$A$10,E{i},REFERENCES!$C$7:$C$10,"Active")=0','Unknown/inactive supplier'),(f'COUNTIFS(REFERENCES!$A$14:$A$16,F{i})=0','Unknown stockpile')]
        for c,label,fraction in [('G','Reported WMT',False),('H','Weighbridge WMT',False),('I','Moisture',True),('J','Grade',True)]:
            valid=f'AND(ISNUMBER({c}{i}),{c}{i}>=0,{c}{i}<=1)' if fraction else f'AND(ISNUMBER({c}{i}),{c}{i}>0)'
            checks.append((f'NOT({valid})',f'Missing/invalid {label}'))
        flags='&'.join(f'IF({test},"{label}; ","")' for test,label in checks)
        math=f'=IF(OR({",".join(test for test,_ in checks)}),{flags},"VALID")'
        signature='&"|"&'.join([f'{c}{i}' if c in ['AB','AD','AE','AJ'] else f'IF({c}{i}="","<blank>",TEXT({c}{i},"0.000000"))' for c in ['AB','AC','AD','AE','AF','AG','AH','AI','AJ']])
        formulas={
          'B':f'=IF(O{i}="YES","Excluded duplicate",IF(R{i}<>"OK","Unresolved",IF(P{i}="Reject","Rejected",IF(OR(S{i}<>"VALID",U{i}>1,P{i}="Hold",AND(T{i}="REVIEW",P{i}<>"Approve")),"Unresolved","Accepted"))))',
          'C':f'=UPPER(TRIM(AB{i}))','D':corrected('AC','AT',4),'E':corrected('AD','AU',5,True),'F':corrected('AE','AV',6,True),
          'G':corrected('AF','AW',7),'H':corrected('AG','AX',8),'I':corrected('AH','AY',9),'J':corrected('AI','AZ',10),
          'K':f'=IFERROR(VLOOKUP(E{i},REFERENCES!$A$7:$C$10,2,FALSE),"UNKNOWN")',
          'L':f'=IFERROR(VLOOKUP(F{i},REFERENCES!$A$14:$B$16,2,FALSE),"UNKNOWN")',
          'M':f'=IF(OR(C{i}<>AB{i},D{i}<>AC{i},E{i}<>AD{i},F{i}<>AE{i},G{i}<>AF{i},H{i}<>AG{i},I{i}<>AH{i},J{i}<>AI{i}),"NORMALIZED/OVERRIDDEN","UNCHANGED")',
          'N':'='+signature,
          'O':f'=IF(COUNTIFS($N$7:$N${end},N{i},$A$7:$A${end},"<"&A{i})>0,"YES","NO")',
          'P':decision,'Q':reason,'R':control,'S':math,
          'T':f'=IF(S{i}<>"VALID","UNAVAILABLE",IF(OR(ABS(W{i})>REFERENCES!$B$20,I{i}>REFERENCES!$B$21,J{i}>REFERENCES!$B$22),"REVIEW","OK"))',
          'U':f'=IF(C{i}="",0,COUNTIFS($C$7:$C${end},C{i},$O$7:$O${end},"NO")-COUNTIFS($C$7:$C${end},C{i},$O$7:$O${end},"NO",$P$7:$P${end},"Reject",$R$7:$R${end},"OK"))',
          'V':f'=IF(AND(ISNUMBER(G{i}),G{i}>0,ISNUMBER(H{i}),H{i}>0),H{i}-G{i},"")',
          'W':f'=IF(ISNUMBER(V{i}),V{i}/G{i},"")',
          'X':f'=IF(ISNUMBER(W{i}),IF(ABS(W{i})>REFERENCES!$B$20,"EXCEPTION","OK"),"UNAVAILABLE")',
          'Y':f'=IF(B{i}="Accepted",1,0)','Z':f'=IF(B{i}="Excluded duplicate",1,0)','AA':f'=IF(OR(B{i}="Unresolved",B{i}="Rejected"),1,0)',
          **original,'AK':f'=IF(COUNTIFS({ovkeys},$A{i},{ovr("BA")},"<>")=0,"",VLOOKUP($A{i},{ov},11,FALSE))'}
        val(rev,f'A{i}',[[row['source_row_id']]])
        rowform.append([formulas[col(j)] for j in range(2,38)])
    form(rev,'B7',rowform);table(rev,f'A6:AK{end}','DeliveryReview')
    for ran in [f'D7:D{end}',f'AC7:AC{end}','AT7:AT46']:fmt(rev,ran,'dd-mmm-yyyy')
    for ran in [f'I7:J{end}',f'W7:W{end}',f'AH7:AI{end}','AY7:AZ46']:fmt(rev,ran,'0.00%')
    for ran in [f'G7:H{end}',f'V7:V{end}',f'AF7:AG{end}','AW7:AX46']:fmt(rev,ran,'#,##0.000')
    rev['warnings']=[dict(range=f'B7:B{end}',formula='$B7="Unresolved"'),dict(range='R7:S'+str(end),formula='$S7<>"VALID"')];rev['printRange']='A1:J19'
    clean=sheet(names[4],'Accepted delivery calculations','SYNTHETIC | One keyed row per source. Non-accepted metrics stay blank. Filter inclusion_status to Accepted.',{'A':17,'B':19,'C':16,'D':15,'E':15,'F':27,'G':22,'H':22,'I':20,'J':20,'K':18,'L':19,'M':24,'N':23,'O':20,'P':19,'Q':16,'R':16,'S':21,'T':20})
    cheads=['source_row_id','delivery_id','delivery_date','supplier_id','stockpile_id','supplier_name','stockpile_name','inclusion_status','weighbridge_WMT','DMT','moisture_fraction','grade_dry_fraction','theoretical_contained_Ni_t','signed_variance_WMT','variance_fraction','weight_exception','accepted_flag','selected_scope_flag','supplier_reported_WMT','accepted_ID_count']
    val(clean,'A6',[cheads]); cf=[]
    for i,row in enumerate(data['raw'],7):
        val(clean,f'A{i}',[[row['source_row_id']]])
        def lookup(index):return f'VLOOKUP($A{i},{reviewrange},{index},FALSE)'
        mapping={2:3,3:4,4:5,5:6,6:11,7:12,9:8,11:9,12:10,14:22,15:23,16:24,19:7}
        f={j:f'=IF($H{i}="Accepted",{lookup(idx)},"")' for j,idx in mapping.items()}
        f[8]='='+lookup(2);f[10]=f'=IF(H{i}="Accepted",I{i}*(1-K{i}),"")';f[13]=f'=IF(H{i}="Accepted",J{i}*L{i},"")';f[17]=f'=IF(H{i}="Accepted",1,0)'
        f[18]=f'=IF(AND(Q{i}=1,DASHBOARD!$B$8="VALID",C{i}>=DASHBOARD!$B$6,C{i}<=DASHBOARD!$D$6,OR(DASHBOARD!$F$6="All",D{i}=DASHBOARD!$F$6),OR(DASHBOARD!$H$6="All",E{i}=DASHBOARD!$H$6)),1,0)'
        f[20]=f'=IF(Q{i}=1,COUNTIFS($B$7:$B${end},B{i},$Q$7:$Q${end},1),0)'
        cf.append([f[j] for j in range(2,21)])
    form(clean,'B7',cf);table(clean,f'A6:T{end}','CleanDeliveries');fmt(clean,f'C7:C{end}','dd-mmm-yyyy');fmt(clean,f'I7:J{end}','#,##0.000');fmt(clean,f'K7:L{end}','0.00%');fmt(clean,f'M7:N{end}','#,##0.000');fmt(clean,f'O7:O{end}','0.00%');fmt(clean,f'S7:S{end}','#,##0.000');clean['printRange']='A1:M19'
    summ=sheet(names[5],'Delivery summaries','SYNTHETIC | SUMIFS summaries of all accepted records. These are formula tables, not PivotTables.',{'A':27,'B':17,'C':19,'D':19,'E':23,'F':23,'G':22,'H':21,'J':5,'K':30,'L':30})
    shead=['supplier_id','accepted_count','WMT','DMT','theoretical_Ni_t','weighted_grade_dry','weighted_moisture','weight_exceptions']
    def summary_block(cell, keys, keycol, name, header):
        val(summ,cell,[header]); rownum=int(''.join(x for x in cell if x.isdigit()))+1
        for i,key in enumerate(keys,rownum):
            val(summ,f'A{i}',[[key]])
            filt=f'{cl(keycol)},A{i}'
            form(summ,f'B{i}',[[f'=COUNTIFS({cl("Q")},1,{filt})',f'=SUMIFS({cl("I")},{cl("Q")},1,{filt})',f'=SUMIFS({cl("J")},{cl("Q")},1,{filt})',f'=SUMIFS({cl("M")},{cl("Q")},1,{filt})',f'=IF(D{i}>0,E{i}/D{i},"Unavailable")',f'=IF(C{i}>0,1-D{i}/C{i},"Unavailable")',f'=COUNTIFS({cl("Q")},1,{filt},{cl("P")},"EXCEPTION")']])
        last=rownum+len(keys)-1;table(summ,f'A{rownum-1}:H{last}',name);fmt(summ,f'C{rownum}:E{last}','#,##0.000');fmt(summ,f'F{rownum}:G{last}','0.00%')
    summary_block('A6',[x['supplier_id'] for x in data['suppliers']],'D','SupplierSummary',shead)
    summary_block('A14',[x['stockpile_id'] for x in data['stockpiles']],'E','StockpileSummary',['stockpile_id']+shead[1:])
    summary_block('A23',[serial((date(2026,9,1)+timedelta(days=i)).isoformat()) for i in range(60)],'C','DailySummary',['delivery_date']+shead[1:]);fmt(summ,'A24:A83','dd-mmm-yyyy')
    val(summ,'K6',[[x] for x in ['Manual genuine PivotTable (not installed)', '1. In Excel, select CLEAN DELIVERIES A6:T316.', '2. Insert > PivotTable > From Table/Range.', '3. Table/Range = CleanDeliveries. Existing Worksheet = SUMMARY!K20.',
         '4. Add inclusion_status to Filters; select Accepted only.', '5. Add supplier_name to Rows. Add accepted_flag, weighbridge_WMT, DMT,', '   theoretical_contained_Ni_t to Values; set every value to SUM.',
         '6. Repeat at SUMMARY!K35 with stockpile_name in Rows.', '7. For a date trend, use delivery_date in Rows; ungroup if Excel groups it.',
         '8. After review changes: Formulas > Calculation Options > Automatic;', '   Calculate Now, then Data > Refresh All. Reapply Accepted filter if needed.',
         '9. Reconcile Pivot grand totals to SUMMARY C7:E10 totals.', '10. Never average grade or moisture in a PivotTable. Divide Ni/DMT,', '    or compute 1-DMT/WMT from matching grand totals outside the Pivot.',
         'Native creation/refresh requires your installed Excel; it was not tested here.']]);summ['printRange']='A1:H18'
    scen=sheet(names[6],'Baseline and user-adjusted scenario','SYNTHETIC | Assumptions, not forecasts. Scope follows DASHBOARD selectors. No clipping of invalid fractions.',{'A':29,'B':20,'C':20,'D':22,'E':20,'F':20,'G':29,'H':22,'I':22,'J':25,'K':40,'L':20,'M':23,'N':18})
    val(scen,'A5',[['Volume change (relative %)','','','Moisture change (pp)','','','Dry-basis grade change (pp)']]);val(scen,'B6',[[.1]]);val(scen,'E6',[[2]]);val(scen,'H6',[[.1]]);fmt(scen,'B6','0.0%');fmt(scen,'E6','0.00" pp"');fmt(scen,'H6','0.00" pp"');scen['inputs']=['B6','E6','H6']
    val(scen,'A8',[['Scenario status']]);form(scen,'B9',[[f'=IF(DASHBOARD!$B$8<>"VALID","BLOCKED: invalid dashboard scope",IF(NOT(AND(ISNUMBER(B6),ISNUMBER(E6),ISNUMBER(H6))),"BLOCKED: all three inputs must be numeric",IF(B6<-1,"BLOCKED: volume below -100%",IF(SUM(B31:B{n+30})=0,"NO ACCEPTED RECORDS IN SCOPE",IF(SUM(N31:N{n+30})>0,"BLOCKED: moisture or grade outside [0,1]","VALID")))))']])
    val(scen,'A11',[['Metric','Baseline','User scenario','Change']]);metrics=[('WMT','C','H'),('DMT','F','L'),('Theoretical contained Ni (t)','G','M')]
    for i,(label,basec,casec) in enumerate(metrics,12):
        val(scen,f'A{i}',[[label]]);form(scen,f'B{i}',[[f'=SUM({basec}31:{basec}{n+30})',f'=IF($B$9="VALID",SUM({casec}31:{casec}{n+30}),"Unavailable")',f'=IF(ISNUMBER(C{i}),C{i}-B{i},"Unavailable")']])
    val(scen,'A15',[['Weighted grade (dry basis)']]);form(scen,'B15',[[ '=IF(B13>0,B14/B13,"Unavailable")','=IF(AND(ISNUMBER(C13),C13>0),C14/C13,"Unavailable")','=IF(AND(ISNUMBER(B15),ISNUMBER(C15)),C15-B15,"Unavailable")']]);fmt(scen,'B12:D14','#,##0.000');fmt(scen,'B15:D15','0.00%')
    val(scen,'A17',[['Moisture sensitivity (other changes zero)']]);val(scen,'G17',[['Grade sensitivity (other changes zero)']])
    val(scen,'A18',[['Moisture delta (pp)','DMT','Theoretical Ni (t)','Status']]);val(scen,'G18',[['Grade delta (pp)','DMT','Theoretical Ni (t)','Status']])
    for i,delta in enumerate([-4,-2,0,2,4],19):
        val(scen,f'A{i}',[[delta]]);val(scen,f'G{i}',[[delta/10]])
        for inputc,frac,basec in [('A','K','moisture'),('G','L','grade')]:
            statusc='D' if inputc=='A' else 'J'; dc='B' if inputc=='A' else 'H'; nc='C' if inputc=='A' else 'I'
            form(scen,f'{statusc}{i}',[[f'=IF(DASHBOARD!$B$8<>"VALID","BLOCKED: scope",IF(SUM({cl("R")})=0,"NO RECORDS",IF(COUNTIFS({cl("R")},1,{cl(frac)},"<"&(-{inputc}{i}/100))+COUNTIFS({cl("R")},1,{cl(frac)},">"&(1-{inputc}{i}/100))>0,"BLOCKED","VALID")))']])
            dcalc=f'$B$13-$B$12*A{i}/100' if inputc=='A' else '$B$13'
            ncalc=f'$B$14-A{i}/100*SUMPRODUCT({cl("R")},{cl("I")},{cl("L")})' if inputc=='A' else f'$B$14+G{i}/100*$B$13'
            form(scen,f'{dc}{i}',[[f'=IF({statusc}{i}="VALID",{dcalc},"Unavailable")',f'=IF({statusc}{i}="VALID",{ncalc},"Unavailable")']])
    fmt(scen,'A19:A23','0.0" pp"');fmt(scen,'G19:G23','0.0" pp"');fmt(scen,'B19:C23','#,##0.000');fmt(scen,'H19:I23','#,##0.000')
    val(scen,'A25',[[x] for x in ['Example: 100 WMT, 30% moisture, 1.50% dry grade gives 70 DMT and 1.05 t theoretical Ni.',
         '+10% volume, +2 pp moisture and +0.10 pp grade gives 110 WMT, 32% moisture, 1.60% grade, 74.8 DMT and 1.1968 t Ni.',
         'Scenarios recompute every scoped accepted delivery. Invalid scenario totals are unavailable; no partial totals are reported.']])
    val(scen,'A30',[['source_row_id','in_scope','baseline_WMT','baseline_moisture','baseline_grade_dry','baseline_DMT','baseline_Ni_t','scenario_WMT','scenario_moisture','scenario_grade_dry','row_validation','scenario_DMT','scenario_Ni_t','invalid_flag']]);sf=[]
    for i,r in enumerate(data['raw'],31):
        val(scen,f'A{i}',[[r['source_row_id']]])
        cleanrange=f"'CLEAN DELIVERIES'!$A$7:$T${end}"
        forms={'B':f'=VLOOKUP($A{i},{cleanrange},18,FALSE)'}
        for c,index in [('C',9),('D',11),('E',12),('F',10),('G',13)]:forms[c]=f'=IF(B{i}=1,VLOOKUP($A{i},{cleanrange},{index},FALSE),"")'
        forms.update(H=f'=IF(AND(B{i}=1,ISNUMBER($B$6),$B$6>=-1),C{i}*(1+$B$6),"")',I=f'=IF(AND(B{i}=1,ISNUMBER($E$6)),D{i}+$E$6/100,"")',J=f'=IF(AND(B{i}=1,ISNUMBER($H$6)),E{i}+$H$6/100,"")',
            K=f'=IF(B{i}=0,"OUT OF SCOPE",IF(NOT(AND(ISNUMBER($B$6),ISNUMBER($E$6),ISNUMBER($H$6))),"BLOCKED: numeric inputs required",IF($B$6<-1,"BLOCKED: volume below -100%",IF(OR(I{i}<0,I{i}>1,J{i}<0,J{i}>1),"BLOCKED: fraction outside [0,1]","VALID"))))',
            L=f'=IF(K{i}="VALID",H{i}*(1-I{i}),"")',M=f'=IF(K{i}="VALID",L{i}*J{i},"")',N=f'=IF(AND(B{i}=1,K{i}<>"VALID"),1,0)')
        sf.append([forms[col(j)] for j in range(2,15)])
    form(scen,'B31',sf);table(scen,f'A30:N{n+30}','ScenarioDeliveries');fmt(scen,f'C31:C{n+30}','#,##0.000');fmt(scen,f'D31:E{n+30}','0.00%');fmt(scen,f'F31:H{n+30}','#,##0.000');fmt(scen,f'I31:J{n+30}','0.00%');fmt(scen,f'L31:M{n+30}','#,##0.000');scen['printRange']='A1:J27';scen['freezeRows']=0;scen['warnings']=[dict(range='B9:H9',formula='$B$9<>"VALID"')]
    dash=sheet(names[7],'Ore delivery dashboard','SYNTHETIC EDUCATIONAL DATA | Selected accepted records. Global data quality is shown separately.',{c:14 for c in [col(j) for j in range(1,14)]})
    dash['freezeRows']=0;dash['freezeColumns']=0;dash['printRange']='A1:M43'
    val(dash,'A5',[['Start date','','End date','','Supplier ID','','Stockpile ID']]);val(dash,'B6',[[serial('2026-09-01')]]);val(dash,'D6',[[serial('2026-10-30')]]);val(dash,'F6',[['All']]);val(dash,'H6',[['All']]);fmt(dash,'B6','dd-mmm-yyyy');fmt(dash,'D6','dd-mmm-yyyy');dash['inputs']=['B6','D6','F6','H6']
    val(dash,'A8',[['Scope status']]);form(dash,'B8',[[ '=IF(NOT(AND(ISNUMBER(B6),ISNUMBER(D6),B6>=1,D6<=2958465,B6<=D6,IFERROR(MOD(B6,1),1)=0,IFERROR(MOD(D6,1),1)=0)),"BLOCKED: invalid dates",IF(AND(F6<>"All",COUNTIFS(REFERENCES!$A$7:$A$10,F6)=0),"BLOCKED: supplier",IF(AND(H6<>"All",COUNTIFS(REFERENCES!$A$14:$A$16,H6)=0),"BLOCKED: stockpile","VALID")))']])
    val(dash,'A10',[['SELECTED ACCEPTED RECORDS']]);labels=['Accepted count','WMT','DMT','Weighted Ni grade','Theoretical Ni (t)','Aggregate moisture','Weight exceptions']
    for i,label in enumerate(labels,11):val(dash,f'A{i}',[[label]])
    for cell,formula in {'D11':f'=IF($B$8="VALID",SUM({cl("R")}),"Unavailable")','D12':f'=IF($B$8="VALID",SUMIFS({cl("I")},{cl("R")},1),"Unavailable")','D13':f'=IF($B$8="VALID",SUMIFS({cl("J")},{cl("R")},1),"Unavailable")','D14':'=IF(AND(ISNUMBER(D13),D13>0),D15/D13,"Unavailable")','D15':f'=IF($B$8="VALID",SUMIFS({cl("M")},{cl("R")},1),"Unavailable")','D16':'=IF(AND(ISNUMBER(D12),D12>0),1-D13/D12,"Unavailable")','D17':f'=IF($B$8="VALID",COUNTIFS({cl("R")},1,{cl("P")},"EXCEPTION"),"Unavailable")'}.items():form(dash,cell,[[formula]])
    fmt(dash,'D12:D13','#,##0.0');fmt(dash,'D14','0.00%');fmt(dash,'D15','#,##0.000');fmt(dash,'D16','0.00%')
    val(dash,'A21',[['SAME SCOPE COMPARISON']]);val(dash,'A22',[['Metric','Baseline','Scenario','Change']])
    for di,si,label in [(23,12,'WMT'),(24,13,'DMT'),(25,14,'Ni tonnes'),(26,15,'Grade (dry)')]:
        val(dash,f'A{di}',[[label]]);form(dash,f'B{di}',[[f'=SCENARIOS!B{si}',f'=SCENARIOS!C{si}',f'=SCENARIOS!D{si}']])
    fmt(dash,'B23:D25','#,##0.000');fmt(dash,'B26:D26','0.00%');val(dash,'A28',[['Scenario status']]);form(dash,'B29',[['=SCENARIOS!B9']]);val(dash,'A30',[['Scenario inputs are on SCENARIOS.']])
    val(dash,'A33',[['GLOBAL QUALITY: ALL RAW ROWS']])
    for i,label,status in [(34,'Raw rows',None),(35,'Accepted','Accepted'),(36,'Excluded duplicate','Excluded duplicate'),(37,'Unresolved','Unresolved'),(38,'Rejected','Rejected')]:
        val(dash,f'A{i}',[[label]]);form(dash,f'D{i}',[[f'=COUNTA({rawids})' if status is None else f'=COUNTIFS({rv("B")},"{status}")']])
    val(dash,'A40',[['Global quality counts do not follow selectors.'],['Ni means theoretical contained nickel.'],['Zero accepted DMT makes grade unavailable.']])
    val(dash,'A49',[['date_label','scoped_WMT']])
    for i in range(60):
        rr=50+i; d=(date(2026,9,1)+timedelta(days=i));val(dash,f'A{rr}',[[d.strftime('%d-%b')]])
        form(dash,f'B{rr}',[[f'=IF($B$8="VALID",SUMIFS({cl("I")},{cl("R")},1,{cl("C")},{serial(d.isoformat())}),"")']])
    val(dash,'J49',[['Case','Theoretical Ni (t)']]);val(dash,'J50',[['Baseline'],['Scenario']]);form(dash,'K50',[['=IF(ISNUMBER(B25),B25,"")'],['=IF(ISNUMBER(C25),C25,"")']])
    dash['charts']=[dict(type='line',ranges=['A49:B109'],position=['F10','M24'],title='Synthetic daily accepted WMT',yFormat='#,##0'),dict(type='bar',ranges=['J49:K51'],position=['F27','M41'],title='Synthetic theoretical Ni (t)',yFormat='#,##0.0')]
    dash['warnings']=[dict(range='B8:H8',formula='$B$8<>"VALID"'),dict(range='B29:E29',formula='$B$29<>"VALID"')]
    qa=sheet(names[8],'Quality assurance and reconciliation','SYNTHETIC | Live checks are separate from executed test evidence in test_report.md. Tolerances use unrounded values.',{'A':49,'B':24,'C':24,'D':19,'E':35,'F':65})
    val(qa,'A6',[['check','actual','expected','tolerance','result','scope_or_interpretation']]);expected=aggregate([x for x in analyze() if x['status']=='Accepted'])
    checks=[('All raw rows accounted for',f'SUM({rv("Y")})+SUM({rv("Z")})+SUM({rv("AA")})',n,0,'Mutually exclusive status flags; all source rows'),
        ('Accepted flags match clean population',f'SUM({cl("Q")})-COUNTIFS({rv("B")},"Accepted")',0,0,'All accepted records'),
        ('Accepted identifiers unique',f'COUNTIFS({cl("T")},">1")',0,0,'No accepted duplicate delivery identifiers'),
        ('Invalid measurements cannot be accepted',f'COUNTIFS({rv("S")},"<>VALID",{rv("B")},"Accepted")',0,0,'Missing/invalid values remain blocked'),
        ('Non-accepted DMT remains blank',f'SUMPRODUCT(({cl("H")}<>"Accepted")*ISNUMBER({cl("J")}))',0,0,'No numeric DMT, including false zero, for non-accepted measurements'),
        ('Supplier WMT reconciliation',f'SUM(SUMMARY!C7:C10)-SUM({cl("I")})',0,1e-7,'Same all-accepted population'),
        ('Stockpile DMT reconciliation',f'SUM(SUMMARY!D15:D17)-SUM({cl("J")})',0,1e-7,'Same all-accepted population'),
        ('Date contained Ni reconciliation',f'SUM(SUMMARY!E24:E83)-SUM({cl("M")})',0,1e-7,'Fixed 60-day reporting window; out-of-window corrections require extension'),
        ('Initial snapshot WMT Python control',f'SUM({cl("I")})',expected['wmt'],1e-7,'Snapshot evidence: legitimate review edits may change this control'),
        ('Initial snapshot DMT Python control',f'SUM({cl("J")})',expected['dmt'],1e-7,'Snapshot evidence: legitimate review edits may change this control'),
        ('Initial snapshot Ni Python control',f'SUM({cl("M")})',expected['contained_ni_t'],1e-7,'Snapshot evidence: legitimate review edits may change this control'),
        ('Duplicate keyed review IDs',f'SUMPRODUCT(({ovkeys}<>"")*(COUNTIF({ovkeys},{ovkeys})>1))',0,0,'Duplicate corrections block the associated source row'),
        ('Orphan keyed review IDs',f'SUMPRODUCT(({ovkeys}<>"")*(COUNTIF({rawids},{ovkeys})=0))',0,0,'Every populated correction key must exist in raw data')]
    for i,(label,actual,exp,tol,note) in enumerate(checks,7):
        val(qa,f'A{i}',[[label]]);val(qa,f'C{i}',[[exp,tol]]);val(qa,f'F{i}',[[note]])
        result=f'=IF(ABS(B{i}-C{i})<=D{i},"PASS","'+('REVIEW: inputs changed' if 'snapshot' in label else 'FAIL')+'")'
        form(qa,f'B{i}',[['='+actual]]);form(qa,f'E{i}',[[result]])
    fmt(qa,'B7:D19','0.0000000');table(qa,'A6:F19','QAChecks');qa['warnings']=[dict(range='E7:E19',formula='LEFT(E7,4)<>"PASS"')]
    fmt(qa,'B7:C11','0');fmt(qa,'B18:C19','0');fmt(qa,'D7:D19','General')
    val(qa,'A22',[[x] for x in ['Recalculation and mutation evidence is in test_report.md and reports/engine_tests.json.',
         'Excel COM was unavailable (logon-session error). No Excel UI or native PivotTable refresh testing is claimed.',
         'Installed formulas and charts are not proof of native Excel interaction. Test in Excel before extending this portfolio model.',
         'RAW DELIVERIES is protected as an accidental-edit safeguard, not a security control. No password is used.']]);qa['printRange']='A1:F25'
    spec['validations']=[dict(sheet='DASHBOARD',range='F6',rule=dict(type='list',formula1='REFERENCES!$G$7:$G$11')),dict(sheet='DASHBOARD',range='H6',rule=dict(type='list',formula1='REFERENCES!$H$7:$H$10')),
       dict(sheet='DATA REVIEW',range='AQ7:AQ46',rule=dict(type='list',formula1=f"'RAW DELIVERIES'!$A$7:$A${end}")),dict(sheet='DATA REVIEW',range='AR7:AR46',rule=dict(type='list',values=['Auto','Approve','Reject','Hold'])),
       dict(sheet='SCENARIOS',range='B6',rule=dict(type='decimal',operator='greaterThanOrEqual',formula1=-1)),dict(sheet='SCENARIOS',range='E6',rule=dict(type='decimal',operator='between',formula1=-100,formula2=100)),dict(sheet='SCENARIOS',range='H6',rule=dict(type='decimal',operator='between',formula1=-100,formula2=100))]
    spec['initial_controls']=expected
    local('tmp/workbook_spec.json').write_text(json.dumps(spec,indent=2),encoding='utf-8')
    print(f'Workbook specification: {n} source rows, {len(names)} required sheets.')

if __name__=='__main__':main()
