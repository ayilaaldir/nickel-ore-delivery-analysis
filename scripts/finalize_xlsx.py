"""Finalize OOXML views, named validations, chart axes and a verified QA blank-check compatibility fix."""
import json
import zipfile
import xml.etree.ElementTree as ET
import copy
from project_paths import local

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('',NS)
def q(name):return '{'+NS+'}'+name

def main():
    spec=json.loads(local('tmp/workbook_spec.json').read_text(encoding='utf-8'))
    book=local('NickelOps.xlsx')
    with zipfile.ZipFile(book) as z: parts={name:z.read(name) for name in z.namelist()}
    wb=ET.fromstring(parts['xl/workbook.xml'])
    dn=wb.find(q('definedNames'))
    if dn is None:
        dn=ET.Element(q('definedNames')); before=wb.find(q('calcPr'))
        wb.insert(list(wb).index(before) if before is not None else len(wb),dn)
    for old in list(dn):
        if old.get('name') in ('SupplierFilterChoices','StockpileFilterChoices','SourceRowIDs','_xlnm.Print_Area'):dn.remove(old)
    for name,formula in [('SupplierFilterChoices',"'REFERENCES'!$G$7:$G$11"),('StockpileFilterChoices',"'REFERENCES'!$H$7:$H$10"),('SourceRowIDs',f"'RAW DELIVERIES'!$A$7:$A${spec['end']}")]:
        e=ET.SubElement(dn,q('definedName'),{'name':name});e.text=formula
    for i,s in enumerate(spec['sheets']):
        e=ET.SubElement(dn,q('definedName'),{'name':'_xlnm.Print_Area','localSheetId':str(i)});a,b=s['printRange'].split(':');e.text=f"'{s['name']}'!${a[0]}${a[1:]}:${b[0]}${b[1:]}"
    calc=wb.find(q('calcPr'))
    if calc is None:calc=ET.SubElement(wb,q('calcPr'))
    calc.set('calcMode','auto');calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
    for view in wb.findall('.//'+q('workbookView')):view.set('activeTab','0');view.set('firstSheet','0')
    parts['xl/workbook.xml']=ET.tostring(wb,encoding='utf-8',xml_declaration=True)
    if 'docProps/core.xml' in parts:
        core=ET.fromstring(parts['docProps/core.xml'])
        dc='http://purl.org/dc/elements/1.1/'
        for tag,value in [('creator','Ridla A. Gani'),('title','NickelOps - Ore Delivery Reconciliation and Scenario Analysis'),('subject','Synthetic educational portfolio'),('description','Fictional delivery data and illustrative assumptions. Independent portfolio work; no real mining operations or financial impact.')]:
            element=core.find('{'+dc+'}'+tag)
            if element is None:element=ET.SubElement(core,'{'+dc+'}'+tag)
            element.text=value
        parts['docProps/core.xml']=ET.tostring(core,encoding='utf-8',xml_declaration=True)
    chartns='http://schemas.openxmlformats.org/drawingml/2006/chart'
    for name in list(parts):
        if '/charts/chart' in name and name.endswith('.xml'):
            chart=ET.fromstring(parts[name])
            for axis in chart.findall('.//{'+chartns+'}valAx'):
                scaling=axis.find('{'+chartns+'}scaling')
                if scaling is not None:
                    minimum=scaling.find('{'+chartns+'}min')
                    if minimum is None:minimum=ET.SubElement(scaling,'{'+chartns+'}min')
                    minimum.set('val','0')
            parts[name]=ET.tostring(chart,encoding='utf-8',xml_declaration=True)
    style_root=ET.fromstring(parts['xl/styles.xml']);cellxfs=style_root.find(q('cellXfs'))
    style_cache={}
    def number_style(cell,fmtid):
        old=int(cell.get('s','0'));key=(old,fmtid)
        if cellxfs[old].get('numFmtId','0')==str(fmtid):return
        if key not in style_cache:
            xf=copy.deepcopy(cellxfs[old]);xf.set('numFmtId',str(fmtid));xf.set('applyNumberFormat','1');style_cache[key]=len(cellxfs);cellxfs.append(xf)
        cell.set('s',str(style_cache[key]))
    # Preserve namespace prefixes on untouched parts. Only the standard worksheet metadata changes.
    for i,s in enumerate(spec['sheets'],1):
        p=f'xl/worksheets/sheet{i}.xml'; root=ET.fromstring(parts[p]);sv=root.find(q('sheetViews'))
        if s['name']=='DATA REVIEW':
            for address,text in [('A2','Review'),('A3','SYNTHETIC DATA')]:
                cell=root.find('.//'+q('c')+'[@r="'+address+'"]')
                for child in list(cell):cell.remove(child)
                cell.set('t','inlineStr');ET.SubElement(ET.SubElement(cell,q('is')),q('t')).text=text
        if s['name']=='QA CHECKS':
            for cell in root.findall('.//'+q('c')):
                address=cell.get('r');letter=''.join(c for c in address if c.isalpha());row=int(''.join(c for c in address if c.isdigit()))
                if letter=='D' and 7<=row<=19:number_style(cell,0)
                if letter in ['B','C'] and (7<=row<=11 or 18<=row<=19):number_style(cell,1)
                if address=='B11' and cell.find(q('f')) is not None and cell.find(q('f')).text.startswith('COUNTIFS'):
                    report=local('reports/qa_compatibility_test.json')
                    if report.exists():
                        evidence=json.loads(report.read_text(encoding='utf-8'))
                        if not evidence['passed']:raise RuntimeError('QA compatibility check did not pass')
                        cell.find(q('f')).text=evidence['formula'].lstrip('=');cell.set('t','n');cell.find(q('v')).text=str(evidence['actual'])
                        result=root.find('.//'+q('c')+'[@r="E11"]');result.set('t','str');result.find(q('v')).text='PASS'
        if sv is not None:
            for view in sv:
                view.set('tabSelected','1' if i==1 else '0');view.set('zoomScale','90' if s['name']=='DASHBOARD' else '85')
                for pane in list(view.findall(q('pane'))):view.remove(pane)
                if s['freezeRows'] or s['freezeColumns']:
                    from build_spec import col
                    attrs={'state':'frozen','topLeftCell':f"{col(s['freezeColumns']+1)}{s['freezeRows']+1}",'activePane':'bottomRight' if s['freezeColumns'] and s['freezeRows'] else ('bottomLeft' if s['freezeRows'] else 'topRight')}
                    if s['freezeRows']:attrs['ySplit']=str(s['freezeRows'])
                    if s['freezeColumns']:attrs['xSplit']=str(s['freezeColumns'])
                    view.insert(0,ET.Element(q('pane'),attrs))
                    for selection in list(view.findall(q('selection'))):view.remove(selection)
                    ET.SubElement(view,q('selection'),{'pane':attrs['activePane'],'activeCell':attrs['topLeftCell'],'sqref':attrs['topLeftCell']})
        for dv in root.findall('.//'+q('dataValidation')):
            formula=dv.find(q('formula1'))
            if formula is not None:
                if dv.get('sqref')=='F6' and s['name']=='DASHBOARD':formula.text='SupplierFilterChoices'
                if dv.get('sqref')=='H6' and s['name']=='DASHBOARD':formula.text='StockpileFilterChoices'
                if dv.get('sqref')=='AQ7:AQ46':formula.text='SourceRowIDs'
            dv.set('showErrorMessage','1');dv.set('errorStyle','stop');dv.set('errorTitle','Check input');dv.set('error','Use the listed choices or numeric limits. Formula gates also validate pasted inputs.')
        if s['name']=='RAW DELIVERIES':
            for old in root.findall(q('sheetProtection')):root.remove(old)
            prot=ET.Element(q('sheetProtection'),{'sheet':'1','objects':'1','scenarios':'1','autoFilter':'0','sort':'1'})
            data=root.find(q('sheetData'));root.insert(list(root).index(data)+1,prot)
        # Standard ordering: pageMargins, pageSetup, headerFooter, drawing/tableParts.
        for tag in ['printOptions','pageMargins','pageSetup','headerFooter']:
            for old in root.findall(q(tag)):root.remove(old)
        idx=len(root)
        for j,child in enumerate(root):
            if child.tag in [q(x) for x in ['rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst']]:idx=j;break
        margins=ET.Element(q('pageMargins'),dict(left='0.3',right='0.3',top='0.4',bottom='0.4',header='0.2',footer='0.2'))
        setup=ET.Element(q('pageSetup'),dict(paperSize='9',orientation='landscape',fitToWidth='1',fitToHeight='1'))
        footer=ET.Element(q('headerFooter'));ET.SubElement(footer,q('oddFooter')).text='&LSYNTHETIC EDUCATIONAL DATA&CRidla A. Gani&RPage &P'
        for node in [margins,setup,footer]:root.insert(idx,node);idx+=1
        prop=root.find(q('sheetPr'))
        if prop is None:prop=ET.Element(q('sheetPr'));root.insert(0,prop)
        ps=prop.find(q('pageSetUpPr'))
        if ps is None:ps=ET.SubElement(prop,q('pageSetUpPr'))
        ps.set('fitToPage','1')
        parts[p]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    cellxfs.set('count',str(len(cellxfs)));parts['xl/styles.xml']=ET.tostring(style_root,encoding='utf-8',xml_declaration=True)
    temp=local('tmp/finalized.xlsx')
    with zipfile.ZipFile(temp,'w',zipfile.ZIP_DEFLATED) as z:
        for name,content in parts.items():z.writestr(name,content)
    temp.replace(book)
    print('Finalized native views, named validation lists, print settings and raw-sheet protection.')

if __name__=='__main__':main()
