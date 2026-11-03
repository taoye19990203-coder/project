"""Synthetic template generated from scratch; contains no uploaded audit data."""
import io
from zipfile import ZipFile, ZIP_DEFLATED
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree as E
from openpyxl import Workbook

C='http://schemas.openxmlformats.org/drawingml/2006/chart'
REL='http://schemas.openxmlformats.org/package/2006/relationships'


def template_bytes():
    d=Document()
    def p(value,number=None,level=None):
        x=d.add_paragraph()
        x.paragraph_format.line_spacing=1.5
        x.paragraph_format.first_line_indent=Pt(24)
        r=x.add_run(value);r.font.name='仿宋';r.font.size=Pt(12)
        if number is not None:
            n=OxmlElement('w:numPr')
            for tag,val in [('ilvl',0),('numId',number)]:
                a=OxmlElement('w:'+tag);a.set(qn('w:val'),str(val));n.append(a)
            x._p.get_or_add_pPr().append(n)
        if level is not None:
            a=OxmlElement('w:outlineLvl');a.set(qn('w:val'),str(level));x._p.get_or_add_pPr().append(a)
        return x
    def field(label):
        x=p('');r=x.add_run(label);r.font.name='仿宋';r.font.size=Pt(12);r.bold=True
        r=x.add_run('测试内容');r.font.name='仿宋';r.font.size=Pt(12);r.bold=False
    p('第一章测试内容',1,0)
    for n,title in [(2,'内控审计整改情况详述'),(3,'滚动内审情况详述')]:
        p(title,1,0)
        p('在2025年1月对模板公司进行的内控审计中，共发现1项缺陷。')
        if n==2:
            p('模板公司整体整改完成率0%。')
            p('模板公司持续关注未完成事项。')
        for chartno in range((n-2)*2+1,(n-2)*2+3):
            x=p('')
            drawing=OxmlElement('w:drawing');inline=OxmlElement('wp:inline')
            graphic=E.SubElement(inline,'{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
            g=E.SubElement(graphic,'{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData',uri=C)
            chart=E.SubElement(g,'{'+C+'}chart');chart.set(qn('r:id'),'chart'+str(chartno))
            drawing.append(inline);x.add_run()._r.append(drawing)
        p('测试流程',6,1)
        if n==2:p('已整改0项，整改中1项，未整改0项')
        p('1.示例问题')
        labels=['问题简述','整改结果','责任部门','原因说明及后续计划','整改时间','整改情况'] if n==2 else ['问题描述','管理建议','问题类别','问题等级','责任部门','整改时间']
        for label in labels:field('【'+label+'】：')
    p('附件一：测试附件')
    b=io.BytesIO();d.save(b)
    with ZipFile(io.BytesIO(b.getvalue())) as z:parts={n:z.read(n) for n in z.namelist()}
    rels=E.fromstring(parts['word/_rels/document.xml.rels'])
    for i in range(1,5):
        E.SubElement(rels,'{'+REL+'}Relationship',Id='chart'+str(i),Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',Target=f'charts/chart{i}.xml')
        root=E.Element('{'+C+'}chartSpace',nsmap={'c':C})
        chart=E.SubElement(E.SubElement(root,'{'+C+'}chart'),'{'+C+'}plotArea')
        for name in (['已整改','整改中','未整改'] if i<3 else ['重大','重要','一般']):
            ser=E.SubElement(chart,'{'+C+'}ser')
            tx=E.SubElement(ser,'{'+C+'}tx');E.SubElement(tx,'{'+C+'}v').text=name
            for tag,cachetype in [('cat','strCache'),('val','numCache')]:
                ref=E.SubElement(E.SubElement(ser,'{'+C+'}'+tag),'{'+C+'}'+('strRef' if tag=='cat' else 'numRef'))
                E.SubElement(ref,'{'+C+'}f').text='Sheet1!$A$2:$A$2'
                E.SubElement(ref,'{'+C+'}'+cachetype)
        parts[f'word/charts/chart{i}.xml']=E.tostring(root)
        fname='Microsoft_Excel_Worksheet'+(str(i-1) if i>1 else '')+'.xlsx'
        r=E.Element('{'+REL+'}Relationships')
        E.SubElement(r,'{'+REL+'}Relationship',Id='data',Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',Target='../embeddings/'+fname)
        parts[f'word/charts/_rels/chart{i}.xml.rels']=E.tostring(r)
        w=Workbook();w.active.title='Sheet1';w.active.append(['类别','已整改','整改中','未整改'])
        out=io.BytesIO();w.save(out);parts['word/embeddings/'+fname]=out.getvalue()
    parts['word/_rels/document.xml.rels']=E.tostring(rels)
    types=E.fromstring(parts['[Content_Types].xml'])
    E.SubElement(types,'{http://schemas.openxmlformats.org/package/2006/content-types}Default',Extension='xlsx',ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    parts['[Content_Types].xml']=E.tostring(types)
    out=io.BytesIO()
    with ZipFile(out,'w',ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)
    return out.getvalue()
