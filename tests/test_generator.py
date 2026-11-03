import io
import json
import re
import threading
import unittest
import urllib.request
from http.server import ThreadingHTTPServer
from zipfile import ZipFile
from docx import Document
from docx.oxml.ns import qn
from lxml import etree as E
from openpyxl import Workbook, load_workbook
from generator import generate, read_rows, InputError, HEADERS, xml, ptext

from template_fixture import template_bytes
TEMPLATE=template_bytes()

def generate(*args, **kwargs):
    from generator import generate as actual
    return actual(*args, template=TEMPLATE, **kwargs)

def workbook(rows=None):
    w=Workbook();s=w.active
    headers=['缺陷发生实体名称','审计类别','编号','涉及业务流程','问题简述','问题描述','问题类别','问题等级','实际整改措施及结果','是否完成整改','问题等级','未整改原因说明及后续整改计划','整改时间节点','整改责任部门']
    s.append(['汇总表']);s.append(headers);s.append([None]*len(headers))
    for r in rows or [
        ['测试公司','2025年内控审计及2026年回头看','1.','测试流程','原问题','问题概述\n详细描述','执行','一般','整改结果\n结果续段','已整改','重要','不适用','不适用','不适用'],
        ['测试公司','2026年滚动审计','1.','新增流程','新问题','第一段\n第二段','设计','重大',None,'不适用','不适用','整改建议\n建议续段','2027-04-30','责任部门\n协助部门']]:s.append(r)
    b=io.BytesIO();w.save(b);return b.getvalue()


def parts(data):
    with ZipFile(io.BytesIO(data)) as z:return {n:z.read(n) for n in z.namelist()}


class GeneratorTests(unittest.TestCase):
    def test_template_is_required(self):
        from generator import generate as actual
        with self.assertRaisesRegex(InputError, '请选择本地 Word'):
            actual(workbook())

    def test_fields_format_numbering_and_charts(self):
        data,info=generate(workbook(),original_period='2025年1月',rolling_period='2027年2月')
        d=Document(io.BytesIO(data));texts=[p.text for p in d.paragraphs]
        self.assertEqual([x['count'] for x in info['chapters']],[1,1])
        self.assertTrue(any('2027年2月' in t for t in texts))
        self.assertIn('【问题简述】：问题概述',texts)
        self.assertIn('【问题等级】：重大缺陷',texts) # 滚动审计使用原问题等级，而非反馈栏的不适用
        self.assertIn('【整改结果】：整改结果',texts)
        self.assertIn('结果续段',texts)
        self.assertIn('第二段',texts)
        self.assertIn('【责任部门】：责任部门；协助部门',texts)
        p=next(p for p in d.paragraphs if p.text.startswith('【整改结果】'))
        template=Document(io.BytesIO(TEMPLATE))
        original=next(p for p in template.paragraphs if p.text.startswith('【整改结果】'))
        self.assertEqual(E.tostring(p._p.pPr),E.tostring(original._p.pPr))
        self.assertTrue(p.runs[0].bold);self.assertFalse(p.runs[1].bold)
        self.assertEqual(p.runs[0].font.name,'仿宋');self.assertEqual(p.runs[0].font.size.pt,12)
        z=parts(data);num=xml(z['word/numbering.xml'])
        h=next(p for p in d.paragraphs if p.text=='内控审计整改情况详述')
        nid=h._p.xpath('./w:pPr/w:numPr/w:numId/@w:val')[0]
        n=num.find(f'{qn("w:num")}[@{qn("w:numId")}="{nid}"]')
        self.assertEqual(n.find('.//'+qn('w:startOverride')).get(qn('w:val')),'2')
        ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
        c=xml(z['word/charts/chart2.xml'])
        self.assertEqual(c.xpath('//c:ser[1]/c:cat//c:pt/c:v/text()',namespaces=ns),['测试流程'])
        self.assertEqual(c.xpath('//c:ser[1]/c:val//c:pt/c:v/text()',namespaces=ns),['1'])
        w=load_workbook(io.BytesIO(z['word/embeddings/Microsoft_Excel_Worksheet1.xlsx']),data_only=True)
        self.assertEqual(w.active.cell(2,1).value,'测试流程');self.assertEqual(w.active.cell(2,2).value,1)
        self.assertFalse(any('模板公司' in t for t in texts))

    def test_full_preserves_outside_chapters_and_parts(self):
        data,info=generate(workbook(),mode='full');a=parts(TEMPLATE);b=parts(data)
        def outside(doc):
            nodes=list(xml(doc).find(qn('w:body')))
            start=next(i for i,n in enumerate(nodes) if ptext(n)=='内控审计整改情况详述')
            end=next(i for i,n in enumerate(nodes) if ptext(n).startswith('附件一'))
            return [E.tostring(n) for n in nodes[:start]+nodes[end:]]
        self.assertEqual(outside(a['word/document.xml']),outside(b['word/document.xml']))
        for name in a:
            if name in ('word/document.xml','word/settings.xml','word/numbering.xml') or re.match(r'word/charts/chart\d.xml$',name) or name.startswith('word/embeddings/'):continue
            self.assertEqual(a[name],b[name],name)
        self.assertTrue(info['warnings'])

    def test_missing_values_marked_not_invented(self):
        w=load_workbook(io.BytesIO(workbook()));s=w.active;s.cell(4,9).value=None
        b=io.BytesIO();w.save(b)
        data,info=generate(b.getvalue())
        self.assertIn('【整改结果】：待补充',[p.text for p in Document(io.BytesIO(data)).paragraphs])
        self.assertEqual(len(info['warnings']),1)

    def test_invalid_inputs_and_multiple_entities(self):
        with self.assertRaises(InputError):generate(b'invalid')
        w=load_workbook(io.BytesIO(workbook()));s=w.active;s.cell(4,10).value='不知道'
        b=io.BytesIO();w.save(b)
        with self.assertRaisesRegex(InputError,'整改状态'):generate(b.getvalue())
        s.cell(4,10).value='已整改';s.cell(5,1).value='另一公司';b=io.BytesIO();w.save(b)
        with self.assertRaisesRegex(InputError,'多个实体'):generate(b.getvalue())
        result,info=generate(b.getvalue(),company='另一公司')
        self.assertEqual([c['count'] for c in info['chapters']],[0,1])
        Document(io.BytesIO(result))

    def test_browser_upload_and_download(self):
        import app
        server=ThreadingHTTPServer(('127.0.0.1',0),app.Handler)
        worker=threading.Thread(target=server.serve_forever,daemon=True);worker.start()
        url=f'http://127.0.0.1:{server.server_port}'
        try:
            page=urllib.request.urlopen(url).read().decode()
            self.assertIn(app.TOKEN,page)
            boundary='----AuditTestBoundary'
            payload=(f'--{boundary}\r\nContent-Disposition: form-data; name="excel"; filename="input.xlsx"\r\nContent-Type: application/octet-stream\r\n\r\n').encode()+workbook()+f'\r\n--{boundary}--\r\n'.encode()
            payload=payload[:-len(f'--{boundary}--\r\n'.encode())]+(f'--{boundary}\r\nContent-Disposition: form-data; name="template"; filename="template.docx"\r\nContent-Type: application/octet-stream\r\n\r\n').encode()+TEMPLATE+f'\r\n--{boundary}--\r\n'.encode()
            for path in ['/analyze','/generate']:
                r=urllib.request.urlopen(urllib.request.Request(url+path,data=payload,headers={'Content-Type':f'multipart/form-data; boundary={boundary}','X-Request-Token':app.TOKEN}))
                data=r.read()
                if path=='/analyze':self.assertEqual(json.loads(data)['chapters'][0]['count'],1)
                else: self.assertEqual(len(Document(io.BytesIO(data)).inline_shapes),4)
        finally:server.shutdown();server.server_close();worker.join()

if __name__=='__main__':unittest.main()
