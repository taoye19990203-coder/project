"""Excel → 审计整改报告。复用模板 XML，保持未修改的 DOCX 部件原样。"""
from __future__ import annotations
import argparse
import io
import json
import posixpath
import re
from collections import Counter, OrderedDict
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED, BadZipFile
from lxml import etree as E
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.utils.datetime import from_excel
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

CHAPTERS = ('内控审计整改情况详述', '滚动内审情况详述')
KINDS = ('管理提升', '风险提示', '设计', '执行')
STATUSES = ('已整改', '整改中', '未整改')
HEADERS = {
 'company':'缺陷发生实体名称', 'audit':'审计类别', 'number':'编号',
 'process':'涉及业务流程', 'title':'问题简述', 'description':'问题描述',
 'kind':'问题类别', 'level':'问题等级', 'result':'实际整改措施及结果',
 'status':'是否完成整改', 'plan':'未整改原因说明及后续整改计划',
 'deadline':'整改时间节点', 'department':'整改责任部门',
}

class InputError(ValueError):
    pass

def safe_package(data):
    try:
        with ZipFile(io.BytesIO(data)) as z:
            if sum(x.file_size for x in z.infolist()) > 100_000_000 or len(z.infolist()) > 5000:
                raise InputError('文件解压后过大，请使用小于100MB的标准 Office 文件。')
    except (OSError, ValueError, BadZipFile) as e:
        raise InputError('文件不是有效的 xlsx/docx 文件。') from e

def text(v):
    return '' if v is None else str(v).strip()

def read_rows(data, company=''):
    safe_package(data)
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        raise InputError('无法读取Excel，请使用未加密的 .xlsx 文件。') from e
    candidates=[]
    for ws in wb:
        for r in range(1,min(ws.max_row,20)+1):
            names={text(ws.cell(r,c).value) for c in range(1,ws.max_column+1)}
            if {'审计类别','问题简述','问题描述'} <= names:
                candidates.append((ws,r)); break
    if len(candidates)!=1:
        raise InputError('需要唯一一张包含“审计类别、问题简述、问题描述”表头的明细工作表。')
    ws, header = candidates[0]
    columns={}
    for key,name in HEADERS.items():
        matches=[c for r in range(header,header+2) for c in range(1,ws.max_column+1) if text(ws.cell(r,c).value)==name]
        if not matches:
            raise InputError(f'缺少必要列：{name}（需保留审计项目组反馈栏）。')
        columns[key]=matches[-1]  # 问题等级取项目组反馈列；滚动审计改用原问题等级。
    original_level = next(c for c in range(1,ws.max_column+1) if text(ws.cell(header,c).value)=='问题等级')
    rows=[]
    warnings=[]
    for r in range(header+2,ws.max_row+1):
        row={k:ws.cell(r,c).value for k,c in columns.items()}
        if not any(text(row[k]) for k in ('title','audit','number')):
            continue
        if not text(row['title']):
            raise InputError(f'第{r}行缺少问题标题，不能跳过该缺陷。')
        for k in row:
            if k!='deadline': row[k]=text(row[k])
        row['source_row']=r
        audit=row['audit']
        if '滚动' in audit: row['chapter']=3
        elif '回头看' in audit or '内控审计' in audit: row['chapter']=2
        else: raise InputError(f'第{r}行审计类别无法识别：{audit}')
        if company and row['company']!=company: continue
        if not row['company'] or not row['process']:
            raise InputError(f'第{r}行实体名称或业务流程为空。')
        row['kind']=row['kind'].replace('内控','').replace('缺陷','')
        if row['kind'] not in KINDS: raise InputError(f'第{r}行问题类别无法识别：{row["kind"]}')
        if row['chapter']==3: row['level']=text(ws.cell(r,original_level).value)
        row['level']=row['level'].replace('缺陷','')
        if row['level'] not in ('重大','重要','一般'):
            raise InputError(f'第{r}行问题等级无法识别：{row["level"]}')
        if row['chapter']==2 and row['status'] not in STATUSES:
            raise InputError(f'第{r}行整改状态需为已整改、整改中或未整改。')
        v=row['deadline']
        if isinstance(v,(float,int)) and 20000 < v < 100000: v=from_excel(v,wb.epoch)
        if isinstance(v,(datetime,date)): v=f'{v.year}年{v.month}月{v.day}日'
        row['deadline']=text(v)
        required=['description','plan','department','deadline']
        if row['chapter']==2: required+=['result']
        for k in required:
            if not row[k]:
                row[k]='待补充'
                warnings.append(f'第{r}行“{HEADERS[k]}”为空，Word中标记为“待补充”。')
        rows.append(row)
    if not rows: raise InputError('未找到可生成的缺陷记录，请检查实体筛选或表格。')
    companies=list(dict.fromkeys(x['company'] for x in rows))
    if len(companies)!=1: raise InputError('Excel含多个实体，请填写本次报告的实体名称：'+ '、'.join(companies))
    for chapter in (2,3):
        selected=[r for r in rows if r['chapter']==chapter]
        ids=[r['number'] for r in selected if r['number']]
        if len(ids)!=len(set(ids)): warnings.append(f'第{chapter}章源表编号有重复；输出按业务流程重新编号。')
    return rows,warnings

def percent(a,b):
    from decimal import Decimal, ROUND_HALF_UP
    return str((Decimal(a*100)/Decimal(b)).quantize(Decimal('1'),rounding=ROUND_HALF_UP))+'%' if b else '不适用'

def stats(rows,warnings):
    return {'company':rows[0]['company'],'warnings':warnings,'chapters':[
        {'chapter':n,'count':len(part),'processes':list(dict.fromkeys(r['process'] for r in part)),
         'kinds':dict(Counter(r['kind'] for r in part)),
         'statuses':dict(Counter(r['status'] for r in part)) if n==2 else {}}
        for n in (2,3) for part in [[r for r in rows if r['chapter']==n]] ]}

def xml(data):
    return E.fromstring(data,parser=E.XMLParser(resolve_entities=False,no_network=True))

def dump(root):
    return E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)

def ptext(p):
    return ''.join(n.text or '' if n.tag==qn('w:t') else '\t' if n.tag==qn('w:tab') else '\n' for n in p.iter() if n.tag in (qn('w:t'),qn('w:tab'),qn('w:br'),qn('w:cr')))

def setnum(p,num_id):
    num=p.find('./'+qn('w:pPr')+'/'+qn('w:numPr')+'/'+qn('w:numId'))
    if num is None: raise InputError('模板标题缺少自动编号设置。')
    num.set(qn('w:val'),str(num_id))

def new_num(numbering,source,start=1):
    old=numbering.find(f'{qn("w:num")}[@{qn("w:numId")}="{source}"]')
    if old is None: raise InputError('模板编号定义不完整。')
    n=deepcopy(old)
    for attr in list(n.attrib):
        if attr!=qn('w:numId'): del n.attrib[attr]
    abstract_id=old.find(qn('w:abstractNumId')).get(qn('w:val'))
    abstract=deepcopy(numbering.find(f'{qn("w:abstractNum")}[@{qn("w:abstractNumId")}="{abstract_id}"]'))
    new_abstract=max(int(x.get(qn('w:abstractNumId'))) for x in numbering.findall(qn('w:abstractNum')))+1
    abstract.set(qn('w:abstractNumId'),str(new_abstract))
    for child in list(abstract):
        if child.tag in (qn('w:nsid'),qn('w:tmpl')): abstract.remove(child)
    abstract.find(f'{qn("w:lvl")}[@{qn("w:ilvl")}="0"]/{qn("w:start")}').set(qn('w:val'),str(start))
    numbering.insert(0,abstract)
    n.find(qn('w:abstractNumId')).set(qn('w:val'),str(new_abstract))
    value=max(int(x.get(qn('w:numId'))) for x in numbering.findall(qn('w:num')))+1
    n.set(qn('w:numId'),str(value))
    for x in list(n):
        if x.tag==qn('w:lvlOverride'): n.remove(x)
    lvl=OxmlElement('w:lvlOverride');lvl.set(qn('w:ilvl'),'0')
    s=OxmlElement('w:startOverride');s.set(qn('w:val'),str(start));lvl.append(s);n.append(lvl)
    numbering.append(n)
    return value

def paragraph(proto,value,label=''):
    """复制段落和字符属性；字段标签与正文分别沿用模板的字符样式。"""
    p=deepcopy(proto)
    runs=proto.findall(qn('w:r'))
    base=next((r for r in runs if r.find(qn('w:t')) is not None),None)
    bodyrun=base
    if label:
        offset=0
        for r in runs:
            t=''.join(n.text or '' for n in r.iter(qn('w:t')))
            if offset>=len(label) and t: bodyrun=r;break
            offset+=len(t)
    for child in list(p):
        if child.tag!=qn('w:pPr'): p.remove(child)
    def add(value,source,bold=None):
        r=OxmlElement('w:r')
        if source is not None and source.find(qn('w:rPr')) is not None: r.append(deepcopy(source.find(qn('w:rPr'))))
        if bold is not None:
            pr=r.find(qn('w:rPr'))
            if pr is None: pr=OxmlElement('w:rPr');r.insert(0,pr)
            b=pr.find(qn('w:b'))
            if b is None: b=OxmlElement('w:b');pr.append(b)
            b.set(qn('w:val'),'1' if bold else '0')
        t=OxmlElement('w:t');t.set(qn('xml:space'),'preserve');t.text=value;r.append(t);p.append(r)
    if label: add(label,base,True)
    add(value,bodyrun,False if label else None)
    return p

def field(proto,label,value):
    lines=value.replace('\r\n','\n').replace('\r','\n').split('\n')
    out=[paragraph(proto,lines[0],label)]
    for line in lines[1:]:
        if line.strip():
            p=paragraph(proto,line)
            # continuation uses the value's font without the label's bold styling
            for b in p.findall('.//'+qn('w:b')): b.set(qn('w:val'),'0')
            out.append(p)
    return out

def summary(part,company,period,rolling=False):
    if not part: return f'本次{company}未列示'+('滚动内审' if rolling else '回头看')+'缺陷。'
    kinds=Counter(r['kind'] for r in part); processes=list(dict.fromkeys(r['process'] for r in part))
    s=f'在{period}对{company}进行的'+('滚动内审工作中' if rolling else '内控审计中')+f'，审计项目组共发现{len(part)}项缺陷，涉及'+ '、'.join(processes)+f'等{len(processes)}项业务流程。其中，管理提升{kinds["管理提升"]}项（不纳入整改考核）、风险提示{kinds["风险提示"]}项（不纳入整改考核）、设计缺陷{kinds["设计"]}项、执行缺陷{kinds["执行"]}项。'
    if not rolling:
        c=Counter(r['status'] for r in part)
        s+=f'其中{c["已整改"]}项问题已完成整改，{c["整改中"]}项问题仍处于整改中状态，{c["未整改"]}项问题未整改。'
    return s+'详情见附件“'+('内控问题及整改建议汇总表' if rolling else '内控缺陷问题汇总表')+'”。'

def update_chart(parts,name,rows,by_kind,rolling):
    ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
    cns='{'+ns['c']+'}'
    root=xml(parts[name])
    cats=['管理提升','风险提示','内控设计','内控执行'] if by_kind else list(dict.fromkeys(r['process'] for r in rows))
    if not cats: cats=['无缺陷']
    relname=posixpath.join(posixpath.dirname(name),'_rels',posixpath.basename(name)+'.rels')
    rels=xml(parts[relname]);target=next(x.get('Target') for x in rels if x.get('Type').endswith('/package'))
    embedded=posixpath.normpath(posixpath.join(posixpath.dirname(name),target))
    wb=load_workbook(io.BytesIO(parts[embedded]));ws=wb.active
    if ws.max_row>1: ws.delete_rows(2,ws.max_row-1)
    for idx,ser in enumerate(root.findall('.//c:ser',ns),2):
        series=ser.find('.//c:tx//c:v',ns).text
        numbers=[]
        for cat in cats:
            selected=[r for r in rows if (r['kind']==cat.replace('内控','') if by_kind else r['process']==cat)]
            numbers.append(sum(r['level' if rolling else 'status']==series for r in selected))
        for tag,cache_tag,values in [('cat','strCache',cats),('val','numCache',numbers)]:
            ref=ser.find('c:'+tag,ns)[0]
            cache=ref.find('c:'+cache_tag,ns)
            for child in list(cache):
                if child.tag!=cns+'formatCode': cache.remove(child)
            E.SubElement(cache,cns+'ptCount',val=str(len(values)))
            for j,v in enumerate(values): E.SubElement(E.SubElement(cache,cns+'pt',idx=str(j)),cns+'v').text=str(v)
            col='A' if tag=='cat' else get_column_letter(idx)
            ref.find('c:f',ns).text=f"'{ws.title}'!${col}$2:${col}${len(cats)+1}"
        for j,(cat,value) in enumerate(zip(cats,numbers),2): ws.cell(j,1,cat);ws.cell(j,idx,value)
    for table in ws.tables.values(): table.ref=f'A1:D{len(cats)+1}'
    out=io.BytesIO();wb.save(out);parts[embedded]=out.getvalue();parts[name]=dump(root)

def generate(excel,template=None,mode='chapters',company='',original_period='',rolling_period=''):
    if mode not in ('chapters','full'): raise InputError('输出范围无效。')
    rows,warnings=read_rows(excel,company)
    company=rows[0]['company']
    if template is None: raise InputError('请选择本地 Word 报告模板（.docx）。程序不包含审计原文件。')
    safe_package(template)
    try:
        with ZipFile(io.BytesIO(template)) as z: parts={n:z.read(n) for n in z.namelist()}
        doc=xml(parts['word/document.xml']);body=doc.find(qn('w:body'));nodes=list(body)
        paras=[p for p in nodes if p.tag==qn('w:p')]
        def heading(title):
            choices=[p for p in paras if ptext(p)==title and p.find('./'+qn('w:pPr')+'/'+qn('w:outlineLvl')) is not None]
            if len(choices)!=1: raise InputError(f'模板中未找到唯一章节标题：{title}')
            return choices[0]
        h2,h3=map(heading,CHAPTERS)
        i2,i3=nodes.index(h2),nodes.index(h3)
        end=next(i for i in range(i3+1,len(nodes)) if nodes[i].tag==qn('w:p') and ptext(nodes[i]).startswith('附件一'))
        if not i2<i3<end: raise InputError('模板章节顺序异常。')
        two=[p for p in nodes[i2:i3] if p.tag==qn('w:p')]
        three=[p for p in nodes[i3:end] if p.tag==qn('w:p')]
        def find(seq,prefix): return next(p for p in seq if ptext(p).startswith(prefix))
        sub2=next(p for p in two if p.find('./'+qn('w:pPr')+'/'+qn('w:outlineLvl')) is not None and p is not h2)
        sub3=next(p for p in three if p.find('./'+qn('w:pPr')+'/'+qn('w:outlineLvl')) is not None and p is not h3)
        numbering=xml(parts['word/numbering.xml'])
        def numid(p):return p.find('./'+qn('w:pPr')+'/'+qn('w:numPr')+'/'+qn('w:numId')).get(qn('w:val'))
        chapter_num=new_num(numbering,numid(h2),2) if mode=='chapters' else None
        if not original_period:
            original_period=re.search(r'在(.+?)对',ptext(two[1])).group(1)
        if not rolling_period:
            rolling_period=re.search(r'在(.+?)对',ptext(three[1])).group(1)
        made=[]
        for chapter,seq,h,sub,period in [(2,two,h2,sub2,original_period),(3,three,h3,sub3,rolling_period)]:
            selected=[r for r in rows if r['chapter']==chapter]
            title=deepcopy(h)
            if chapter_num is not None:setnum(title,chapter_num)
            made.append(title);made.append(paragraph(seq[1],summary(selected,company,period,chapter==3)))
            if chapter==2:
                c=Counter(r['status'] for r in selected);assessed=[r for r in selected if r['kind'] in ('设计','执行')]
                completed=sum(r['status']=='已整改' for r in assessed)
                made.append(paragraph(two[2],f'{company}整体整改完成率{percent(c["已整改"],len(selected))}（仅考虑内控设计与内控执行缺陷，整改完成率为{percent(completed,len(assessed))}）。'))
                made.append(paragraph(two[3],ptext(two[3]).replace(re.search(r'对(.+?)进行',ptext(two[1])).group(1),company)))
            # chart drawings retain original placement, dimensions and all theme/style references
            for p in seq:
                if p.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart') is not None: made.append(deepcopy(p))
            groups=OrderedDict()
            for r in selected:groups.setdefault(r['process'],[]).append(r)
            subgroup_num=new_num(numbering,numid(sub))
            title_proto=next(p for p in seq if re.match(r'^1[.．]',ptext(p)))
            for process,items in groups.items():
                p=paragraph(sub,process);setnum(p,subgroup_num);made.append(p)
                if chapter==2:
                    c=Counter(r['status'] for r in items)
                    made.append(paragraph(find(two,'已整改'),f'已整改{c["已整改"]}项，整改完成率{percent(c["已整改"],len(items))}；整改中{c["整改中"]}项；未整改{c["未整改"]}项'))
                for no,row in enumerate(items,1):
                    made.append(paragraph(title_proto,f'{no}.{row["title"]}'))
                    if chapter==2:
                        fields=[('【问题简述】：',row['description'].splitlines()[0]),('【整改结果】：',row['result']),('【责任部门】：',row['department'].replace('\n','；')),('【原因说明及后续计划】：',row['plan']),('【整改时间】：',row['deadline']),('【整改情况】：',row['status'])]
                    else:
                        fields=[('【问题描述】：',row['description']),('【管理建议】：',row['plan']),('【问题类别】：',row['kind']+('缺陷' if row['kind'] in ('设计','执行') else '')),('【问题等级】：',row['level']+'缺陷'),('【责任部门】：',row['department'].replace('\n','；')),('【整改时间】：',row['deadline'])]
                    for label,value in fields:made.extend(field(find(seq,label),label,value))
                    made.append(paragraph(title_proto,''))
        if mode=='full':
            for p in nodes[i2:end]: body.remove(p)
            for offset,p in enumerate(made):body.insert(i2+offset,p)
            warnings.append('完整报告中仅更新第二、三章；封面、第一章、目录和附件保留模板原文，请按本次审计情况复核。')
        else:
            section=next((n.find('.//'+qn('w:sectPr')) for n in nodes[i2:] if n.find('.//'+qn('w:sectPr')) is not None),body.find(qn('w:sectPr')))
            sect=deepcopy(section)
            # Section header/footer references are inherited from preceding sections.
            inherited={}
            for prior in doc.iter(qn('w:sectPr')):
                for ref in prior:
                    if ref.tag in (qn('w:headerReference'),qn('w:footerReference')):
                        inherited[(ref.tag,ref.get(qn('w:type')))]=ref
                if prior is section: break
            for key,ref in inherited.items():
                if not any(x.tag==key[0] and x.get(qn('w:type'))==key[1] for x in sect):sect.insert(0,deepcopy(ref))
            for p in list(body):body.remove(p)
            body.extend(made)
            if sect is not None:body.append(sect)
        parts['word/document.xml']=dump(doc);parts['word/numbering.xml']=dump(numbering)
        # Resolve chart references from the supplied template, not assumptions about file names.
        rels=xml(parts['word/_rels/document.xml.rels'])
        relmap={r.get('Id'):r.get('Target') for r in rels}
        for chapter,seq in [(2,two),(3,three)]:
            charts=[c for p in seq for c in p.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')]
            if len(charts)!=2:raise InputError('模板每章需保留原有两张统计图。')
            for j,c in enumerate(charts):
                name=posixpath.normpath(posixpath.join('word',relmap[c.get(qn('r:id'))]))
                update_chart(parts,name,[r for r in rows if r['chapter']==chapter],j==0,chapter==3)
        # Ask Word to refresh fields when opening, preserving other settings.
        settings=xml(parts['word/settings.xml']);update=settings.find(qn('w:updateFields'))
        if update is None:update=OxmlElement('w:updateFields');settings.append(update)
        update.set(qn('w:val'),'true');parts['word/settings.xml']=dump(settings)
    except InputError: raise
    except Exception as e: raise InputError('Word模板结构不兼容，请选择与原报告章节结构一致的 Word 模板。') from e
    out=io.BytesIO()
    with ZipFile(out,'w',ZIP_DEFLATED) as z:
        for n,b in parts.items():z.writestr(n,b)
    info=stats(rows,warnings);info.update(original_period=original_period,rolling_period=rolling_period,mode=mode)
    return out.getvalue(),info

def main():
    p=argparse.ArgumentParser(description='按原报告格式生成审计整改报告第二、三章')
    p.add_argument('excel',type=Path);p.add_argument('-o','--output',type=Path,default=Path('审计整改报告_第二三章.docx'))
    p.add_argument('--template',type=Path,required=True);p.add_argument('--mode',choices=['chapters','full'],default='chapters')
    p.add_argument('--company',default='');p.add_argument('--original-period',default='');p.add_argument('--rolling-period',default='')
    a=p.parse_args()
    try:
        if a.output.resolve() in [a.excel.resolve(),a.template.resolve()]:raise InputError('输出文件不能覆盖输入文件。')
        data,info=generate(a.excel.read_bytes(),a.template.read_bytes(),a.mode,a.company,a.original_period,a.rolling_period)
        a.output.parent.mkdir(parents=True,exist_ok=True);a.output.write_bytes(data)
        a.output.with_suffix('.json').write_text(json.dumps(info,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(info,ensure_ascii=False,indent=2)); print('已生成：',a.output)
    except (InputError,OSError) as e:p.exit(1,str(e)+'\n')

if __name__=='__main__':main()
