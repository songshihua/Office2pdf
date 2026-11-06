"""Local Word-to-PDF conversion; never modify the user's source document."""
from pathlib import Path
from tempfile import TemporaryDirectory
from contextlib import contextmanager
import uuid
import shutil
from zipfile import ZipFile
from io import BytesIO
import base64
import json
import os
import re
import subprocess
import sys
import ctypes
import xml.etree.ElementTree as ET
ET.register_namespace('', 'http://www.w3.org/2000/svg')
ET.register_namespace('xlink', 'http://www.w3.org/1999/xlink')

import fitz
from fontTools.ttLib import TTFont, TTCollection
from fontTools.pens.basePen import BasePen
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream, NumberObject
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtGui import QPdfWriter, QPainter, QPageSize, QPageLayout, QFontDatabase, QFont
from PyQt5.QtCore import QSizeF, QMarginsF, QRectF

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'svg': 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
}

def register_fonts(app):
    fonts = Path(os.environ.get('WINDIR', r'C:\Windows')) / 'Fonts'
    for name in ('msyh.ttc', 'msyhbd.ttc', 'simsun.ttc', 'simhei.ttf', 'times.ttf',
                 'timesbd.ttf', 'timesi.ttf', 'arial.ttf', 'arialbd.ttf', 'ariali.ttf',
                 'calibri.ttf', 'calibrib.ttf', 'cambria.ttc'):
        if (fonts / name).exists():
            QFontDatabase.addApplicationFont(str(fonts / name))
    app.setFont(QFont('Microsoft YaHei', 10))

def native_export(source, folder, log):
    def ps(s):
        return "'" + str(s).replace("'", "''") + "'"
    pdf, positions = folder / 'native.pdf', folder / 'positions.json'
    script = r'''
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding
function Invoke-Office([scriptblock]$Action) {
 for ($attempt=0; $attempt -lt 30; $attempt++) {
  try { return (& $Action) } catch {
   $busy=($_.Exception.ToString() -match '80010001|8001010A|RPC_E_CALL_REJECTED|RPC_E_SERVERCALL_RETRYLATER')
   if ($attempt -ge 29 -or -not $busy) { throw }
   Start-Sleep -Milliseconds 500
  }
 }
}
$app=$null; $doc=$null; $quit=$false; $stage='连接文字处理程序'
try {
 foreach ($prog in @('kwps.Application','Word.Application')) {
  try { $app=New-Object -ComObject $prog; $engine=$prog; break } catch { }
 }
 if ($null -eq $app) { throw '找不到可用的 WPS 或 Microsoft Word。请先安装并正常打开一次文字处理程序。' }
 $stage='读取文字处理程序状态'
 $initialCount=Invoke-Office { $app.Documents.Count }
 $quit=($initialCount -eq 0)
 if ($quit) { Invoke-Office { $app.Visible=$false } }
 $oldAlerts=Invoke-Office { $app.DisplayAlerts }
 Invoke-Office { $app.DisplayAlerts=0 }
 try { $oldSecurity=$app.AutomationSecurity; $app.AutomationSecurity=3 } catch { }
 $stage='打开文档'
 $doc=Invoke-Office { $app.Documents.Open(SOURCE,$false,$true) }
 $stage='计算分页'
 Invoke-Office { $doc.Repaginate() }
 $items=@()
 $stage='读取图片位置'
 $shapeCount=Invoke-Office { $doc.InlineShapes.Count }
 for($i=1;$i -le $shapeCount;$i++) {
  $item=Invoke-Office {
   $s=$doc.InlineShapes.Item($i);$r=$s.Range
   [pscustomobject]@{Index=$i; Width=$s.Width; Height=$s.Height; Page=$r.Information(3); X=$r.Information(5); Y=$r.Information(6)}
  }
  $items += $item
 }
 ConvertTo-Json -InputObject $items -Depth 4 | Set-Content -LiteralPath POSITIONS -Encoding UTF8
 $stage='导出 PDF'
 Invoke-Office { $doc.ExportAsFixedFormat(DEST,17) }
 Write-Output ('ENGINE:'+ $engine)
} catch {
 Write-Output ('ERROR:'+ $stage+'：'+$_.Exception.Message)
 exit 1
} finally {
 if ($null -ne $doc) { try { Invoke-Office { $doc.Close(0) } } catch { } }
 if ($null -ne $app) {
  try { $app.DisplayAlerts=$oldAlerts } catch { }
  try { if ($null -ne $oldSecurity) { $app.AutomationSecurity=$oldSecurity } } catch { }
  if ($quit) { try { Invoke-Office { $app.Quit() } } catch { } }
 }
}
'''.replace('SOURCE', ps(source)).replace('POSITIONS', ps(positions)).replace('DEST', ps(pdf))
    encoded = base64.b64encode(script.encode('utf-16le')).decode('ascii')
    command = [str(Path(os.environ.get('WINDIR', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe'),
               '-NoProfile', '-STA', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded]
    flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
    log('正在调用 WPS / Word 排版并导出，请稍候…')
    # Frozen Qt applications must not pass their private Qt/DLL search paths to
    # WPS, which ships its own Qt libraries and can otherwise reject COM calls.
    child_env=os.environ.copy()
    for key in list(child_env):
        if key.startswith(('QT_','QML')) or key in ('PYTHONHOME','PYTHONPATH'):
            child_env.pop(key,None)
    bundle=getattr(sys,'_MEIPASS',None)
    if bundle:
        child_env['PATH']=os.pathsep.join(p for p in child_env.get('PATH','').split(os.pathsep) if not p.lower().startswith(str(bundle).lower()))
        ctypes.windll.kernel32.SetDllDirectoryW(None)
    try:
        proc = subprocess.run(command, capture_output=True, timeout=600, creationflags=flags,env=child_env)
    finally:
        if bundle:ctypes.windll.kernel32.SetDllDirectoryW(str(bundle))
    stdout = proc.stdout.decode('utf-8', errors='replace')
    stderr = proc.stderr.decode('utf-8', errors='replace')
    if proc.returncode or not pdf.exists():
        detail=stdout.split('ERROR:',1)[-1] if 'ERROR:' in stdout else (stderr or stdout)
        raise RuntimeError('文字处理程序未完成导出。请先确认文件可以正常打开、没有密码或弹窗。\n' + detail[-1800:])
    engine = 'Microsoft Word' if 'Word.Application' in stdout else 'WPS'
    log(f'{engine} 排版导出完成。')
    records = json.loads(positions.read_text(encoding='utf-8-sig')) if positions.exists() else []
    return pdf, records if isinstance(records, list) else [records]

def source_figures(source, folder):
    if source.suffix.lower() not in ('.docx', '.docm'):
        return []
    with ZipFile(source) as z:
        # Normalize uncommon backslash entries without writing into the source.
        names = {n.replace('\\', '/'): n for n in z.namelist()}
        def read(n): return z.read(names[n])
        rel = {e.get('Id'): e.get('Target') for e in ET.fromstring(read('word/_rels/document.xml.rels'))}
        xml = ET.fromstring(read('word/document.xml'))
        figures = []
        for index, el in enumerate(xml.findall('.//w:drawing', NS), 1):
            inline = el.find('wp:inline', NS)
            svg = el.find('.//svg:svgBlip', NS)
            crop = el.find('.//a:srcRect', NS)
            record = {'svg': None, 'crop': {}, 'inline': inline is not None}
            if crop is not None:
                record['crop'] = {k: float(v) / 100000 for k, v in crop.attrib.items()}
            if svg is not None:
                rid = svg.get('{' + NS['r'] + '}embed')
                target = rel.get(rid, '')
                name = 'word/' + target.lstrip('/') if not target.startswith('/') else target.lstrip('/')
                if name in names:
                    dest = folder / f'figure-{index:03d}.svg'
                    dest.write_bytes(read(name));record['svg'] = dest
            figures.append(record)
        return figures

def compatible_svg(source):
    """Replace gradient color servers with their representative opaque color.
    Shape geometry, text, layout and data remain unchanged; do not trace bitmaps.
    """
    root = ET.parse(source).getroot()
    colors = {}
    def color(s):
        s = s.strip()
        if re.fullmatch(r'#[0-9a-fA-F]{6}', s):return tuple(int(s[i:i+2], 16) for i in (1,3,5))
        if re.fullmatch(r'#[0-9a-fA-F]{3}', s):return tuple(int(c*2, 16) for c in s[1:])
        nums = re.findall(r'\d+', s)
        return tuple(map(int, nums[:3])) if s.startswith('rgb(') and len(nums)>=3 else None
    for el in root.iter():
        if el.tag.split('}')[-1] not in ('linearGradient','radialGradient'):continue
        stops = []
        for stop in el:
            style = dict(re.findall(r'([\w-]+)\s*:\s*([^;]+)', stop.get('style','')))
            c = color(stop.get('stop-color', style.get('stop-color','#000000')))
            if c:stops.append(c)
        if stops:
            colors[el.get('id')] = '#'+''.join(f'{round((a+b)/2):02x}' for a,b in zip(stops[0], stops[-1]))
    for el in root.iter():
        for key,value in list(el.attrib.items()):
            if 'url(' in value:
                el.set(key, re.sub(r'url\(\s*[\"\']?#([^\)\"\']+)[\"\']?\s*\)',
                                   lambda m: colors.get(m.group(1),m.group(0)),value))
    return ET.tostring(root,encoding='utf-8')

def render_svg(source, destination, compatibility):
    data = compatible_svg(source) if compatibility else source.read_bytes()
    renderer = QSvgRenderer(data)
    if not renderer.isValid():raise ValueError('无法解析 SVG')
    box = renderer.viewBoxF();w,h = box.width(),box.height()
    if w <= 0 or h <= 0:raise ValueError('SVG 尺寸为空')
    writer = QPdfWriter(str(destination));writer.setResolution(72)
    writer.setPageSize(QPageSize(QSizeF(w,h),QPageSize.Point))
    writer.setPageMargins(QMarginsF(0,0,0,0),QPageLayout.Point)
    painter = QPainter(writer)
    if not painter.isActive():raise RuntimeError('不能创建矢量页面')
    try:renderer.render(painter,QRectF(0,0,w,h))
    finally:painter.end()
    return w,h

def restore_vectors(document, figures, positions, folder, compatibility, log, warnings):
    inline = [f for f in figures if f['inline']]
    if len(inline) != len(positions):
        if any(f['svg'] for f in figures):warnings.append('部分图形与排版对象无法一一对应，已保留排版引擎的导出结果。')
        return 0
    restored = 0;used = set()
    for index,(fig,pos) in enumerate(zip(inline,positions),1):
        if not fig['svg']:continue
        pno = int(pos['Page'])-1
        if not 0<=pno<len(document):continue
        page = document[pno];candidates=[]
        for im in page.get_image_info(xrefs=True):
            xref=im.get('xref',0)
            if not xref or (pno,xref) in used:continue
            box=fitz.Rect(im['bbox']);w,h=box.width,box.height
            expected_w,expected_h=float(pos['Width']),float(pos['Height'])
            if abs(w-expected_w)>max(9,expected_w*.06) or abs(h-expected_h)>max(9,expected_h*.06):continue
            score=abs(box.x0-float(pos['X']))+abs(box.y0-float(pos['Y']))+abs(w-expected_w)+abs(h-expected_h)
            candidates.append((score,im))
        if not candidates:
            # Word may have already exported the graphic as genuine vector content.
            log(f'图形 {index} 没有对应的整幅位图，保留原导出内容。');continue
        _,im=min(candidates,key=lambda x:x[0]);box=fitz.Rect(im['bbox'])
        try:
            target_file=folder/f'vector-{index:03d}.pdf'
            w,h=render_svg(fig['svg'],target_file,compatibility)
            with fitz.open(target_file) as vector:
                crop=fig['crop'];clip=fitz.Rect(w*crop.get('l',0),h*crop.get('t',0),w*(1-crop.get('r',0)),h*(1-crop.get('b',0)))
                if clip.is_empty or not fitz.Rect(0,0,w,h).contains(clip):raise ValueError('图片采用扩展裁切，保留原显示')
                # WPS raster capture adds ~1.5pt white padding; Word pictures need none.
                dw=box.width-float(pos['Width']);dh=box.height-float(pos['Height'])
                px=1.5 if 2.5<=dw<=4.5 else 0;py=1.5 if 2.5<=dh<=4.5 else 0
                target=fitz.Rect(box.x0+px,box.y0+py,box.x1-px,box.y1-py)
                # Insert first; if validation fails, keep the native image untouched.
                page.show_pdf_page(target,vector,0,clip=clip,keep_proportion=False,overlay=True)
                page.delete_image(im['xref']);used.add((pno,im['xref']));restored+=1
                log(f'已恢复第 {pno+1} 页的矢量图形 {index}。')
        except Exception as exc:
            warnings.append(f'图形 {index} 未作矢量替换：{exc}')
    return restored

class OutlinePen(BasePen):
    def __init__(self,glyphs,origin,direction,size,upem,height):
        super().__init__(glyphs);self.origin=origin;self.direction=direction;self.scale=size/upem;self.height=height;self.commands=[]
    def point(self,p):
        x,y=p;dx,dy=self.direction;ox,oy=self.origin;s=self.scale
        return f'{ox+s*(x*dx+y*dy):.7f} {self.height-oy+s*(-x*dy+y*dx):.7f}'
    def _moveTo(self,p):self.commands.append(self.point(p)+' m')
    def _lineTo(self,p):self.commands.append(self.point(p)+' l')
    def _curveToOne(self,a,b,c):self.commands.append(' '.join(self.point(p) for p in (a,b,c))+' c')
    def _closePath(self):self.commands.append('h')
    def _endPath(self):pass

def outline_math(source,destination,folder,log):
    def font_key(name):return re.sub(r'[^a-z0-9]','',name.split('+')[-1].lower())
    original=fitz.open(source);fonts={};page_fonts=[]
    for page in original:
        mapping={}
        for ft in page.get_fonts(full=True):
            key=font_key(ft[3])
            if not any(s in key for s in ('cambriamath','symbolmt','stix','latinmodernmath')):continue
            if ft[0] not in fonts:
                _,_,_,data=original.extract_font(ft[0])
                if not data:raise RuntimeError('公式字体未嵌入，无法安全转为轮廓。')
                if data[:4]==b'ttcf':
                    coll=TTCollection(BytesIO(data))
                    font=next((f for f in coll.fonts if 'math' in (f['name'].getDebugName(6) or '').lower()),coll.fonts[0])
                else:font=TTFont(BytesIO(data))
                fonts[ft[0]]=font
            mapping[ft[4]]=ft[0]
        page_fonts.append(mapping)
    if not fonts:
        original.close();Path(destination).write_bytes(Path(source).read_bytes());return 0
    reader=PdfReader(source);writer=PdfWriter();writer.clone_document_from_reader(reader)
    for pno,page in enumerate(writer.pages):
        if not page_fonts[pno]:continue
        content=ContentStream(page.get_contents(),writer);ops=[];current=None;mode=0;stack=[]
        for operands,operator in content.operations:
            if operator==b'q':stack.append((current,mode))
            elif operator==b'Q' and stack:current,mode=stack.pop()
            elif operator==b'Tf':current=str(operands[0]).lstrip('/')
            elif operator==b'Tr':mode=int(operands[0])
            if operator in (b'Tj',b'TJ',b"'",b'"') and current in page_fonts[pno]:
                ops.extend([([NumberObject(3)],b'Tr'),(operands,operator),([NumberObject(mode)],b'Tr')])
            else:ops.append((operands,operator))
        content.operations=ops;page.replace_contents(content)
    hidden=folder/'hidden-math.pdf'
    with hidden.open('wb') as out:writer.write(out)
    fixed=fitz.open(hidden);total=0
    for pno,page in enumerate(original):
        byname={font_key(ft[3]):fonts[ft[0]] for ft in page.get_fonts(full=True) if ft[0] in fonts}
        commands=['q'];count=0
        for span in page.get_texttrace():
            font=byname.get(font_key(span['font']))
            if font is None or span['type']!=0:continue
            if span['wmode']!=0:raise RuntimeError('不支持竖排公式，已停止生成以避免丢失字符。')
            glyphs=font.getGlyphSet();order=font.getGlyphOrder();upem=font['head'].unitsPerEm
            color=span['color'];commands.append(' '.join(str(v) for v in color)+(' rg' if len(color)==3 else ' g'))
            for _,gid,origin,_ in span['chars']:
                if gid<0:continue
                if gid>=len(order):raise RuntimeError('公式字形编号超出字体范围。')
                pen=OutlinePen(glyphs,origin,span['dir'],span['size'],upem,page.rect.height)
                glyphs[order[gid]].draw(pen)
                if pen.commands:commands.extend(pen.commands);commands.append('f');count+=1
        if count:
            commands.append('Q');p=fixed[pno];p.wrap_contents()
            xref=fixed.get_new_xref();fixed.update_object(xref,'<<>>');fixed.update_stream(xref,'\n'.join(commands).encode('ascii'))
            fixed.xref_set_key(p.xref,'Contents','['+' '.join(f'{x} 0 R' for x in p.get_contents()+[xref])+']')
            total+=count
    fixed.save(destination,garbage=4,deflate=True);fixed.close();original.close()
    log(f'公式兼容处理完成：{total} 个数学字形转为矢量轮廓。')
    return total

def unique_output(folder,stem):
    dest=folder/(stem+'.pdf');i=2
    while dest.exists():dest=folder/f'{stem}_{i}.pdf';i+=1
    return dest

@contextmanager
def conversion_workspace(output_dir):
    # Use inherited Windows permissions so the Office COM process can write here.
    base=Path(output_dir).resolve()
    folder=base/('.转换临时_'+uuid.uuid4().hex)
    folder.mkdir()
    try:yield folder
    finally:
        resolved=folder.resolve()
        if resolved.parent==base and resolved.name.startswith('.转换临时_'):
            shutil.rmtree(resolved,ignore_errors=True)

def convert(source,output_dir,compatibility=True,log=lambda _:None):
    source=Path(source).resolve();output_dir=Path(output_dir).resolve()
    if not source.is_file():raise FileNotFoundError('找不到选择的 Word 文件。')
    if source.suffix.lower() not in ('.docx','.doc','.docm'):raise ValueError('请选择 .docx、.doc 或 .docm 文件。')
    output_dir.mkdir(parents=True,exist_ok=True)
    warnings=[]
    with conversion_workspace(output_dir) as tmp:
        folder=Path(tmp);native,positions=native_export(source,folder,log)
        figures=source_figures(source,folder)
        document=fitz.open(native);pages=len(document)
        restored=restore_vectors(document,figures,positions,folder,compatibility,log,warnings)
        log(f'共 {pages} 页，恢复 {restored} 幅 SVG 矢量图。')
        vectors=folder/'vectors.pdf';document.save(vectors,garbage=4,deflate=True);document.close()
        prepared=folder/'prepared.pdf';glyphs=0
        if compatibility:glyphs=outline_math(vectors,prepared,folder,log)
        else:prepared.write_bytes(vectors.read_bytes())
        # Validate every page can render before publishing the result.
        log('正在检查全部页面…')
        with fitz.open(prepared) as check:
            if len(check)!=pages:raise RuntimeError('转换后的页数发生变化，未保存结果。')
            for page in check:
                if page.rect.is_empty:raise RuntimeError('发现空尺寸页面。')
                page.get_pixmap(matrix=fitz.Matrix(.25,.25),alpha=False)
        dest=unique_output(output_dir,source.stem+('_兼容版' if compatibility else '_高清版'))
        # Publish only after the complete conversion and rendering checks succeed.
        part=dest.with_suffix('.pdf.part')
        try:part.write_bytes(prepared.read_bytes());part.replace(dest)
        finally:
            if part.exists():part.unlink()
    return {'path':str(dest),'pages':pages,'vectors':restored,'math_glyphs':glyphs,'warnings':warnings}
