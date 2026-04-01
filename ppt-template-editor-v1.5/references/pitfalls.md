# pitfalls.md — 污染清理 + 常见问题速查 + 血泪教训

## think-cell 污染清理

```python
import zipfile, re

# Step 1: 检测
with zipfile.ZipFile('/home/claude/template.pptx', 'r') as z:
    bad = ['THINKCELL','think-cell','TCLayout','oleObj','AlternateContent','custDataLst']
    for name in z.namelist():
        if not (name.endswith('.xml') or name.endswith('.rels')): continue
        c = z.read(name).decode('utf-8', errors='replace')
        hits = [k for k in bad if k.lower() in c.lower()]
        if hits: print(f"  {name}: {hits}")

# Step 2: 清理
with zipfile.ZipFile('/home/claude/template.pptx', 'r') as z:
    all_files = {name: z.read(name) for name in z.namelist()}

def clean_xml(content):
    def handle_alt(m):
        fb = re.search(r'<mc:Fallback[^>]*>(.*?)</mc:Fallback>', m.group(0), re.DOTALL)
        if fb:
            fc = fb.group(1).strip()
            if any(k in fc.lower() for k in ['oleobj','tclayout','think-cell']): return ''
            return fc
        return ''
    content = re.sub(r'<mc:AlternateContent[^>]*>.*?</mc:AlternateContent>', handle_alt, content, flags=re.DOTALL)
    content = re.sub(r'<p:oleObj[^>]*/>', '', content)
    content = re.sub(r'<p:oleObj[^>]*>.*?</p:oleObj>', '', content, flags=re.DOTALL)
    content = re.sub(r'\s*<p:custDataLst>.*?</p:custDataLst>', '', content, flags=re.DOTALL)
    return content

EMPTY_TAG = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
tag_count = xml_count = rels_count = 0

for name in all_files:
    if re.match(r'ppt/tags/tag\d+\.xml$', name):
        all_files[name] = EMPTY_TAG; tag_count += 1

for name in list(all_files.keys()):
    if not re.match(r'ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster)\d+\.xml$', name): continue
    c = all_files[name].decode('utf-8', errors='replace'); orig = c; c = clean_xml(c)
    if c != orig: all_files[name] = c.encode('utf-8'); xml_count += 1

for name in list(all_files.keys()):
    if not name.endswith('.rels'): continue
    c = all_files[name].decode('utf-8', errors='replace'); orig = c
    c = re.sub(r'<Relationship[^>]*(?:oleObject|TCLayout)[^/]*/>\s*', '', c)
    if c != orig: all_files[name] = c.encode('utf-8'); rels_count += 1

bins = [n for n in list(all_files.keys()) if re.match(r'ppt/embeddings/oleObject\d+\.bin$', n)]
for name in bins: del all_files[name]

ct = all_files['[Content_Types].xml'].decode('utf-8')
ct = re.sub(r'<Default Extension="bin"[^/]*/>', '', ct)
ct = re.sub(r'<(?:Default|Override)[^>]*oleObject[^/]*/>', '', ct)
all_files['[Content_Types].xml'] = ct.encode('utf-8')

for key in ['ppt/presentation.xml', 'docProps/app.xml']:
    if key in all_files:
        c = all_files[key].decode('utf-8')
        c = re.sub(r'\s*<p:custDataLst>.*?</p:custDataLst>', '', c, flags=re.DOTALL)
        c = re.sub(r'\s*<vt:lpstr>[^<]*(?:think-cell|TCLayout|OLE)[^<]*</vt:lpstr>', '', c, flags=re.IGNORECASE)
        all_files[key] = c.encode('utf-8')

with zipfile.ZipFile('/home/claude/template_clean.pptx', 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in sorted(all_files.items()): z.writestr(name, data)

from pptx import Presentation
prs = Presentation('/home/claude/template_clean.pptx')
print(f"✅ 清理完成: tag{tag_count} xml{xml_count} rels{rels_count} bin{len(bins)} | {len(prs.slides)}张")
```

## 污染位置速查表

| 位置 | 污染内容 | 正确处理 |
|------|---------|---------|
| ppt/tags/tag*.xml | THINKCELL... | **替换为空壳**（不删文件！） |
| ppt/slides/slide*.xml | custDataLst+oleObj+AlternateContent | 删除这些XML元素 |
| ppt/slides/_rels/*.rels | oleObject关系引用 | 删除对应Relationship行 |
| ppt/embeddings/oleObject*.bin | OLE二进制数据 | 删除文件+更新Content_Types |
| ppt/presentation.xml | custDataLst | 删除该元素 |
| docProps/app.xml | 嵌入OLE服务器+think-cell | 删除相关vt:lpstr行 |

## 常见问题速查

| 现象 | 根因 | 解决 |
|------|------|------|
| PowerPoint报"修复并删除" | broken references | 检查rels里是否有引用但文件不存在 |
| 英文残留（最常见） | 多run碎片，字符串替换静默失败 | map_slide找出，rewrite_sp整体重写 |
| 文字写进去但看不见 | fontRef→lt1(白色)，make_para()无显式颜色 | 改用make_para_colored('000000') |
| 标题内容上下颠倒 | 按sp索引号写，但sp[12].y>sp[13].y | survey()按y坐标确认顺序 |
| 图标与文字错位 | N个图标只写1个make_para段落 | N个图标=N个make_para() |
| 右侧/底部空白 | 未做坐标普查，漏写cx=5"的大文字框 | survey()找cx最大的sp |
| 大数字版式失调 | sz>1800在双栏版式里比例失调 | 大数字最大sz=1800 |
| p:pic断链 | 从其他slide复制spTree时带入了p:pic | 删除所有p:pic元素 |
| sldIdLst幻灯片数量虚报 | re.findall扫描整个pres.xml误匹配slideMaster | 只扫描sldIdLst内部的rId |
| 会话结束文件丢失 | /home/claude/是临时目录 | 质检通过后立即cp到/mnt/user-data/outputs/ |
