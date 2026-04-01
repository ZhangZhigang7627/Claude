# incremental.md — 增量修改（单页/几页更新）
<!-- v1.3 · 2026-03-31 · 加版本号，内容无变更 -->

## 定位目标页

```python
import re

def get_slide_map(unpacked_dir='unpacked'):
    with open(f'{unpacked_dir}/ppt/presentation.xml') as f: pres = f.read()
    with open(f'{unpacked_dir}/ppt/_rels/presentation.xml.rels') as f: rels = f.read()
    sld_list = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', pres, re.DOTALL).group(1)
    rids = re.findall(r'r:id="(rId\d+)"', sld_list)
    slide_map = {}
    for i, rid in enumerate(rids):
        m = re.search(rf'Id="{rid}"[^>]*Target="slides/(slide\d+\.xml)"', rels)
        if m: slide_map[i+1] = m.group(1)
    for page, fname in slide_map.items():
        with open(f'{unpacked_dir}/ppt/slides/{fname}') as f: c = f.read()
        spans = [(ms.start(),ms.end()) for ms in re.finditer(r'<p:sp>.*?</p:sp>',c,re.DOTALL)]
        title = next((t for s,e in spans[:2]
                      if (t:=''.join(re.findall(r'<a:t[^>]*>([^<]*)</a:t>',c[s:e])).strip()) and len(t)>2), '')
        print(f'  第{page:3d}页 → {fname}  {title[:40]}')
    return slide_map
```

## 四种操作

### A. 修改内容（最常见）

```python
import sys; sys.path.insert(0,'/home/claude')
from ppt_builder import *

slide_map = get_slide_map('unpacked')
target = 73  # 第5页对应的slideN数字

survey(target)   # ≥20sp时必做
map_slide(target)
c = load(target)
c = rewrite_sp(c, 0, make_para('新标题', bold=True, sz=2400))
save(target, c)
verify(target)
```

### B. 换版式

```python
import subprocess
# 1. 添加新源slide
r = subprocess.run(['python3', '/mnt/skills/public/pptx/scripts/add_slide.py', 'unpacked/', 'slideX.xml'],
                   capture_output=True, text=True)
# 2. 更新sldIdLst（用新rId替换旧页的rId）
# 3. 必须跑clean.py清理旧的孤立slide
subprocess.run(['python3', '/mnt/skills/public/pptx/scripts/clean.py', 'unpacked/'])
# 4. 写入新内容（同操作A）
```

### C. 插入新页

```python
import subprocess, re
r = subprocess.run(['python3', '/mnt/skills/public/pptx/scripts/add_slide.py', 'unpacked/', 'slideN.xml'],
                   capture_output=True, text=True)
new_rid = re.search(r'rId\d+', r.stdout).group(0)

# 在sldIdLst指定位置插入新的<p:sldId>
with open('unpacked/ppt/presentation.xml') as f: pres = f.read()
# ... 在目标位置插入 f'<p:sldId id="750" r:id="{new_rid}"/>'
with open('unpacked/ppt/presentation.xml', 'w') as f: f.write(pres)
# 插入不需要clean.py
```

### D. 删除页面

```python
import re, subprocess

slide_map = get_slide_map('unpacked')
target_file = slide_map[8]  # 删第8页

with open('unpacked/ppt/_rels/presentation.xml.rels') as f: rels = f.read()
m = re.search(rf'Id="(rId\d+)"[^>]*Target="slides/{target_file}"', rels)
if m:
    rid = m.group(1)
    with open('unpacked/ppt/presentation.xml') as f: pres = f.read()
    pres = re.sub(rf'\s*<p:sldId[^>]*r:id="{rid}"[^/]*/>', '', pres)
    with open('unpacked/ppt/presentation.xml', 'w') as f: f.write(pres)
    print(f'✅ 已移除 {target_file}')

# 删除后必须跑clean.py
subprocess.run(['python3', '/mnt/skills/public/pptx/scripts/clean.py', 'unpacked/'])
```

## 规则表

| 场景 | 需要clean.py | 需要全量质检 |
|------|------------|------------|
| 只改内容 | ❌ | ❌ 只验证改动页 |
| 换版式 | ✅ 必须 | ❌ 只验证改动页 |
| 插入页面 | ❌ | ❌ 只验证新页 |
| 删除页面 | ✅ 必须 | ❌ 验证相邻页 |
| 大范围≥5页 | ✅ | ✅ 全量 |

## 快速验证（不需要跑全量质检）

```python
# 只检查改动页的英文残留
import re
for fname in ['slide73.xml', 'slide74.xml']:
    with open(f'unpacked/ppt/slides/{fname}') as f: c = f.read()
    texts = re.findall(r'<a:t>([^<]{8,})</a:t>', c)
    bad = [t for t in texts if re.search(r'[a-zA-Z]{5,}',t) and not re.search(r'[一-龥]',t)]
    print(f'{fname}: {"✅" if not bad else f"❌ {bad[0][:40]}"}')
```

```bash
# 只生成改动页的高清图（如第5-6页）
python3 /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 -f 5 -l 6 output.pdf /home/claude/check
```
