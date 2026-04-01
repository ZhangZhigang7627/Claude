# qa-scripts.md — 全量质检脚本（8.1-8.6）

所有质检项必须通过才能输出文件。

## 合并质检脚本（一次运行8.1-8.5）

```python
import zipfile, re
from pptx import Presentation

pptx = '/home/claude/output.pptx'
OUTPUT_LANG = 'zh'  # 'zh'=中文输出检测英文残留，'en'=英文输出检测中文残留，'en-en'=跳过残留检测

prs = Presentation(pptx)
print(f'8.1 结构: ✅ {len(prs.slides)} 张')

bad_kw = ['THINKCELL','think-cell','TCLayout','oleObj','AlternateContent','custDataLst']
pollution = []; issues = []; color_risks = []; sp_seq = []

with zipfile.ZipFile(pptx) as z:
    names = set(z.namelist())
    for name in names:
        if not (name.endswith('.xml') or name.endswith('.rels')): continue
        c = z.read(name).decode('utf-8','replace')
        if any(k.lower() in c.lower() for k in bad_kw): pollution.append(name)

    pres = z.read('ppt/presentation.xml').decode()
    sld_list = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', pres, re.DOTALL).group(1)
    rids = re.findall(r'r:id="(rId\d+)"', sld_list)
    rels = z.read('ppt/_rels/presentation.xml.rels').decode()

    for rid in rids:
        m = re.search(rf'Id="{rid}"[^>]*Target="slides/(slide\d+\.xml)"', rels)
        if not m: continue
        fname = m.group(1)
        c = z.read(f'ppt/slides/{fname}').decode('utf-8','replace')

        # 8.3 残留检测
        if OUTPUT_LANG != 'en-en':
            texts = re.findall(r'<a:t>([^<]{8,})</a:t>', c)
            if OUTPUT_LANG == 'zh':
                bad = [t for t in texts if re.search(r'[a-zA-Z]{5,}',t) and not re.search(r'[一-龥]',t)]
            else:
                bad = [t for t in texts if re.search(r'[一-龥]{3,}',t) and not re.search(r'[a-zA-Z]',t)]
            if bad: issues.append(f'  {fname}: {bad[0][:50]}')

        # 8.4 颜色继承风险
        spans = [(ms.start(),ms.end()) for ms in re.finditer(r'<p:sp>.*?</p:sp>',c,re.DOTALL)]
        for i,(s,e) in enumerate(spans):
            sp = c[s:e]
            if not re.findall(r'<a:t[^>]*>[^<]{1,}</a:t>', sp): continue
            spPr = re.search(r'<p:spPr>(.*?)</p:spPr>', sp, re.DOTALL)
            has_bg = bool(spPr and re.search(r'<a:(solidFill|gradFill)', spPr.group(1)))
            light_ref = bool(re.search(r'<a:fontRef[^>]*>.*?<a:schemeClr val="(lt1|lt2|bg1|bg2)"', sp, re.DOTALL))
            no_color = not bool(re.search(r'<a:rPr[^>]*>.*?<a:solidFill', sp, re.DOTALL))
            if (has_bg or light_ref) and no_color:
                t = ''.join(re.findall(r'<a:t[^>]*>([^<]*)</a:t>', sp)).strip()
                color_risks.append(f'  {fname} sp[{i}]: "{t[:35]}"')

        # 8.5 版式多样性
        n = len(re.findall(r'<p:sp>', c))
        tier = '○' if n<=6 else ('·' if n<=22 else '★')
        sp_seq.append((fname, n, tier))

print(f'8.2 污染: {"✅ 零" if not pollution else f"❌ {pollution}"}')
print(f'8.3 残留: {"✅ 无" if not issues else f"❌ {len(issues)}处"}')
for e in issues: print(e)
print(f'8.4 颜色风险: {"✅ 无" if not color_risks else f"❌ {len(color_risks)}处 → 改用make_para_colored()"}')
for r in color_risks: print(r)
sp_vals = [n for _,n,_ in sp_seq]
var = max(sp_vals) - min(sp_vals)
print(f'8.5 版式多样: sp范围{min(sp_vals)}-{max(sp_vals)}, 差值={var} {"✅" if var>=20 else "❌"}')
for fname,n,t in sp_seq: print(f'  {fname}: {n}sp {t}')

all_ok = not pollution and not issues and not color_risks and var >= 20
print(f'\n{"✅ 8.1-8.5全部通过" if all_ok else "❌ 有问题，修复后重新质检"}')
```

## 8.6 视觉检查（高清单页图）

```bash
# 全量
python3 /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf /home/claude/output.pptx
pdftoppm -jpeg -r 150 /home/claude/output.pdf /home/claude/slide
# 生成 slide-01.jpg, slide-02.jpg ...

# 只检查某几页（如第3-4页）
pdftoppm -jpeg -r 150 -f 3 -l 4 /home/claude/output.pdf /home/claude/slide
```

用 `view` 工具逐页检查：
- [ ] 文字无溢出（高清图下清晰可见）
- [ ] 无英文残留
- [ ] 版式有变化（不连续3页相同）
- [ ] 标题粗体大、正文细体小
- [ ] 每页内容充实，无大面积空白
- [ ] 图标与文字对齐（段落数=图标数）
- [ ] 有背景色的框里文字可见（不是白字白底）

## 质检汇总报告

```
═══════ 质检报告 ═══════
8.1 结构验证   ✅/❌
8.2 污染检查   ✅/❌
8.3 残留文字   ✅/❌
8.4 颜色风险   ✅/❌
8.5 版式多样   ✅/❌
8.6 视觉检查   ✅/❌（人工确认）
总体结论: ✅全部通过 / ❌有问题
```
