---
name: ppt-template-editor
metadata:
  version: "1.5"
  updated: "2026-03-31"
  changelog: "1.5:新增get_ph()解决标题顺序陷阱+del_gf()解决graphicFrame盲区+圆圈颜色修复;1.4:make_ph_para+save内置force_color;1.3:大页数分批规则;1.2:schemeClr/xfrm铁律;1.1:Step0/session/效率规则;1.0:初始"
description: >
  【优先级最高，覆盖所有PPT相关任务】以下任何情况都必须使用本Skill，严禁使用ppt-generator：
  1. 用户上传了.pptx格式的模板文件；
  2. 用户说"继续上次的PPT任务"、"继续任务"、"读取session文件继续"——即使本次对话没有上传.pptx，
     只要session文件或历史记录中提到了.pptx模板，就必须用本Skill；
  3. 用户说"用这个模板做PPT"、"按模板样式生成"、"在模板上填内容"；
  4. 任务上下文中存在.pptx模板路径（如/mnt/user-data/uploads/*.pptx）。
  【技术原因】ppt-template-editor通过复制模板slide保留完整母版引用链（slide→slideLayout→slideMaster→theme），
  生成的PPT自动继承模板的背景/配色/字体/图形元素。有模板就绝不能用ppt-generator。
  核心原则：复用.pptx模板的版式和样式，只替换文字内容，质检通过后必须输出到/mnt/user-data/outputs/。
---

# ppt-template-editor · v1.5 · 2026-03-31

输入：内容文档（Word/MD）+ PPTX模板 → 输出：高质量多页PPT  
支持全量生成（5-35页）和增量修改（单页/几页）。

---

## ⚠️ 第0步（每次必做）

**0.1 声明Skill**：回复用户"我将使用 `ppt-template-editor v1.5`"，等待确认。

**0.2 检查文件**：
```python
import glob
uploads = '/mnt/user-data/uploads'
pptx = glob.glob(f'{uploads}/*.pptx')
docx = glob.glob(f'{uploads}/*.docx')
print(f"✅ 模板:{pptx}  内容:{docx}" if pptx and docx else "❌ 文件缺失")
```

---

## ⚠️ 九条铁律（违反即失败）

1. **只用 `rewrite_sp()` 写内容**，禁止字符串匹配替换

2. **【v1.5核心】始终用 `get_ph()` 定位placeholder，禁止硬编码 `sp[0]`/`sp[1]`**：
   ```python
   ph = get_ph(c)
   c = rs(c, ph['title'], mph('主标题', bold=True))   # ✅ 永远正确
   c = rs(c, ph['body'],  mp('副标题', sz=1300))       # ✅ 永远正确
   # ❌ 禁止：rs(c, 0, mph(...)) / rs(c, 1, mph(...))   ← 该模板64/68张slide会写反！
   ```

3. **【v1.5核心】graphicFrame两种处理策略（表格可写入中文，无需删除）**：
   - **策略A（保留表格，写入中文内容）**：`scan_table()` 侦查结构 → `rtc()` 写入每个单元格，保留原有图标/样式/边框
   - **策略B（删除表格，改用普通sp布局）**：`del_gf()` 删除，再用 `rs()` 写普通内容
   ```python
   # 策略A：保留表格写中文（保留图标、边框等美观样式）
   scan_table(slide_num)             # 先侦查：几行几列
   c = load(slide_num)
   ph = get_ph(c)
   c = rs(c, ph['title'], mph('标题'))
   c = rtc(c, 0, 1, mtp('左列标题', bold=True))  # row=标题行, col=左文字列
   c = rtc(c, 1, 1, mtp('内容文字', sz=1100))
   c = rtc(c, 1, 3, mtp('右列内容', sz=1100))
   save(slide_num, c)

   # 策略B：删除表格（版式完全不适用时）
   c = load(slide_num)
   c = del_gf(c)    # 对无graphicFrame的slide也安全
   ph = get_ph(c)
   c = rs(c, ph['title'], mph('标题'))
   save(slide_num, c)
   ```
   该模板5张slide含graphicFrame：slide08/12/13/22/64，完整函数和用法见 `references/core-tools.md`。

4. **【v1.5】圆圈编号颜色规则**：
   - 白底描边圆（bg1白色填充+accent3蓝边框）→ `mpc(color='62B5E5')`
   - 有色填充圆（accent色有色填充）→ `mpc(color='FFFFFF')`
   - 不确定时：检查spPr的solidFill vs ln颜色

5. **【v1.4】文字颜色：`save()` 已内置自动保护**：
   - `make_para()`写的普通文字：`save()`自动注入333333，无需手动处理
   - `make_para_colored()`写的显式颜色：保留不变，确保颜色正确

6. **质检通过后立即 `cp` 到 `/mnt/user-data/outputs/`**

7. **新会话开始先读 session 文件**

8. **sldIdLst 必须完整替换，只保留新页**

9. **缩略图一次性全量查看**

---

## 标准流程（9步）

| 步骤 | 操作 | 参考文档 |
|------|------|---------|
| **0. 确认+检查** | 声明Skill → 等确认 → 检查文件 | 见上方 |
| **1. 确认需求** | PPT范围/语言/页数/叙事主线 | 见下方 |
| **2. Session文件** | 创建/读取断点状态 | 见下方 |
| **3. 污染清理** | 检测并清理think-cell污染 | `references/pitfalls.md` |
| **4. 解包+侦查** | unpack → 缩略图(全量看) → sp普查 → **graphicFrame检测** → 版式规划 | `references/layout-rules.md` |
| **5. 克隆+注册** | add_slide.py × N → sldIdLst完整替换 | 见核心命令速查 |
| **6. 写入内容** | **del_gf() → get_ph() → rs()** → verify | `references/core-tools.md` |
| **7. 清理打包** | clean.py → pack.py | `references/pitfalls.md` |
| **8. 质检** | 8.1结构 8.1b母版链 8.2污染 8.3残留 8.4颜色 8.5多样 8.6视觉（逐页） | `references/qa-scripts.md` |
| **9. 输出** | cp到outputs + present_files + 更新session | |

> **理想总轮数：2轮**（轮1=准备+写入，轮2=质检+输出）。

## ⚠️ 大页数任务分批规则

| 页数范围 | 策略 |
|---------|------|
| ≤20页 | 单次会话完整完成 |
| 21-30页 | **提前告知用户需要2次会话** |
| >30页 | **必须分批，每批≤15页** |

---

## 第一步：确认需求

```
1. PPT范围：完整独立 / 章节片段 / 单页增量
2. 模板语言 → 输出语言
3. 页数目标
4. 排版规范：字体YaHei, 正文sz=1200+ls=150000，placeholder字号由母版决定
5. 叙事主线（2-3句话）
```

---

## 第二步：Session文件

```
## 任务信息
- 模板/内容/输出路径；语言；排版（placeholder字号继承母版）
## 版式规划
| 页 | 源slide | 是否含graphicFrame | 内容 | 状态 |
## 当前状态 / 输出文件
```

---

## 第四步侦查：graphicFrame检测（新增必做项）

侦查阶段必须扫描所有候选slide是否含graphicFrame：

```python
import re, glob
for sf in glob.glob('unpacked/ppt/slides/slide*.xml'):
    with open(sf) as f: c = f.read()
    gfs = re.findall(r'<p:graphicFrame>.*?</p:graphicFrame>', c, re.DOTALL)
    if gfs:
        texts = [t for t in re.findall(r'<a:t[^>]*>([^<]+)</a:t>', ''.join(gfs)) if t.strip()]
        snum = re.search(r'slide(\d+)', sf).group(1)
        print(f"⚠️ slide{snum} 含{len(gfs)}个graphicFrame: {texts[:4]}")
```

**该模板含graphicFrame的slide（优先避免使用）**：slide08, slide12, slide13, slide22, slide64

---

## 核心命令速查

```bash
python3 /mnt/skills/public/pptx/scripts/office/unpack.py template.pptx unpacked/
python3 /mnt/skills/public/pptx/scripts/thumbnail.py template.pptx thumb --cols 5
python3 /mnt/skills/public/pptx/scripts/add_slide.py unpacked/ slideN.xml
python3 /mnt/skills/public/pptx/scripts/clean.py unpacked/
python3 /mnt/skills/public/pptx/scripts/office/pack.py unpacked/ output.pptx --original template.pptx
```

### sldIdLst完整替换（第5步后必执行）

```python
import re, subprocess
sources = [42, 11, 45]   # 按规划填写
ORIGINAL_SLIDE_COUNT = 68  # 原模板slide总数

for src in sources:
    subprocess.run(['python3','/mnt/skills/public/pptx/scripts/add_slide.py','unpacked/',f'slide{src}.xml'])

with open('unpacked/ppt/_rels/presentation.xml.rels') as f: rels=f.read()
new_slides=sorted([(int(re.search(r'slide(\d+)',m.group(0)).group(1)),m.group(1))
    for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="slides/(slide(\d+))\.xml"',rels)
    if int(re.search(r'slide(\d+)',m.group(0)).group(1))>ORIGINAL_SLIDE_COUNT],key=lambda x:x[0])
entries=[f'<p:sldId id="{700+i}" r:id="{rid}"/>' for i,(snum,rid) in enumerate(new_slides)]
new_lst='<p:sldIdLst>\n        '+'\n        '.join(entries)+'\n    </p:sldIdLst>'
with open('unpacked/ppt/presentation.xml') as f: pres=f.read()
pres=re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>',new_lst,pres,flags=re.DOTALL)
with open('unpacked/ppt/presentation.xml','w') as f: f.write(pres)
print(f'✅ sldIdLst已替换，输出{len(new_slides)}张')
```

```python
import shutil
shutil.copy('/home/claude/output.pptx', '/mnt/user-data/outputs/最终文件名.pptx')
```

---

## 参考文档索引

| 文档 | 内容 | 何时读 |
|------|------|--------|
| `references/core-tools.md` | **v1.5完整代码**（get_ph/del_gf/三大陷阱速查）| 写内容前必读 |
| `references/layout-rules.md` | sp普查 + 版式选择 + 坐标判断 | 规划版式时 |
| `references/qa-scripts.md` | 全量质检脚本（8.1-8.6六项） | 质检时 |
| `references/pitfalls.md` | 污染清理 + 血泪教训一~五 | 遇到报错或异常时 |
| `references/incremental.md` | 增量修改四种操作 | 单页修改时 |
