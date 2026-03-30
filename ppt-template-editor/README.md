# ppt-template-editor

> **Priority: Highest** — This skill must be used for ALL PPT tasks when a `.pptx` template is involved.

A PowerPoint skill that generates or modifies slides by reusing an existing `.pptx` template's layout and style, replacing only text content.

---

## What It Does

**Input:** Content document (Word `.docx` / Markdown) + PPTX template  
**Output:** High-quality multi-page PPT (5–35 pages, full generation or incremental edit)

| Approach | How It Works | Result |
|---|---|---|
| **ppt-template-editor** ✅ | Copies template slides, preserving the full master chain: `slide → slideLayout → slideMaster → theme` | Background, colors, fonts, and graphic elements are automatically inherited |
| **ppt-generator** ❌ | Creates PPT from scratch, no master reference | Style is always lost — never use when a template exists |

---

## When to Use This Skill

Use this skill immediately when **any** of the following conditions are true:

1. ✅ User uploads a `.pptx` template file
2. ✅ User says "continue previous PPT task" or "resume session" — even if no new `.pptx` is uploaded this turn, the session file will reference one
3. ✅ User says "make PPT with this template", "generate with template style", or "fill content into template"
4. ✅ A `.pptx` template path exists in the task context (e.g. `/mnt/user-data/uploads/*.pptx`)

---

## The 9-Step Workflow

| Step | Action | Reference |
|------|--------|-----------|
| **0. Confirm** | Declare skill → check files → stop if missing | Step 0 above |
| **1. Confirm requirements** | Scope / language / page count / layout rules / narrative arc | "Step 1" below |
| **2. Create session file** | Write task state to `outputs/session_xxx.md` | "Session Format" below |
| **3. Clean contamination** | Detect and remove think-cell pollution | `references/pitfalls.md` |
| **4. Unpack template** | `unpack.py` → thumbnails → sp survey | `references/layout-rules.md` |
| **5. Plan layouts** | Select source slides, sp delta ≥ 20, three tiers complete | `references/layout-rules.md` |
| **6. Write content** | ppt_builder library: survey → RS → verify per page | `references/core-tools.md` |
| **7. Pack** | ① hardcode schemeClr → ② inject xfrm → ③ clean.py → ④ pack.py | `references/pitfalls.md` Lessons 1 & 2 |
| **8. QA** | 8.1 Structure · 8.2 Contamination · 8.3 Residuals · 8.4 Colors · 8.4b White-on-white · 8.5 Diversity · 8.6 Visual | `references/qa-scripts.md` |
| **9. Output** | Copy to `outputs/` + update session | See below |

---

## ⚠️ Seven Iron Rules (Failure if Violated)

1. **Use `rewrite_sp()` only** — string matching replacement silently fails on multi-run text
2. **Shapes with background fill use `make_para_colored(color='000000')`** — never use `FFFFFF` (white-on-white fails QA)
3. **Copy to `/mnt/user-data/outputs/` immediately after QA passes** — `/home/claude/` is cleared when session ends
4. **Read the session file first at the start of a new session** — `outputs/session_任务名.md`
5. **`sldIdLst` must be fully replaced, keeping only new pages** — the output file contains only the N pages generated this turn; all original template pages must be removed from sldIdLst. See Step 5.
6. **Hardcode schemeClr before packing** — otherwise PowerPoint colors drift to purple. See `references/pitfalls.md` Lesson 1.
7. **Inject title/body placeholder xfrm into `<p:spPr>` before packing** — otherwise title positions are misaligned in PowerPoint. See `references/pitfalls.md` Lesson 2.

---

## Step 1 — Confirm Requirements

```
1. PPT Scope:   Full / Chapter-only (no cover/TOC) / Single page increment
2. Language:    Template → Output: EN→ZH / EN→EN / ZH→ZH / ZH→EN
3. Page target: Lean(5-15) / Standard(15-25) / Full(25-40)
4. Layout rules (from template or user-specified):
   Default: font=Microsoft YaHei, title sz=2400, subtitle sz=1600,
            body sz=1200 + line_spacing=150000 (1.5x), big numbers sz=1800 (cap)
5. Narrative arc (2-3 sentences): Why → How → What first
```

Language rules:
- **ZH output**: in `save()`, replace `lang="zh-CN"`, font = Microsoft YaHei
- **EN output**: keep `lang="en-US"`, preserve template fonts
- **Residual detection**: ZH output checks for EN residuals; EN output checks for ZH residuals; EN→EN skips this check

---

## Step 2 — Session File Format

**New task:** create `/mnt/user-data/outputs/session_任务名.md`  
**Resume task:** read session file first, resume from the checkpoint

```
## Task Info
- Template: /mnt/user-data/uploads/xxx.pptx (cleaned version: template_clean.pptx)
- Content: /mnt/user-data/uploads/xxx.docx
- Output: /mnt/user-data/outputs/xxx.pptx
- Language: EN template → ZH output
- Layout: font=YaHei, body sz=1200 + ls=150000, title sz=2400

## Layout Plan
| Page | Source Slide | sp | Content | Status |
|------|--------------|----|---------|--------|
| 1    | slide42      | 22 | Four construction achievements | Done |
| 2    | slide11      | 31 | Policy responsiveness | In progress |

## Current Status
- Completed: 1 confirm · 2 session · 3 contamination clean · 4 unpack · 5 plan · 6 write (page 1)
- Next: Write page 2 (slide70)

## Output
- (fill in after QA passes)
```

Update the session file immediately after each key step.

---

## Core Commands

```bash
# Unpack template
python3 /mnt/skills/public/pptx/scripts/office/unpack.py template_clean.pptx unpacked/
# Thumbnails
python3 /mnt/skills/public/pptx/scripts/thumbnail.py template.pptx thumb --cols 5
# Copy source slide (one at a time; auto: file copy + rels + Content_Types; NOT sldIdLst)
python3 /mnt/skills/public/pptx/scripts/add_slide.py unpacked/ slideN.xml
# Clean
python3 /mnt/skills/public/pptx/scripts/clean.py unpacked/
# Pack
python3 /mnt/skills/public/pptx/scripts/office/pack.py unpacked/ output.pptx --original template_clean.pptx
# QA screenshots (HD)
python3 /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf /home/claude/slide
pdftoppm -jpeg -r 150 -f N -l M output.pdf /home/claude/slide  # pages N to M only
```

### After Step 5 — sldIdLst Full Replacement

```python
import re, subprocess

sources = [42, 11, 45, 41, 17, 37]   # ← fill in per layout plan
ORIGINAL_SLIDE_COUNT = 68             # ← must match actual count

for src in sources:
    r = subprocess.run(
        ['python3', '/mnt/skills/public/pptx/scripts/add_slide.py', 'unpacked/', f'slide{src}.xml'],
        capture_output=True, text=True)
    print(r.stdout.strip())

with open('unpacked/ppt/_rels/presentation.xml.rels') as f: rels = f.read()
new_slides = sorted(
    [(int(re.search(r'slide(\d+)', m.group(0)).group(1)), m.group(1))
     for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="slides/(slide(\d+))\.xml"', rels)
     if int(re.search(r'slide(\d+)', m.group(0)).group(1)) > ORIGINAL_SLIDE_COUNT],
    key=lambda x: x[0])

entries = [f'<p:sldId id="{700+i}" r:id="{rid}"/>' for i,(snum,rid) in enumerate(new_slides)]
new_lst = '<p:sldIdLst>\n        ' + '\n        '.join(entries) + '\n    </p:sldIdLst>'
with open('unpacked/ppt/presentation.xml') as f: pres = f.read()
pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', new_lst, pres, flags=re.DOTALL)
with open('unpacked/ppt/presentation.xml', 'w') as f: f.write(pres)
print(f'✅ sldIdLst replaced, output {len(new_slides)} pages (original removed)')
```

### Step 7 Sequence (Must Follow This Order)

Read `references/pitfalls.md` Lessons 1 & 2 for full code.

```
① hardcode_scheme_colors() — all new slides + slideMaster + slideLayout
② inject_title_xfrm()      — sp[0] (body13) and sp[1] (title) of all new slides
③ python3 clean.py unpacked/
④ python3 pack.py unpacked/ output.pptx --original template_clean.pptx
```

### Step 9 — Output (Mandatory)

```python
import shutil
shutil.copy('/home/claude/output.pptx', '/mnt/user-data/outputs/最终文件名.pptx')
# Then call present_files, then update session file
```

---

## Reference Documents

| Document | Content | When to Read |
|----------|---------|-------------|
| `references/core-tools.md` | ppt_builder.py full code + writing rules | Once before writing content |
| `references/layout-rules.md` | sp survey + layout selection + coordinate judgment + golden standard | When planning layouts |
| `references/qa-scripts.md` | Full QA scripts (8.1–8.6 six items) | During QA step |
| `references/pitfalls.md` | Contamination cleanup + common issues + hard lessons | When encountering errors or anomalies |
| `references/incremental.md` | Four incremental operations + quick validation | For single-page modifications |

---

## Quick Decision Flowchart

```
User requests PPT task?
├─ YES → Is there a .pptx template involved?
│         ├─ YES → Use ppt-template-editor ✅
│         └─ NO  → Check if user said "continue" with session file referencing .pptx
│                   └─ YES → Use ppt-template-editor ✅
├─ NO  → Is this a brand new PPT from scratch (no template)?
          └─ YES → Use ppt-generator (separate skill)
```

---

## Tech Background

**Why ppt-template-editor preserves styles:**

A PowerPoint template has a hierarchical structure:

```
slide (page)
  └─ slideLayout (page layout/master page)
       └─ slideMaster (master slide)
            └─ theme (colors, fonts, effects)
```

When you copy a slide from the template, the entire reference chain is preserved. PowerPoint resolves styles by looking up the chain at render time — so background, colors, fonts, and graphic elements all come through automatically.

If you create slides from scratch, there is no master chain. PowerPoint has no reference to resolve, so styles are lost.

**Conclusion:** Always use the template editor when a template exists. Only use ppt-generator for template-free tasks.