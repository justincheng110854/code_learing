#!/usr/bin/env python3
"""
Thesis-to-PPT Web Matching Tool — 拖拽论文，勾选匹配，导出PPTX

Usage:
    python thesis2ppt_web.py [--port 5000]
    Then open http://127.0.0.1:5000 in browser.
"""

import argparse
import json
import os
import re
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from thesis2ppt import (
    ThesisParser, PPTBuilder, map_sections_to_slides,
    extract_images_from_docx, find_image_references,
    ensure_image_png, summarize_section, generate_ppt,
)

from flask import Flask, request, jsonify, send_file, render_template_string

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Session storage
# ---------------------------------------------------------------------------
SESSION = {}


# ---------------------------------------------------------------------------
# HTML — single page with drag-drop, section list, image grid, export
# ---------------------------------------------------------------------------
HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>毕业论文答辩PPT — 拖拽匹配工具</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"Microsoft YaHei","SimHei",sans-serif;background:#f0f2f5;color:#333;display:flex;flex-direction:column;height:100vh}
/* ── Header ── */
.header{background:#003366;color:#fff;padding:12px 20px;display:flex;align-items:center;justify-content:space-between}
.header h1{font-size:19px}
.header .hint{font-size:12px;opacity:.8}
/* ── Drop Zone ── */
.drop-zone{border:3px dashed #c0c0c0;border-radius:10px;margin:16px 20px 0;padding:24px;text-align:center;transition:all .2s;background:#fff;cursor:pointer}
.drop-zone.dragover{border-color:#0066CC;background:#e6f0ff}
.drop-zone.has-file{border-color:#28a745;background:#f0fff4;border-style:solid}
.drop-zone .dz-icon{font-size:36px;margin-bottom:6px}
.drop-zone .dz-text{font-size:15px;color:#555}
.drop-zone .dz-sub{font-size:12px;color:#999;margin-top:4px}
.drop-zone .dz-file{font-size:13px;color:#28a745;font-weight:bold;margin-top:4px;display:none}
.drop-zone.has-file .dz-file{display:block}
.drop-zone.has-file .dz-placeholder{display:none}
.drop-zone input[type=file]{display:none}
/* ── Meta row ── */
.meta-row{display:flex;gap:8px;padding:10px 20px;background:#fff;border-bottom:1px solid #eee;flex-wrap:wrap;align-items:center}
.meta-row input{padding:6px 10px;border:1px solid #ccc;border-radius:4px;font-size:12px}
.meta-row input.m-title{width:280px}
.meta-row input.m-author{width:80px}
.meta-row input.m-advisor{width:80px}
.meta-row input.m-uni{width:130px}
.meta-row input.m-date{width:80px}
.meta-row label{font-size:12px;font-weight:bold}
/* ── Main panels ── */
.main{display:flex;flex:1;overflow:hidden}
.panel{overflow-y:auto;padding:14px}
.panel-left{width:42%;border-right:1px solid #ddd;background:#fff}
.panel-right{width:58%;background:#fafafa}
.panel h2{font-size:14px;margin-bottom:10px;padding-bottom:6px;border-bottom:2px solid #0066CC;color:#003366}
/* ── Section items ── */
.section-item{padding:10px 12px;margin-bottom:6px;border:2px solid #e0e0e0;border-radius:6px;cursor:pointer;transition:all .12s;display:flex;align-items:center;gap:10px}
.section-item:hover{border-color:#0066CC;background:#f0f6ff}
.section-item.selected{border-color:#0066CC;background:#e6f0ff;box-shadow:0 0 0 2px rgba(0,102,204,.15)}
.section-num{font-size:11px;color:#0066CC;font-weight:bold;min-width:28px}
.section-info{flex:1;min-width:0}
.section-title{font-size:13px;font-weight:bold}
.section-preview{font-size:11px;color:#999;margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.section-match{font-size:11px;color:#28a745;font-weight:bold;margin-top:2px}
.section-match .none{color:#bbb}
.unmatch-btn{font-size:10px;color:#dc3545;cursor:pointer;margin-left:4px;text-decoration:underline}
/* ── Image grid ── */
.image-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px}
.image-card{border:2px solid #e0e0e0;border-radius:6px;overflow:hidden;cursor:pointer;transition:all .12s;background:#fff}
.image-card:hover{border-color:#0066CC;transform:translateY(-1px);box-shadow:0 3px 10px rgba(0,0,0,.08)}
.image-card.selected{border-color:#28a745;box-shadow:0 0 0 3px rgba(40,167,69,.2)}
.image-card.matched{border-color:#ffc107;opacity:.65}
.image-card img{width:100%;height:110px;object-fit:contain;background:#f5f5f5;display:block}
.image-card .cap{padding:5px 7px;font-size:10px;color:#666;max-height:28px;overflow:hidden;line-height:1.3}
/* ── Bottom bar ── */
.bottom-bar{background:#fff;border-top:2px solid #ddd;padding:10px 20px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.match-summary{display:flex;gap:10px;flex-wrap:wrap;font-size:11px}
.match-summary span{background:#e6f0ff;padding:3px 7px;border-radius:3px;white-space:nowrap}
.export-area{display:flex;gap:8px;align-items:center}
.export-area button{padding:8px 16px;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:bold}
.btn-export{background:#28a745;color:#fff}
.btn-export:hover{background:#1e7e34}
.btn-reset{background:#fff;color:#dc3545;border:1px solid #dc3545}
.export-status{font-size:11px;color:#888}
.empty{text-align:center;color:#bbb;padding:50px 20px;font-size:14px}
.empty .big{font-size:40px;margin-bottom:8px}
/* ── Toast ── */
.toast{position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#333;color:#fff;padding:10px 24px;border-radius:6px;font-size:13px;z-index:999;opacity:0;transition:opacity .3s}
.toast.show{opacity:1}
.spinner{display:inline-block;width:14px;height:14px;border:2px solid #fff;border-top-color:transparent;border-radius:50%;animation:spin .6s linear infinite;margin-right:6px;vertical-align:middle}
@keyframes spin{to{transform:rotate(360deg)}}
</style>
</head>
<body>

<div class="header">
  <h1>毕业论文答辩PPT — 拖拽匹配工具</h1>
  <span class="hint" id="stepHint">拖拽论文文件到下方区域开始</span>
</div>

<!-- Drag & Drop Zone -->
<div class="drop-zone" id="dropZone">
  <div class="dz-placeholder">
    <div class="dz-icon">📂</div>
    <div class="dz-text">拖拽论文文件到此处</div>
    <div class="dz-sub">支持 PDF / DOCX / TXT / Markdown</div>
  </div>
  <div class="dz-file" id="dzFileName"></div>
  <input type="file" id="fileInput" accept=".pdf,.docx,.txt,.md">
</div>

<!-- Metadata -->
<div class="meta-row">
  <label>自动提取:</label>
  <input type="text" class="m-title" id="mtitle" placeholder="论文标题">
  <input type="text" class="m-author" id="mauthor" placeholder="作者">
  <input type="text" class="m-advisor" id="madvisor" placeholder="导师">
  <input type="text" class="m-uni" id="muni" placeholder="学校/学院">
  <input type="text" class="m-date" id="mdate" placeholder="日期">
</div>

<!-- Main content -->
<div class="main">
  <div class="panel panel-left" id="sectionPanel">
    <h2>📋 章节列表（点击选中，再点右侧图片匹配）</h2>
    <div id="sectionList"><div class="empty"><div class="big">📄</div>拖拽论文文件后自动加载</div></div>
  </div>
  <div class="panel panel-right" id="imagePanel">
    <h2>🖼️ 图片库（点击图片匹配到当前选中章节）</h2>
    <div id="imageGrid" class="image-grid"><div class="empty"><div class="big">🖼️</div>加载论文后显示图片</div></div>
  </div>
</div>

<!-- Bottom bar -->
<div class="bottom-bar" id="bottomBar" style="display:none">
  <div class="match-summary" id="matchSummary">
    <span style="color:#999">尚未匹配</span>
  </div>
  <div class="export-area">
    <span class="export-status" id="exportStatus"></span>
    <button class="btn-reset" onclick="resetAll()">重置</button>
    <button class="btn-export" onclick="doExport()">⬇ 导出 PPTX</button>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
/* ── State ── */
let sections=[], images=[], matches={}, selSection=null, sessionId='';
// matches format: {section_title: [img_filename, ...]}

/* ── Drag & Drop ── */
const dz = document.getElementById('dropZone');
const fi = document.getElementById('fileInput');

dz.addEventListener('click', ()=>fi.click());
fi.addEventListener('change', ()=>handleFile(fi.files[0]));

dz.addEventListener('dragover', e=>{e.preventDefault();dz.classList.add('dragover')});
dz.addEventListener('dragleave', ()=>dz.classList.remove('dragover'));
dz.addEventListener('drop', e=>{
  e.preventDefault(); dz.classList.remove('dragover');
  if(e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});

async function handleFile(file){
  if(!file){return}
  document.getElementById('dzFileName').textContent = '已选择: ' + file.name;
  dz.classList.add('has-file');
  toast('正在解析论文...', true);

  let fd = new FormData();
  fd.append('file', file);
  fd.append('meta', JSON.stringify({
    title: g('mtitle').value, author: g('mauthor').value,
    advisor: g('madvisor').value, university: g('muni').value,
    date: g('mdate').value
  }));

  try {
    let r = await fetch('/api/load', {method:'POST', body:fd});
    let d = await r.json();
    if(d.error){toast('加载失败: '+d.error);return}

    sections = d.sections; images = d.images; sessionId = d.session_id || '';
    sections.forEach(s=>matches[s.title]=[]);
    if(d.meta){
      if(d.meta.title) g('mtitle').value=d.meta.title;
      if(d.meta.author) g('mauthor').value=d.meta.author;
      if(d.meta.university) g('muni').value=d.meta.university;
    }
    renderAll();
    g('bottomBar').style.display='flex';
    g('stepHint').textContent = `已加载 ${sections.length} 个章节, ${images.length} 张图片`;
    toast('加载完成!');
  }catch(e){toast('网络错误: '+e.message)}
}

/* ── Render ── */
function g(id){return document.getElementById(id)}

function renderAll(){renderSections();renderImages();updateSummary()}

function renderSections(){
  g('sectionList').innerHTML = sections.map((s,i)=>{
    let imgs = matches[s.title]||[];
    let label = imgs.length
      ? '✅ '+imgs.length+' 张图: '+imgs.map(f=>f.replace(/\.\w+$/,'')).join(', ')
      : '<span class="none">○ 未匹配</span>';
    return `<div class="section-item${selSection===s.title?' selected':''}" id="sec${i}" onclick="pickSection('${esc(s.title)}')">
      <div class="section-num">${esc(s.section_num)}</div>
      <div class="section-info">
        <div class="section-title">${esc(s.title)}</div>
        <div class="section-preview">${esc(s.content_preview||'')}</div>
        <div class="section-match">${label}
          ${imgs.length?`<span class="unmatch-btn" onclick="event.stopPropagation();unmatchSec('${esc(s.title)}')">清除</span>`:''}
        </div>
      </div>
    </div>`;
  }).join('');
}

function renderImages(){
  let h = '';
  if(!images.length){h='<div class="empty"><div class="big">🖼️</div>无可提取图片</div>'}
  else {
    // Build a set of which section each image belongs to
    let imgOwner = {};
    for(let [sec, arr] of Object.entries(matches)){
      (arr||[]).forEach(f=>{imgOwner[f]=sec});
    }
    images.forEach((img,i)=>{
      let owner = imgOwner[img.filename];
      let cls = '';
      if(selSection && owner===selSection) cls=' selected';
      else if(owner) cls=' matched';
      h+=`<div class="image-card${cls}" onclick="pickImage('${esc(img.filename)}')">
        <img src="${img.url}" alt="" loading="lazy">
        <div class="cap">${esc(img.caption||img.filename)}</div></div>`;
    });
  }
  g('imageGrid').innerHTML=h;
}

function pickSection(title){
  selSection = title;
  document.querySelectorAll('.section-item').forEach(e=>e.classList.remove('selected'));
  let idx = sections.findIndex(s=>s.title===title);
  if(idx>=0) document.getElementById('sec'+idx).classList.add('selected');
  renderImages();
}

function pickImage(fname){
  if(!selSection){toast('请先在左侧选中一个章节');return}
  let arr = matches[selSection]||[];
  let idx = arr.indexOf(fname);
  if(idx>=0){
    // Already matched to this section → remove
    arr.splice(idx,1);
  } else {
    // Remove from other sections if matched elsewhere
    for(let [sec, imgs] of Object.entries(matches)){
      let i2 = (imgs||[]).indexOf(fname);
      if(i2>=0 && sec!==selSection){
        if(!confirm('该图片已匹配到 "'+sec+'"，移到当前章节？'))return;
        imgs.splice(i2,1); break;
      }
    }
    arr.push(fname);
  }
  matches[selSection]=arr;
  renderAll();
}

function unmatchSec(title){matches[title]=[];renderAll()}

function resetAll(){
  if(!confirm('清除所有匹配？'))return;
  sections.forEach(s=>matches[s.title]=[]);
  selSection=null;renderAll();
}

function updateSummary(){
  let total=0, totalImgs=0;
  sections.forEach(s=>{let n=(matches[s.title]||[]).length;if(n)total++;totalImgs+=n;});
  g('matchSummary').innerHTML = total
    ? sections.filter(s=>(matches[s.title]||[]).length).map(s=>
        `<span>${esc(s.title)} → ${(matches[s.title]||[]).map(f=>f.replace(/\.\w+$/,'')).join(', ')}</span>`
      ).join('')+` <b>(${total}章节/${totalImgs}图)</b>`
    : '<span style="color:#999">点击左侧章节 + 右侧图片进行匹配（可多图）</span>';
}

async function doExport(){
  g('exportStatus').innerHTML='<span class="spinner"></span>生成中...';
  let meta = {
    title:g('mtitle').value, author:g('mauthor').value,
    advisor:g('madvisor').value, university:g('muni').value,
    date:g('mdate').value
  };
  let mapping = sections.map(s=>({section_title:s.title, images:matches[s.title]||[]}));
  try {
    let r = await fetch('/api/export',{method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({mapping,meta,session_id:sessionId})});
    if(!r.ok){g('exportStatus').textContent='导出失败';return}
    let blob = await r.blob();
    let a = document.createElement('a');
    a.href=URL.createObjectURL(blob); a.download='答辩PPT.pptx';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    g('exportStatus').textContent='导出完成!';
    toast('PPTX 已下载');
    setTimeout(()=>g('exportStatus').textContent='',3000);
  }catch(e){g('exportStatus').textContent='网络错误'}
}

function toast(msg,spin){
  let t = g('toast');
  t.innerHTML = (spin?'<span class="spinner"></span>':'') + msg;
  t.classList.add('show');
  clearTimeout(t._tid);
  if(!spin) t._tid = setTimeout(()=>t.classList.remove('show'),2000);
}
function esc(s){return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# API Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/api/load", methods=["POST"])
def api_load():
    """Accept uploaded thesis file, parse it, return sections + images."""
    if "file" not in request.files:
        return jsonify({"error": "未上传文件"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "文件名为空"}), 400

    meta_raw = request.form.get("meta", "{}")
    try:
        meta = json.loads(meta_raw)
    except json.JSONDecodeError:
        meta = {}

    # Save uploaded file to temp
    ext = os.path.splitext(file.filename)[1].lower()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    file.save(tmp.name)
    tmp.close()
    thesis_path = tmp.name

    try:
        # Parse thesis
        parser = ThesisParser()
        sections = parser.parse(thesis_path)
        slide_plan = map_sections_to_slides(sections)

        # Extract images if DOCX
        image_refs = []
        image_dir = None
        if ext == ".docx":
            image_dir = os.path.join(tempfile.gettempdir(), "thesis_ppt_imgs")
            try:
                extract_images_from_docx(thesis_path, image_dir)
                image_refs = find_image_references(thesis_path)
                # Convert all EMF to PNG immediately so browser can display them
                for ref in image_refs:
                    fname = os.path.basename(ref.get("filename", ""))
                    if fname.lower().endswith(".emf"):
                        ensure_image_png(image_dir, fname)
            except Exception:
                pass

        # Store in session
        sid = str(id(tmp))
        SESSION[sid] = {
            "thesis_path": thesis_path,
            "sections": sections,
            "slide_plan": slide_plan,
            "image_dir": image_dir,
            "image_refs": image_refs,
            "metadata": meta,
        }

        # Build response: sections
        sec_list = []
        for item in slide_plan:
            content = item.get("content", "")
            preview = content[:80].replace("\n", " ") if content else ""
            sec_list.append({
                "title": item["title"],
                "section_num": item["section_num"],
                "content_preview": preview,
            })

        # Build response: images
        img_list = []
        for ref in image_refs:
            fname = os.path.basename(ref.get("filename", ""))
            if not fname:
                continue
            img_list.append({
                "filename": fname,
                "caption": ref.get("caption", "")[:80],
                "url": f"/api/image/{sid}/{fname}",
            })

        # Auto-detect metadata
        auto_meta = {}
        for sec in sections:
            if sec.get("level") == 0:
                content = sec.get("content", "")
                lines = content.split("\n")
                for line in lines:
                    line = line.strip()
                    if len(line) > 5 and not any(
                        kw in line for kw in ["学院", "专业", "年级", "学号", "学生", "指导", "摘要", "关键词"]
                    ):
                        auto_meta["title"] = line
                        break
                m = re.search(r"学生姓名[：:]\s*(.*)", content)
                if m:
                    auto_meta["author"] = m.group(1).strip()
                m = re.search(r"学\s*院[：:]\s*(.*)", content)
                if m:
                    auto_meta["university"] = m.group(1).strip()
                break

        return jsonify({
            "session_id": sid,
            "sections": sec_list,
            "images": img_list,
            "meta": auto_meta,
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/image/<sid>/<path:filename>")
def api_image(sid, filename):
    """Serve extracted image for a session."""
    session = SESSION.get(sid, {})
    image_dir = session.get("image_dir")
    if not image_dir:
        return "No session", 404
    for fname in [filename, os.path.splitext(filename)[0] + ".png"]:
        fp = os.path.join(image_dir, os.path.basename(fname))
        if os.path.exists(fp):
            ext = os.path.splitext(fp)[1].lower()
            mt = {".png":"image/png",".jpg":"image/jpeg",".jpeg":"image/jpeg",
                  ".gif":"image/gif",".bmp":"image/bmp"}.get(ext,"image/png")
            return send_file(fp, mimetype=mt)
    return "Not found", 404


@app.route("/api/export", methods=["POST"])
def api_export():
    """Generate PPTX from mapping."""
    data = request.get_json()
    mapping = data.get("mapping", [])
    meta = data.get("meta", {})
    sid = data.get("session_id", "")

    session = SESSION.get(sid, {})
    thesis_path = session.get("thesis_path")
    image_dir = session.get("image_dir")

    if not thesis_path:
        return jsonify({"error": "请先加载论文"}), 400

    # Convert EMF images (handle both single 'image' and array 'images')
    for m in mapping:
        imgs = m.get("images") or ([m.get("image")] if m.get("image") else [])
        for fname in imgs:
            if fname and fname.lower().endswith(".emf") and image_dir:
                ensure_image_png(image_dir, fname)

    # Write mapping JSON (use 'image' for backward compat with thesis2ppt.py)
    compat_mapping = []
    for m in mapping:
        imgs = m.get("images") or ([m.get("image")] if m.get("image") else [])
        compat_mapping.append({
            "section_title": m["section_title"],
            "image": imgs[0] if imgs else None,
            "images": imgs,  # Extra images beyond the first
        })

    mapping_path = os.path.join(tempfile.gettempdir(), "img_map.json")
    with open(mapping_path, "w", encoding="utf-8") as f:
        json.dump({"mappings": compat_mapping}, f, ensure_ascii=False, indent=2)

    output_path = os.path.join(tempfile.gettempdir(), "defense_out.pptx")
    generate_ppt(
        filepath=thesis_path,
        output_path=output_path,
        title=meta.get("title", ""),
        author=meta.get("author", ""),
        advisor=meta.get("advisor", ""),
        university=meta.get("university", ""),
        date_str=meta.get("date", ""),
        image_mapping_json=mapping_path,
    )

    return send_file(
        output_path, as_attachment=True,
        download_name="答辩PPT.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


# ---------------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser(description="Thesis-to-PPT Web Tool")
    p.add_argument("--port", type=int, default=5000)
    p.add_argument("--host", default="127.0.0.1")
    args = p.parse_args()
    print(f"\n  毕业论文答辩PPT 拖拽匹配工具")
    print(f"  ==============================")
    print(f"  浏览器打开: http://{args.host}:{args.port}")
    print(f"  按 Ctrl+C 退出\n")
    app.run(host=args.host, port=args.port, debug=False)


if __name__ == "__main__":
    main()
