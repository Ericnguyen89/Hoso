# -*- coding: utf-8 -*-
"""
Web app (Flask) tạo bìa hồ sơ lưu trữ hàng loạt từ Excel.

Chạy trực tiếp, KHÔNG cần venv:
    python3 app.py
(Lần đầu thiếu thư viện, app sẽ tự cài rồi chạy lại.)
Sau đó mở:  http://127.0.0.1:5019
"""

import io
import os
import sys
import uuid
import datetime
import importlib.util
import subprocess


def ensure_deps():
    """Tự cài các thư viện còn thiếu bằng đúng Python đang chạy, rồi nạp lại.
    Nhờ vậy app chạy được mà không cần tạo/kích hoạt môi trường ảo (venv)."""
    need = []
    for module, pip_name in (("flask", "flask"),
                             ("docx", "python-docx"),
                             ("openpyxl", "openpyxl"),
                             ("lxml", "lxml")):
        if importlib.util.find_spec(module) is None:
            need.append(pip_name)
    if not need:
        return

    print("• Thiếu thư viện: %s — đang tự cài đặt..." % ", ".join(need))
    base = [sys.executable, "-m", "pip", "install", "--quiet"]
    # Trong venv: cài thẳng. Ngoài venv: ưu tiên user site, rồi phá khóa PEP668.
    in_venv = sys.prefix != getattr(sys, "base_prefix", sys.prefix)
    attempts = ([[]] if in_venv else [["--user"]]) + [["--break-system-packages"], []]

    last = None
    for extra in attempts:
        r = subprocess.run(base + extra + need,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        if r.returncode == 0:
            break
        last = r.stdout
    else:
        if last:
            sys.stdout.buffer.write(last)
        sys.exit("Không tự cài được. Hãy chạy thủ công:\n"
                 "  %s -m pip install %s" % (sys.executable, " ".join(need)))

    # Nạp lại tiến trình để dùng thư viện vừa cài
    os.execv(sys.executable, [sys.executable] + sys.argv)


ensure_deps()

from flask import (Flask, request, jsonify, send_file,  # noqa: E402
                   render_template_string, abort)

import tao_bia  # noqa: E402

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

HERE = os.path.dirname(os.path.abspath(__file__))
# Bộ nhớ tạm cho file kết quả: token -> {bytes, name, time}
STORE = {}


def cleanup():
    """Xóa kết quả cũ hơn 1 giờ."""
    now = datetime.datetime.now()
    for k in list(STORE):
        if (now - STORE[k]["time"]).total_seconds() > 3600:
            STORE.pop(k, None)


def pick_default(pattern, exclude=()):
    return tao_bia.autofind(HERE, pattern, exclude)


@app.route("/")
def index():
    return render_template_string(PAGE)


# Thư mục chứa tệp mẫu .docx (chỉ dev cấu hình, người dùng không thấy)
TEMPLATE_DIR = os.environ.get("BIA_TEMPLATE_DIR", HERE)


def _template_bytes():
    """Lấy bytes của tệp mẫu .docx cố định phía server."""
    import glob
    cands = [p for p in glob.glob(os.path.join(TEMPLATE_DIR, "*.docx"))
             if not os.path.basename(p).startswith("Bia_ho_so")]
    if not cands:
        raise ValueError("Server chưa có tệp mẫu .docx (liên hệ quản trị).")
    with open(cands[0], "rb") as f:
        return f.read()


def _load_data(req):
    """Lấy bytes Excel từ request (bắt buộc người dùng tải lên)."""
    data_file = req.files.get("data")
    if not (data_file and data_file.filename):
        raise ValueError("Vui lòng chọn tệp Excel (.xlsx).")
    if not data_file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("Tệp dữ liệu phải có đuôi .xlsx")
    return data_file.read(), data_file.filename


@app.route("/analyze", methods=["POST"])
def analyze():
    """Đọc dữ liệu, trả về số bìa + vài dòng xem trước (không tạo file)."""
    try:
        data_bytes, data_name = _load_data(request)
        import openpyxl
        ws = openpyxl.load_workbook(io.BytesIO(data_bytes),
                                    data_only=True).worksheets[0]
        rows = tao_bia.read_rows_ws(ws)
        if not rows:
            raise ValueError("Không đọc được dòng dữ liệu nào từ Excel.")
        preview = [{
            "ho_so_so": r["ho_so_so"],
            "title": r["title"],
            "start": r["start"],
            "end": r["end"],
            "so_trang": r["so_trang"],
            "so_tl": r["so_tl"],
            "thbq": r["thbq"],
            "org": r["org"],
        } for r in rows[:8]]
        return jsonify(ok=True, count=len(rows), preview=preview,
                       data_name=data_name)
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 400


@app.route("/generate", methods=["POST"])
def generate():
    """Tạo file docx, lưu tạm, trả token để tải về."""
    cleanup()
    try:
        data_bytes, data_name = _load_data(request)
        out_bytes, rows = tao_bia.generate_from_bytes(_template_bytes(), data_bytes)
        token = uuid.uuid4().hex
        base = os.path.splitext(os.path.basename(data_name))[0]
        STORE[token] = {
            "bytes": out_bytes,
            "name": "Bia_ho_so_%s.docx" % base,
            "time": datetime.datetime.now(),
        }
        return jsonify(ok=True, count=len(rows), token=token,
                       filename=STORE[token]["name"],
                       size=len(out_bytes))
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 400


@app.route("/download/<token>")
def download(token):
    item = STORE.get(token)
    if not item:
        abort(404)
    return send_file(
        io.BytesIO(item["bytes"]),
        mimetype="application/vnd.openxmlformats-officedocument."
                 "wordprocessingml.document",
        as_attachment=True,
        download_name=item["name"],
    )


# --------------------------- Giao diện (HTML) ---------------------------------
PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tạo bìa hồ sơ lưu trữ</title>
<style>
  :root{
    --bg1:#0f172a; --bg2:#1e293b;
    --card:#ffffff; --ink:#0f172a; --muted:#64748b; --line:#e2e8f0;
    --brand:#4f46e5; --brand2:#7c3aed; --ok:#16a34a; --bad:#dc2626;
    --shadow:0 10px 40px rgba(2,6,23,.12);
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",
       Roboto,Helvetica,Arial,"Apple Color Emoji",sans-serif;color:var(--ink);
       background:radial-gradient(1200px 600px at 10% -10%,#312e81 0%,transparent 60%),
                  radial-gradient(1000px 500px at 100% 0%,#6d28d9 0%,transparent 55%),
                  linear-gradient(180deg,var(--bg1),var(--bg2));
       min-height:100vh;padding:32px 16px;}
  .wrap{max-width:980px;margin:0 auto;}
  header{color:#fff;text-align:center;margin-bottom:26px;}
  header .badge{display:inline-flex;align-items:center;gap:8px;
       background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);
       padding:6px 14px;border-radius:999px;font-size:13px;backdrop-filter:blur(6px);}
  header h1{font-size:30px;margin:14px 0 6px;font-weight:800;letter-spacing:-.02em;}
  header p{margin:0;color:#c7d2fe;font-size:15px;}
  .card{background:var(--card);border-radius:20px;box-shadow:var(--shadow);
        padding:26px;}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
  @media (max-width:720px){.grid{grid-template-columns:1fr}}
  .drop{border:2px dashed #cbd5e1;border-radius:16px;padding:22px;text-align:center;
        cursor:pointer;transition:.18s;background:#f8fafc;position:relative;}
  .drop:hover{border-color:var(--brand);background:#f5f3ff;}
  .drop.drag{border-color:var(--brand);background:#eef2ff;transform:translateY(-2px);}
  .drop .ico{width:46px;height:46px;border-radius:12px;display:grid;place-items:center;
        margin:0 auto 10px;font-size:22px;color:#fff;
        background:linear-gradient(135deg,var(--brand),var(--brand2));}
  .drop .t{font-weight:700;font-size:15px}
  .drop .s{color:var(--muted);font-size:13px;margin-top:4px}
  .drop .fname{margin-top:10px;font-size:13px;font-weight:600;color:var(--brand);
        word-break:break-all;}
  .drop.has{border-style:solid;border-color:#c7d2fe;background:#fff;}
  input[type=file]{display:none}
  .actions{display:flex;gap:12px;margin-top:20px;flex-wrap:wrap}
  button{font:inherit;border:0;border-radius:12px;padding:13px 22px;font-weight:700;
         cursor:pointer;transition:.15s;display:inline-flex;align-items:center;gap:8px;}
  button:disabled{opacity:.55;cursor:not-allowed}
  .btn-primary{background:linear-gradient(135deg,var(--brand),var(--brand2));color:#fff;
         box-shadow:0 8px 20px rgba(79,70,229,.35);}
  .btn-primary:hover:not(:disabled){transform:translateY(-1px)}
  .btn-ghost{background:#f1f5f9;color:#334155}
  .btn-ghost:hover{background:#e2e8f0}
  .hint{color:var(--muted);font-size:13px;margin-top:14px;line-height:1.6}
  .panel{margin-top:22px;display:none}
  .stat{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:16px}
  .stat .box{flex:1;min-width:140px;background:#f8fafc;border:1px solid var(--line);
        border-radius:14px;padding:14px 16px;}
  .stat .n{font-size:26px;font-weight:800;color:var(--brand)}
  .stat .l{font-size:13px;color:var(--muted);margin-top:2px}
  table{width:100%;border-collapse:collapse;font-size:13px;margin-top:6px;
        overflow:hidden;border-radius:12px;border:1px solid var(--line)}
  th,td{padding:9px 11px;text-align:left;border-bottom:1px solid var(--line);
        vertical-align:top}
  th{background:#f8fafc;font-weight:700;color:#334155;font-size:12px;
     text-transform:uppercase;letter-spacing:.03em}
  tr:last-child td{border-bottom:0}
  td.title{max-width:320px}
  .tablewrap{overflow-x:auto;border-radius:12px}
  .toast{position:fixed;left:50%;bottom:26px;transform:translateX(-50%) translateY(20px);
        background:#0f172a;color:#fff;padding:12px 18px;border-radius:12px;
        box-shadow:var(--shadow);opacity:0;transition:.25s;pointer-events:none;z-index:9;
        font-size:14px;max-width:90vw}
  .toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
  .toast.bad{background:#7f1d1d}
  .spin{width:16px;height:16px;border:2px solid rgba(255,255,255,.4);
        border-top-color:#fff;border-radius:50%;animation:sp .7s linear infinite;display:none}
  @keyframes sp{to{transform:rotate(360deg)}}
  .dlcard{margin-top:18px;background:linear-gradient(135deg,#ecfdf5,#f0fdfa);
        border:1px solid #bbf7d0;border-radius:16px;padding:18px;display:none;
        align-items:center;gap:16px;flex-wrap:wrap}
  .dlcard .ic{width:44px;height:44px;border-radius:12px;background:var(--ok);color:#fff;
        display:grid;place-items:center;font-size:22px}
  .dlcard .meta{flex:1;min-width:160px}
  .dlcard .meta .a{font-weight:700}
  .dlcard .meta .b{color:var(--muted);font-size:13px}
  footer{color:#a5b4fc;text-align:center;font-size:13px;margin-top:24px}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <span class="badge">📚 Lưu trữ • Tự động hóa</span>
    <h1>Tạo bìa hồ sơ hàng loạt</h1>
    <p>Tải lên tệp mẫu (.docx) và dữ liệu (.xlsx) — hệ thống xuất bìa A4 cho mỗi hồ sơ.</p>
  </header>

  <div class="card">
    <label class="drop" id="dropData">
      <div class="ico">📊</div>
      <div class="t">Chọn tệp Excel (.xlsx)</div>
      <div class="s">Kéo thả vào đây hoặc bấm để chọn từ máy</div>
      <div class="fname" id="nameData"></div>
      <input type="file" id="fileData" accept=".xlsx,.xlsm">
    </label>

    <div class="actions">
      <button class="btn-primary" id="btnGen">
        <span class="spin" id="spinGen"></span>✨ Tạo bìa &amp; tải về
      </button>
      <button class="btn-ghost" id="btnAnalyze">🔍 Xem trước</button>
    </div>

    <div class="hint">
      Cấu trúc Excel (sheet đầu, có thể có dòng tiêu đề):
      <b>A</b>=Hồ sơ số · <b>B</b>=Nội dung/Tiêu đề · <b>C</b>=Thời gian (bắt đầu-kết thúc)
      · <b>D</b>=Số trang · <b>E</b>=Số TL · <b>F</b>=Thời hạn BQ · <b>G</b>=Dòng phông.
    </div>

    <div class="dlcard" id="dlcard">
      <div class="ic">✓</div>
      <div class="meta">
        <div class="a" id="dlName"></div>
        <div class="b" id="dlInfo"></div>
      </div>
      <a id="dlLink"><button class="btn-primary">⬇️ Tải file Word</button></a>
    </div>

    <div class="panel" id="panel">
      <div class="stat">
        <div class="box"><div class="n" id="stCount">0</div><div class="l">Tổng số bìa</div></div>
        <div class="box"><div class="n" id="stData" style="font-size:15px"></div><div class="l">Tệp dữ liệu</div></div>
      </div>
      <div style="font-weight:700;margin:6px 0 8px">Xem trước (tối đa 8 hồ sơ đầu)</div>
      <div class="tablewrap">
        <table>
          <thead><tr>
            <th>Hồ sơ số</th><th>Tiêu đề</th><th>Bắt đầu</th><th>Kết thúc</th>
            <th>Trang</th><th>Số TL</th><th>THBQ</th>
          </tr></thead>
          <tbody id="tbody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <footer>Bìa A4 • giữ nguyên font, khung &amp; bố cục của tệp mẫu</footer>
</div>

<div class="toast" id="toast"></div>

<script>
const $=s=>document.querySelector(s);
const fileData=$("#fileData");

function bind(drop, input, nameEl){
  drop.addEventListener("click",()=>input.click());
  input.addEventListener("change",()=>{
    if(input.files[0]){ nameEl.textContent=input.files[0].name; drop.classList.add("has"); }
  });
  ["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{
    e.preventDefault();drop.classList.add("drag");}));
  ["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{
    e.preventDefault();drop.classList.remove("drag");}));
  drop.addEventListener("drop",e=>{
    const f=e.dataTransfer.files[0]; if(!f)return;
    input.files=e.dataTransfer.files; nameEl.textContent=f.name; drop.classList.add("has");
  });
}
bind($("#dropData"),fileData,$("#nameData"));

let toastTimer;
function toast(msg,bad){
  const t=$("#toast"); t.textContent=msg; t.className="toast show"+(bad?" bad":"");
  clearTimeout(toastTimer); toastTimer=setTimeout(()=>t.className="toast",3200);
}
function esc(s){return (s??"").toString().replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));}

function payload(){
  const fd=new FormData();
  if(fileData.files[0]) fd.append("data",fileData.files[0]);
  return fd;
}
function hasFile(){
  if(!fileData.files[0]){ toast("Vui lòng chọn tệp Excel (.xlsx)",true); return false; }
  return true;
}
function renderPreview(d){
  $("#stCount").textContent=d.count;
  if(d.data_name) $("#stData").textContent=d.data_name;
  const tb=$("#tbody"); tb.innerHTML="";
  (d.preview||[]).forEach(r=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`<td><b>${esc(r.ho_so_so)}</b></td>
      <td class="title">${esc(r.title)}</td>
      <td>${esc(r.start)}</td><td>${esc(r.end)}</td>
      <td>${esc(r.so_trang)}</td><td>${esc(r.so_tl)}</td><td>${esc(r.thbq)}</td>`;
    tb.appendChild(tr);
  });
  $("#panel").style.display="block";
}

$("#btnAnalyze").addEventListener("click",async()=>{
  if(!hasFile()) return;
  try{
    const res=await fetch("/analyze",{method:"POST",body:payload()});
    const d=await res.json();
    if(!d.ok) return toast(d.error,true);
    renderPreview(d); $("#dlcard").style.display="none";
    toast("Đã phân tích: "+d.count+" hồ sơ");
  }catch(e){toast("Lỗi: "+e,true);}
});

$("#btnGen").addEventListener("click",async()=>{
  if(!hasFile()) return;
  const btn=$("#btnGen"), sp=$("#spinGen");
  btn.disabled=true; sp.style.display="inline-block";
  try{
    const res=await fetch("/generate",{method:"POST",body:payload()});
    const d=await res.json();
    if(!d.ok) return toast(d.error,true);
    // xem trước kèm theo (gọi analyze để có bảng)
    try{
      const a=await(await fetch("/analyze",{method:"POST",body:payload()})).json();
      if(a.ok) renderPreview(a);
    }catch(_){}
    $("#dlName").textContent=d.filename;
    $("#dlInfo").textContent=d.count+" bìa • "+(d.size/1024).toFixed(0)+" KB";
    $("#dlLink").href="/download/"+d.token;
    $("#dlcard").style.display="flex";
    toast("Đã tạo "+d.count+" bìa ✔");
  }catch(e){toast("Lỗi: "+e,true);}
  finally{btn.disabled=false; sp.style.display="none";}
});
</script>
</body>
</html>
"""

if __name__ == "__main__":
    import socket
    ip = socket.gethostbyname(socket.gethostname())
    print("Truy cập:  http://127.0.0.1:5019   |   http://%s:5019" % ip)
    app.run(host="0.0.0.0", port=5019, debug=False)
