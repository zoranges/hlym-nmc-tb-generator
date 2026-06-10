"""HLYM NMC TB Auto Generator — Flask API Server.

三合一后端: YMC (Section A) + YMAC (Section B) + LCP (Section C).

启动:
    python server.py
    然后打开 http://localhost:8080
"""
from __future__ import annotations

import io
import json
import os
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Any

from flask import Flask, Response, jsonify, render_template, request, send_file

# ---------------------------------------------------------------------------
# Flask 初始化
# ---------------------------------------------------------------------------

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

PROJECT_ROOT = Path(__file__).resolve().parent.parent
_STRUCT_PATH = os.environ.get("STRUCT_PATH", "")
if _STRUCT_PATH:
    STRUCT_PATH = Path(_STRUCT_PATH)
else:
    STRUCT_PATH = PROJECT_ROOT / "NMC Project" / "[POC] TB Structure Explanation.xlsx"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# 加载各个 app 的核心逻辑 (跳过 Streamlit UI 代码)
# ---------------------------------------------------------------------------

def _load_core(script_name: str) -> dict[str, Any]:
    """通过 exec 加载 app 的核心函数, 避开 Streamlit UI 代码."""
    import types

    script_path = Path(__file__).resolve().parent / script_name
    with open(script_path, "r", encoding="utf-8") as f:
        code = f.read()
    # 找 UI 分割标记
    for marker in ["# Part 5: Streamlit UI", "# Part 6: Streamlit UI"]:
        idx = code.find(marker)
        if idx >= 0:
            code = code[:idx]
            break

    mod_name = f"core_{script_name.replace('.py', '')}"
    mod = types.ModuleType(mod_name)
    mod.__file__ = str(script_path)
    mod.__builtins__ = __builtins__
    sys.modules[mod_name] = mod

    # 占位 streamlit 模块, 防止 import streamlit 失败
    fake_st = types.ModuleType("streamlit")
    sys.modules.setdefault("streamlit", fake_st)

    exec(code, mod.__dict__)
    return mod.__dict__


ymc_ns = _load_core("ymc_app.py")
ymac_ns = _load_core("ymac_app.py")
lcp_ns = _load_core("lcp_app.py")

print("[init] Core logic loaded:")
print(f"  YMC:  {sorted(k for k in ymc_ns if k.startswith('generate_section') or k == 'load_pcl')}")
print(f"  YMAC: {sorted(k for k in ymac_ns if k.startswith('generate_section') or k == 'load_cit')}")
print(f"  LCP:  {sorted(k for k in lcp_ns if k.startswith('generate_section') or k == 'load_cit')}")


# ---------------------------------------------------------------------------
# 通用: 保存上传文件
# ---------------------------------------------------------------------------

def _save_upload(uploaded_file) -> Path:
    tmp_dir = Path(tempfile.mkdtemp())
    dest = tmp_dir / uploaded_file.filename
    uploaded_file.save(str(dest))
    return dest


# ---------------------------------------------------------------------------
# 路由
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    """主页面."""
    return render_template("index.html")


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})


# ---------------------------------------------------------------------------
# YMC (Section A)
# ---------------------------------------------------------------------------

@app.route("/api/generate/ymc", methods=["POST"])
def generate_ymc():
    """生成 Section A."""
    try:
        # 文件
        if "file" in request.files and request.files["file"].filename:
            struct_arg = _save_upload(request.files["file"])
        elif STRUCT_PATH.exists():
            struct_arg = STRUCT_PATH
        else:
            return jsonify({"error": "No file provided and built-in sample not found"}), 400

        tb_no = request.form.get("tb_no", "TB-25-137")
        target_factory = request.form.get("target_factory", "F3")

        # 加载数据
        pcl = ymc_ns["load_pcl"](struct_arg)
        cu_fa_map = ymc_ns["load_cu_fa_code"](struct_arg)
        ln_bom_set = ymc_ns["load_ln_bom"](struct_arg)
        qcps_sub = ymc_ns["load_qcps_sub"](struct_arg)
        qcps_mlr = ymc_ns["load_qcps_ml"](struct_arg, "4. QCPS - ML (R)")
        qcps_mll = ymc_ns["load_qcps_ml"](struct_arg, "5. QCPS - ML (L)")
        ypl = ymc_ns["load_ypl"](struct_arg)
        email_map = ymc_ns["load_email"](struct_arg)

        model = str(pcl.get("MODEL", ["?"]).iloc[0]) if not pcl.empty else "?"
        model_name = str(pcl.get("MODEL_NAME", ["?"]).iloc[0]) if not pcl.empty and "MODEL_NAME" in pcl.columns else ""

        rows, logs = ymc_ns["generate_section_a"](
            pcl, cu_fa_map, ln_bom_set, qcps_sub, qcps_mlr, qcps_mll, ypl, email_map,
        )

        # 写 Excel
        xlsx_bytes = ymc_ns["write_section_a_excel"](rows, tb_no, target_factory, model, model_name)

        # 保存供下载
        out_name = f"{tb_no}-{target_factory}-SectionA.xlsx"
        out_path = OUTPUT_DIR / out_name
        out_path.write_bytes(xlsx_bytes)

        # 构建行数据
        rows_json = [
            {
                "no": i, "sec": r.sec, "lvl1": r.lvl1, "lvl2": r.lvl2, "lvl3": r.lvl3, "lvl4": r.lvl4,
                "part_no": r.part_no, "part_name": r.part_name, "qty": r.qty,
                "supplier": r.supplier, "use_in": r.use_in, "ex_new": r.ex_new, "remarks": r.remarks,
            }
            for i, r in enumerate(rows, 1)
        ]

        ex_count = sum(1 for r in rows if r.ex_new == "EX")
        new_count = sum(1 for r in rows if r.ex_new == "NEW")

        return jsonify({
            "success": True,
            "rows": rows_json,
            "logs": logs,
            "stats": {"total": len(rows), "ex": ex_count, "new": new_count, "model": model, "model_name": model_name},
            "download_url": f"/api/download/{out_name}",
        })
    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


# ---------------------------------------------------------------------------
# YMAC (Section B)
# ---------------------------------------------------------------------------

@app.route("/api/generate/ymac", methods=["POST"])
def generate_ymac():
    """生成 Section B."""
    try:
        if "file" in request.files and request.files["file"].filename:
            struct_arg = _save_upload(request.files["file"])
        elif STRUCT_PATH.exists():
            struct_arg = STRUCT_PATH
        else:
            return jsonify({"error": "No file provided and built-in sample not found"}), 400

        tb_no = request.form.get("tb_no", "TB-25-137")
        target_factory = request.form.get("target_factory", "F3")

        pcl = ymac_ns["load_pcl"](struct_arg)
        cu_fa_map = ymac_ns["load_cu_fa_code"](struct_arg)
        ln_bom_set = ymac_ns["load_ln_bom"](struct_arg)
        qcps_sub = ymac_ns["load_qcps_sub"](struct_arg)
        qcps_mlr = ymac_ns["load_qcps_ml"](struct_arg, "4. QCPS - ML (R)")
        qcps_mll = ymac_ns["load_qcps_ml"](struct_arg, "5. QCPS - ML (L)")
        ypl = ymac_ns["load_ypl"](struct_arg)
        cit_a, cit_b = ymac_ns["load_cit"](struct_arg)
        email_map = ymac_ns["load_email"](struct_arg)

        model = str(pcl.get("MODEL", ["?"]).iloc[0]) if not pcl.empty else "?"
        model_name = str(pcl.get("MODEL_NAME", ["?"]).iloc[0]) if not pcl.empty and "MODEL_NAME" in pcl.columns else ""

        rows, logs = ymac_ns["generate_section_b"](
            pcl, cu_fa_map, ln_bom_set, qcps_sub, qcps_mlr, qcps_mll,
            ypl, cit_a, cit_b, email_map,
        )

        xlsx_bytes = ymac_ns["write_section_b_excel"](rows, tb_no, target_factory, model, model_name)

        out_name = f"{tb_no}-{target_factory}-SectionB.xlsx"
        out_path = OUTPUT_DIR / out_name
        out_path.write_bytes(xlsx_bytes)

        rows_json = [
            {
                "no": i, "sec": r.sec, "lvl1": r.lvl1, "lvl2": r.lvl2, "lvl3": r.lvl3, "lvl4": r.lvl4,
                "part_no": r.part_no, "part_name": r.part_name, "qty": r.qty,
                "supplier": r.supplier, "use_in": r.use_in, "ex_new": r.ex_new, "remarks": r.remarks,
            }
            for i, r in enumerate(rows, 1)
        ]

        ex_count = sum(1 for r in rows if r.ex_new == "EX")
        new_count = sum(1 for r in rows if r.ex_new == "NEW")
        f3_count = sum(1 for r in rows if r.use_in == "F3")
        vendor_count = sum(1 for r in rows if r.use_in != "F3")

        return jsonify({
            "success": True,
            "rows": rows_json,
            "logs": logs,
            "stats": {
                "total": len(rows), "ex": ex_count, "new": new_count,
                "f3": f3_count, "vendor": vendor_count,
                "model": model, "model_name": model_name,
            },
            "download_url": f"/api/download/{out_name}",
        })
    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


# ---------------------------------------------------------------------------
# LCP (Section C)
# ---------------------------------------------------------------------------

@app.route("/api/generate/lcp", methods=["POST"])
def generate_lcp():
    """生成 Section C (VMI)."""
    try:
        if "file" in request.files and request.files["file"].filename:
            struct_arg = _save_upload(request.files["file"])
        elif STRUCT_PATH.exists():
            struct_arg = STRUCT_PATH
        else:
            return jsonify({"error": "No file provided and built-in sample not found"}), 400

        tb_no = request.form.get("tb_no", "TB-25-137")
        target_factory = request.form.get("target_factory", "F3")

        pcl = lcp_ns["load_pcl"](struct_arg)
        cu_fa_map = lcp_ns["load_cu_fa_code"](struct_arg)
        ln_bom_set = lcp_ns["load_ln_bom"](struct_arg)
        qcps_sub = lcp_ns["load_qcps_sub"](struct_arg)
        cit_matrix = lcp_ns["load_cit_matrix"](struct_arg)

        model = str(pcl.get("MODEL", ["?"].iloc[0]) if not pcl.empty else "?"
        model_name = str(pcl.get("MODEL_NAME", ["?"].iloc[0]) if not pcl.empty and "MODEL_NAME" in pcl.columns else ""

        groups, logs, total_rows = lcp_ns["generate_section_c"](
            qcps_sub, pcl, cit_matrix, cu_fa_map, ln_bom_set,
        )

        xlsx_bytes = lcp_ns["write_section_c_excel"](groups, tb_no, target_factory, model, model_name)

        out_name = f"{tb_no}-{target_factory}-SectionC.xlsx"
        out_path = OUTPUT_DIR / out_name
        out_path.write_bytes(xlsx_bytes)

        # 扁平化所有行
        all_rows = []
        for g in groups:
            for r in g.rows:
                all_rows.append(r)

        rows_json = [
            {
                "no": i, "sec": r.sec, "lvl1": r.lvl1, "lvl2": r.lvl2, "lvl3": r.lvl3, "lvl4": r.lvl4,
                "part_no": r.part_no, "part_name": r.part_name, "qty": r.qty,
                "supplier": r.supplier, "use_in": r.use_in, "ex_new": r.ex_new, "remarks": r.remarks,
            }
            for i, r in enumerate(all_rows, 1)
        ]

        return jsonify({
            "success": True,
            "rows": rows_json,
            "logs": logs,
            "stats": {
                "total": total_rows, "groups": len(groups),
                "color_groups": sum(1 for g in groups if g.color_mark),
                "model": model, "model_name": model_name,
            },
            "download_url": f"/api/download/{out_name}",
        })
    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


# ---------------------------------------------------------------------------
# 下载
# ---------------------------------------------------------------------------

@app.route("/api/download/<filename>")
def download_file(filename: str):
    """下载生成的 Excel 文件."""
    # 安全检查
    if ".." in filename or "/" in filename or "\\" in filename:
        return "Invalid filename", 400
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        return "File not found", 404
    return send_file(
        str(file_path),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


# ---------------------------------------------------------------------------
# 启动
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print(f"[server] Struct file: {STRUCT_PATH} (exists={STRUCT_PATH.exists()})")
    print(f"[server] Output dir: {OUTPUT_DIR}")
    print(f"[server] Starting on http://localhost:8080")
    app.run(host="0.0.0.0", port=8080, debug=True)
