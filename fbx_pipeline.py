import json
import subprocess
import tempfile
from pathlib import Path


BLENDER_SCRIPT = Path(__file__).with_name("blender_fbx_optimize.py")


def _run_blender_job(blender_path: Path, payload: dict):
    if not blender_path.exists():
        raise ValueError(f"Blender 路径不存在: {blender_path}")
    if not BLENDER_SCRIPT.exists():
        raise ValueError("缺少 blender_fbx_optimize.py")

    with tempfile.TemporaryDirectory(prefix="mesh_opt_blender_") as temp_dir_raw:
        temp_dir = Path(temp_dir_raw)
        config_path = temp_dir / "config.json"
        report_path = temp_dir / "report.json"
        payload = dict(payload)
        payload["report_path"] = str(report_path)
        config_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        cmd = [
            str(blender_path),
            "--background",
            "--factory-startup",
            "--python",
            str(BLENDER_SCRIPT),
            "--",
            str(config_path),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="ignore")
        if result.returncode != 0:
            detail = (result.stderr or "").strip() or (result.stdout or "").strip()
            raise RuntimeError(f"Blender 处理失败: {detail}")
        if not report_path.exists():
            detail = (result.stdout or "").strip() or "未生成报告"
            raise RuntimeError(f"Blender 处理失败: {detail}")
        return json.loads(report_path.read_text(encoding="utf-8"))


def optimize_model(source_path: Path, output_path: Path, blender_path: Path, params: dict):
    payload = {
        "mode": "optimize",
        "source_path": str(source_path),
        "output_path": str(output_path),
        "params": params,
    }
    return _run_blender_job(blender_path, payload)


def inspect_model(source_path: Path, blender_path: Path):
    payload = {
        "mode": "inspect",
        "source_path": str(source_path),
    }
    return _run_blender_job(blender_path, payload)
