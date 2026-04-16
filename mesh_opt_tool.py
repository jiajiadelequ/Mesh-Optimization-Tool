import json
import queue
import subprocess
import threading
import traceback
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pymeshlab as ml


APP_TITLE = "OBJ Batch Optimizer"
DEFAULT_MESHLAB = r"E:\MeshLab\meshlab.exe"
SETTINGS_PATH = Path(__file__).with_name("mesh_opt_tool_settings.json")
NODE_CONVERTER = Path(__file__).with_name("obj_to_drc.js")


@dataclass
class ModelMeta:
    path: Path
    display_name: str
    group_name: str
    source_size: int


class ModelRow:
    def __init__(self, parent, app, meta: ModelMeta):
        self.app = app
        self.meta = meta
        self.var_selected = tk.BooleanVar(value=True)
        self.var_ratio = tk.StringVar(value="0.10")
        self.var_quality = tk.StringVar(value="0.5")
        self.var_preserve_boundary = tk.BooleanVar(value=True)
        self.var_preserve_normal = tk.BooleanVar(value=True)
        self.var_preserve_topology = tk.BooleanVar(value=True)
        self.var_result = tk.StringVar(value="结果: 未处理")

        self.frame = ttk.Frame(parent, padding=(8, 6))
        self.frame.columnconfigure(1, weight=1)
        self.frame.columnconfigure(2, weight=0)

        left = ttk.Frame(self.frame)
        left.grid(row=0, column=0, sticky="nw", padx=(0, 10))
        ttk.Checkbutton(left, variable=self.var_selected).grid(row=0, column=0, sticky="w")

        body = ttk.Frame(self.frame)
        body.grid(row=0, column=1, sticky="nsew")
        body.columnconfigure(0, weight=1)

        title = f"{meta.display_name} [{meta.path.name}]"
        ttk.Label(body, text=title).grid(row=0, column=0, sticky="w")

        subtitle = f"分组: {meta.group_name or '-'} | size: {meta.source_size / 1024 / 1024:.2f} MB"
        ttk.Label(body, text=subtitle).grid(row=1, column=0, sticky="w")

        options = ttk.Frame(body)
        options.grid(row=2, column=0, sticky="ew", pady=(4, 0))

        ttk.Label(options, text="比例").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.var_ratio, width=7).grid(row=0, column=1, sticky="w", padx=(0, 12))

        ttk.Label(options, text="Quality").grid(row=0, column=2, sticky="w")
        ttk.Entry(options, textvariable=self.var_quality, width=6).grid(row=0, column=3, sticky="w", padx=(0, 12))

        ttk.Checkbutton(options, text="Boundary", variable=self.var_preserve_boundary).grid(row=0, column=4, sticky="w")
        ttk.Checkbutton(options, text="Normal", variable=self.var_preserve_normal).grid(row=0, column=5, sticky="w")
        ttk.Checkbutton(options, text="Topology", variable=self.var_preserve_topology).grid(row=0, column=6, sticky="w")

        self.lbl_result = ttk.Label(body, textvariable=self.var_result)
        self.lbl_result.grid(row=3, column=0, sticky="w", pady=(6, 0))

        actions = ttk.Frame(self.frame)
        actions.grid(row=0, column=2, sticky="ne", padx=(10, 0))
        ttk.Button(actions, text="减面", command=self.simplify_one).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(actions, text="预览", command=self.preview_one).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(actions, text="重置到 1/10", command=self.reset_defaults).grid(row=0, column=2)
    def grid(self, **kwargs):
        self.frame.grid(**kwargs)

    def reset_defaults(self):
        self.var_ratio.set("0.10")
        self.var_quality.set("0.5")
        self.var_preserve_boundary.set(True)
        self.var_preserve_normal.set(True)
        self.var_preserve_topology.set(True)
        self.var_result.set("结果: 未处理")

    def current_params(self):
        ratio = float(self.var_ratio.get().strip())
        quality = float(self.var_quality.get().strip())
        return {
            "ratio": ratio,
            "qualitythr": quality,
            "preserveboundary": self.var_preserve_boundary.get(),
            "preservenormal": self.var_preserve_normal.get(),
            "preservetopology": self.var_preserve_topology.get(),
        }

    def simplify_one(self):
        self.app.start_single_job(self, preview=False)

    def preview_one(self):
        self.app.start_single_job(self, preview=True)


class MeshOptApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1280x820")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.rows = []
        self.job_queue: "queue.Queue[tuple[str, str]]" = queue.Queue()
        self.worker_running = False
        self.load_running = False
        self.loading_total = 0

        self.var_source_dir = tk.StringVar()
        self.var_output_dir = tk.StringVar()
        self.var_drc_input_dir = tk.StringVar()
        self.var_drc_output_dir = tk.StringVar()
        self.var_meshlab_path = tk.StringVar(value=DEFAULT_MESHLAB)
        self.var_select_all = tk.BooleanVar(value=True)
        self.var_generate_data_json = tk.BooleanVar(value=True)
        self.var_status = tk.StringVar(value="就绪")

        self._load_settings()
        self._build_ui()
        self.root.after(150, self._poll_queue)

        source_dir = Path(self.var_source_dir.get()) if self.var_source_dir.get() else Path.cwd().parent / "obj_out"
        if source_dir.exists():
            self.var_source_dir.set(str(source_dir))
            self.load_models_async()

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        top = ttk.Frame(self.root, padding=10)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        top.columnconfigure(4, weight=1)

        ttk.Label(top, text="原始 OBJ 目录").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.var_source_dir).grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Button(top, text="选择", command=self.choose_source_dir).grid(row=0, column=2, padx=(0, 12))

        ttk.Label(top, text="输出目录").grid(row=0, column=3, sticky="w")
        ttk.Entry(top, textvariable=self.var_output_dir).grid(row=0, column=4, sticky="ew", padx=(6, 6))
        ttk.Button(top, text="选择", command=self.choose_output_dir).grid(row=0, column=5)
        ttk.Button(top, text="打包并复制到 PICO", command=self.start_export_zip_to_pico).grid(row=0, column=6, padx=(8, 0))

        ttk.Label(top, text="MeshLab").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.var_meshlab_path).grid(row=1, column=1, columnspan=4, sticky="ew", padx=(6, 6), pady=(8, 0))
        ttk.Button(top, text="选择", command=self.choose_meshlab).grid(row=1, column=5, pady=(8, 0))

        ttk.Label(top, text="DRC 输入目录").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.var_drc_input_dir).grid(row=2, column=1, columnspan=4, sticky="ew", padx=(6, 6), pady=(8, 0))
        ttk.Button(top, text="选择", command=self.choose_drc_input_dir).grid(row=2, column=5, pady=(8, 0))

        ttk.Label(top, text="DRC 输出目录").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.var_drc_output_dir).grid(row=3, column=1, columnspan=4, sticky="ew", padx=(6, 6), pady=(8, 0))
        ttk.Button(top, text="选择", command=self.choose_drc_output_dir).grid(row=3, column=5, pady=(8, 0))
        ttk.Button(top, text="批量 OBJ->DRC", command=self.start_batch_drc_job).grid(row=3, column=6, padx=(8, 0))
        ttk.Button(top, text="打包并复制到 PICO", command=self.start_export_drc_zip_to_pico).grid(row=3, column=7, padx=(8, 0))

        controls = ttk.Frame(top)
        controls.grid(row=4, column=0, columnspan=6, sticky="w", pady=(10, 0))
        ttk.Checkbutton(controls, text="全选", variable=self.var_select_all, command=self.toggle_select_all).grid(row=0, column=0, padx=(0, 12))
        ttk.Checkbutton(controls, text="输出后生成 data.json", variable=self.var_generate_data_json).grid(row=0, column=1, padx=(0, 12))
        ttk.Button(controls, text="刷新模型列表", command=self.load_models_async).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(controls, text="批量减面已勾选", command=self.start_batch_job).grid(row=0, column=3, padx=(0, 8))
        ttk.Button(controls, text="保存设置", command=self.save_settings).grid(row=0, column=4)

        center = ttk.Frame(self.root, padding=(10, 0, 10, 0))
        center.grid(row=1, column=0, sticky="nsew")
        center.columnconfigure(0, weight=1)
        center.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(center, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(center, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        bottom = ttk.Frame(self.root, padding=10)
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.var_status).grid(row=0, column=0, sticky="w")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def choose_source_dir(self):
        path = filedialog.askdirectory(initialdir=self._resolve_initial_dir(self.var_source_dir.get()))
        if path:
            self.var_source_dir.set(path)
            self.save_settings()
            self.load_models_async()

    def choose_output_dir(self):
        path = filedialog.askdirectory(initialdir=self._resolve_initial_dir(self.var_output_dir.get()))
        if path:
            self.var_output_dir.set(path)
            self.save_settings()

    def choose_drc_input_dir(self):
        path = filedialog.askdirectory(initialdir=self._resolve_initial_dir(self.var_drc_input_dir.get()))
        if path:
            self.var_drc_input_dir.set(path)
            self.save_settings()

    def choose_drc_output_dir(self):
        path = filedialog.askdirectory(initialdir=self._resolve_initial_dir(self.var_drc_output_dir.get()))
        if path:
            self.var_drc_output_dir.set(path)
            self.save_settings()

    def choose_meshlab(self):
        path = filedialog.askopenfilename(initialdir=self._resolve_initial_dir(Path(self.var_meshlab_path.get()).parent if self.var_meshlab_path.get() else Path.cwd()), filetypes=[("Executable", "*.exe")])
        if path:
            self.var_meshlab_path.set(path)
            self.save_settings()

    def toggle_select_all(self):
        checked = self.var_select_all.get()
        for row in self.rows:
            row.var_selected.set(checked)

    def _find_data_json(self, source_dir: Path):
        candidates = [source_dir / "data.json", source_dir.parent / "data.json"]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return None

    def _resolve_initial_dir(self, raw_path):
        if not raw_path:
            return str(Path.cwd())
        path = Path(raw_path)
        if path.exists() and path.is_dir():
            return str(path)
        for parent in [path] + list(path.parents):
            if parent.exists() and parent.is_dir():
                return str(parent)
        return str(Path.cwd())

    def _load_name_map(self, source_dir: Path):
        mapping = {}
        data_json = self._find_data_json(source_dir)
        if not data_json:
            return mapping
        try:
            data = json.loads(data_json.read_text(encoding="utf-8"))
            config_by_id = {item["id"]: item for item in data.get("config_json", [])}
            for item in data.get("result_urls", []):
                config = config_by_id.get(item.get("id"), {})
                stem = item.get("filename")
                if stem:
                    mapping[stem] = {
                        "display_name": config.get("filename") or stem,
                        "group_name": config.get("group_name") or "",
                    }
        except Exception:
            pass
        return mapping

    def load_models_async(self):
        if self.load_running:
            return
        self.load_running = True
        self.var_status.set("正在后台加载模型列表...")
        threading.Thread(target=self._load_models_worker, daemon=True).start()

    def _load_models_worker(self):
        source_dir = Path(self.var_source_dir.get().strip()) if self.var_source_dir.get().strip() else Path.cwd().parent / "obj_out"
        if not source_dir.exists():
            self.job_queue.put(("loaded_models", {"rows": [], "status": "源目录不存在"}))
            return

        name_map = self._load_name_map(source_dir)
        obj_paths = sorted(
            (path for path in source_dir.glob("*.obj") if len(path.stem) == 36 and path.stem.count("-") == 4),
            key=lambda p: p.stat().st_size,
            reverse=True,
        )
        self.loading_total = len(obj_paths)
        loaded_rows = []
        for index, path in enumerate(obj_paths):
            try:
                meta_info = name_map.get(path.stem, {})
                meta = ModelMeta(path=path, display_name=meta_info.get("display_name", path.stem), group_name=meta_info.get("group_name", ""), source_size=path.stat().st_size)
                loaded_rows.append(meta)
                self.job_queue.put(("load_progress", f"正在后台加载模型列表... {index + 1}/{len(obj_paths)}"))
            except Exception as exc:
                loaded_rows.append(f"{path.name} 读取失败: {exc}")
                self.job_queue.put(("load_progress", f"正在后台加载模型列表... {index + 1}/{len(obj_paths)}"))

        self.job_queue.put(
            (
                "loaded_models",
                {
                    "rows": loaded_rows,
                    "status": f"已加载 {len(obj_paths)} 个模型。源目录只读，输出仅写入目标目录。",
                },
            )
        )

    def validate_common_inputs(self):
        source_dir = Path(self.var_source_dir.get().strip())
        output_dir = Path(self.var_output_dir.get().strip()) if self.var_output_dir.get().strip() else None
        meshlab_path = Path(self.var_meshlab_path.get().strip()) if self.var_meshlab_path.get().strip() else None

        if not source_dir.exists():
            raise ValueError("原始 OBJ 目录不存在")
        if output_dir is None:
            raise ValueError("请先设置输出目录")
        output_dir.mkdir(parents=True, exist_ok=True)
        return source_dir, output_dir, meshlab_path

    def validate_drc_inputs(self):
        drc_input_dir = Path(self.var_drc_input_dir.get().strip()) if self.var_drc_input_dir.get().strip() else None
        drc_output_dir = Path(self.var_drc_output_dir.get().strip()) if self.var_drc_output_dir.get().strip() else None
        if drc_input_dir is None or not drc_input_dir.exists():
            raise ValueError("请先设置有效的 DRC 输入目录")
        if drc_output_dir is None:
            raise ValueError("请先设置 DRC 输出目录")
        if not shutil.which("node"):
            raise ValueError("未找到 node，无法执行 OBJ->DRC 转换")
        drc_output_dir.mkdir(parents=True, exist_ok=True)
        if not NODE_CONVERTER.exists():
            raise ValueError("缺少 obj_to_drc.js 转换脚本")
        return drc_input_dir, drc_output_dir

    def _target_faces_from_row(self, row: ModelRow):
        params = row.current_params()
        return max(0.0, params["ratio"])

    def start_single_job(self, row: ModelRow, preview: bool):
        try:
            self.validate_common_inputs()
            row.current_params()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        action = "预览" if preview else "减面"
        self.var_status.set(f"正在{action}: {row.meta.path.name}")
        threading.Thread(target=self._run_single_job, args=(row, preview), daemon=True).start()

    def start_batch_job(self):
        selected_rows = [row for row in self.rows if row.var_selected.get()]
        if not selected_rows:
            messagebox.showwarning(APP_TITLE, "请先勾选至少一个模型")
            return
        try:
            self.validate_common_inputs()
            for row in selected_rows:
                row.current_params()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        if self.worker_running:
            messagebox.showwarning(APP_TITLE, "已有任务在运行")
            return
        self.worker_running = True
        threading.Thread(target=self._run_batch_job, args=(selected_rows,), daemon=True).start()

    def start_batch_drc_job(self):
        if self.worker_running:
            messagebox.showwarning(APP_TITLE, "已有任务在运行")
            return
        try:
            drc_input_dir, drc_output_dir = self.validate_drc_inputs()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.worker_running = True
        self.var_status.set(f"正在批量转换 OBJ -> DRC: {drc_input_dir}")
        threading.Thread(target=self._run_batch_drc_job, args=(drc_input_dir, drc_output_dir), daemon=True).start()

    def _run_single_job(self, row: ModelRow, preview: bool):
        try:
            _, output_dir, meshlab_path = self.validate_common_inputs()
            output_path = output_dir / row.meta.path.name
            if preview:
                output_path = output_dir / f"{row.meta.path.stem}_preview.obj"
            result = simplify_obj(row.meta.path, output_path, row.current_params(), self._target_faces_from_row(row))
            msg = (
                f"{row.meta.path.name}: target={result['target_faces']} | "
                f"{result['orig_faces']} -> {result['new_faces']} faces | "
                f"{result['output_size_mb']:.2f} MB"
            )
            if preview:
                if not meshlab_path or not meshlab_path.exists():
                    raise ValueError("MeshLab 路径不存在，无法预览")
                subprocess.Popen([str(meshlab_path), str(output_path)])
                msg += " | 已调用 MeshLab 预览"
            self.job_queue.put(("single_done", (row, msg, result)))
        except Exception as exc:
            self.job_queue.put(("error", f"{row.meta.path.name}: {exc}"))

    def _run_batch_job(self, rows):
        try:
            source_dir, output_dir, _ = self.validate_common_inputs()
            report = []
            for index, row in enumerate(rows, start=1):
                result = simplify_obj(row.meta.path, output_dir / row.meta.path.name, row.current_params(), self._target_faces_from_row(row))
                report.append(result)
                self.job_queue.put((
                    "batch_item_done",
                    (
                        row,
                        f"[{index}/{len(rows)}] {row.meta.path.name}: "
                        f"target={result['target_faces']} | {result['orig_faces']} -> {result['new_faces']} faces | "
                        f"{result['output_size_mb']:.2f} MB",
                        result,
                    ),
                ))
            if self.var_generate_data_json.get():
                write_output_data_json(source_dir, output_dir)
            report_path = output_dir / "tool_simplify_report.json"
            report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
            self.job_queue.put(("info", f"批量完成，输出目录: {output_dir}"))
        except Exception:
            self.job_queue.put(("error", traceback.format_exc()))
        finally:
            self.worker_running = False

    def start_export_zip_to_pico(self):
        if self.worker_running:
            messagebox.showwarning(APP_TITLE, "已有任务在运行")
            return
        try:
            _, output_dir, _ = self.validate_common_inputs()
            if not output_dir.exists():
                raise ValueError("输出目录不存在")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.worker_running = True
        self.var_status.set(f"正在打包并复制到 PICO: {output_dir}")
        threading.Thread(target=self._run_export_zip_to_pico, args=(output_dir,), daemon=True).start()

    def _run_export_zip_to_pico(self, output_dir: Path):
        try:
            zip_path = create_output_zip(output_dir)
            pico_status = try_copy_zip_to_pico_ords(zip_path)
            self.job_queue.put(("info", f"打包完成: {zip_path.name} | PICO: {pico_status}"))
        except Exception:
            self.job_queue.put(("error", traceback.format_exc()))
        finally:
            self.worker_running = False

    def start_export_drc_zip_to_pico(self):
        if self.worker_running:
            messagebox.showwarning(APP_TITLE, "已有任务在运行")
            return
        try:
            _, drc_output_dir = self.validate_drc_inputs()
            if not drc_output_dir.exists():
                raise ValueError("DRC 输出目录不存在")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.worker_running = True
        self.var_status.set(f"正在打包并复制 DRC 到 PICO: {drc_output_dir}")
        threading.Thread(target=self._run_export_zip_to_pico, args=(drc_output_dir,), daemon=True).start()

    def _run_batch_drc_job(self, obj_dir: Path, drc_output_dir: Path):
        try:
            obj_files = sorted(
                (path for path in obj_dir.glob("*.obj") if len(path.stem) == 36 and path.stem.count("-") == 4),
                key=lambda p: p.stat().st_size,
                reverse=True,
            )
            if not obj_files:
                raise ValueError("输出目录里没有可转换的 OBJ 文件")

            for index, obj_path in enumerate(obj_files, start=1):
                drc_path = drc_output_dir / f"{obj_path.stem}.drc"
                cmd = ["node", str(NODE_CONVERTER), str(obj_path), str(drc_path)]
                subprocess.run(cmd, check=True, capture_output=True, text=True)
                self.job_queue.put(("status", f"[{index}/{len(obj_files)}] OBJ->DRC: {obj_path.name}"))

            self._write_drc_data_json(obj_dir, drc_output_dir)
            self.job_queue.put(("info", f"OBJ->DRC 完成，输出目录: {drc_output_dir}"))
        except subprocess.CalledProcessError as exc:
            detail = exc.stderr.strip() or exc.stdout.strip() or str(exc)
            self.job_queue.put(("error", f"OBJ->DRC 失败: {detail}"))
        except Exception:
            self.job_queue.put(("error", traceback.format_exc()))
        finally:
            self.worker_running = False

    def _poll_queue(self):
        try:
            while True:
                level, text = self.job_queue.get_nowait()
                if level == "loaded_models":
                    self._apply_loaded_models(text)
                elif level == "load_progress":
                    self.var_status.set(text)
                elif level == "single_done":
                    row, msg, result = text
                    self._update_row_result(row, result)
                    self.var_status.set(f"完成: {msg}")
                    messagebox.showinfo(APP_TITLE, msg)
                elif level == "batch_item_done":
                    row, msg, result = text
                    self._update_row_result(row, result)
                    self.var_status.set(msg)
                else:
                    self.var_status.set(text)
                    if level == "error":
                        messagebox.showerror(APP_TITLE, text)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)

    def _apply_loaded_models(self, payload):
        for row in self.rows:
            row.frame.destroy()
        self.rows.clear()

        for child in self.inner.winfo_children():
            child.destroy()

        for index, item in enumerate(payload["rows"]):
            if isinstance(item, ModelMeta):
                row = ModelRow(self.inner, self, item)
                row.grid(row=index, column=0, sticky="ew", pady=(0, 8))
                self.rows.append(row)
            else:
                ttk.Label(self.inner, text=item).grid(row=index, column=0, sticky="w")

        self.load_running = False
        self.var_status.set(payload["status"])
        self.save_settings()

    def _update_row_result(self, row: ModelRow, result: dict):
        row.var_result.set(
            "结果: "
            f"target={result['target_faces']}, "
            f"actual={result['new_faces']} faces, "
            f"size={result['output_size_mb']:.2f} MB"
        )

    def save_settings(self):
        payload = {
            "source_dir": self.var_source_dir.get(),
            "output_dir": self.var_output_dir.get(),
            "drc_input_dir": self.var_drc_input_dir.get(),
            "drc_output_dir": self.var_drc_output_dir.get(),
            "meshlab_path": self.var_meshlab_path.get(),
            "generate_data_json": self.var_generate_data_json.get(),
        }
        SETTINGS_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

    def _load_settings(self):
        if not SETTINGS_PATH.exists():
            self.var_source_dir.set(str(Path.cwd().parent / "obj_out"))
            self.var_output_dir.set(str(Path.cwd().parent / "tool_output"))
            self.var_drc_input_dir.set(str(Path.cwd().parent / "tool_output"))
            self.var_drc_output_dir.set(str(Path.cwd().parent / "tool_output_drc"))
            return
        try:
            payload = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            self.var_source_dir.set(payload.get("source_dir", str(Path.cwd().parent / "obj_out")))
            self.var_output_dir.set(payload.get("output_dir", str(Path.cwd().parent / "tool_output")))
            self.var_drc_input_dir.set(payload.get("drc_input_dir", str(Path.cwd().parent / "tool_output")))
            self.var_drc_output_dir.set(payload.get("drc_output_dir", str(Path.cwd().parent / "tool_output_drc")))
            self.var_meshlab_path.set(payload.get("meshlab_path", DEFAULT_MESHLAB))
            self.var_generate_data_json.set(payload.get("generate_data_json", True))
        except Exception:
            self.var_source_dir.set(str(Path.cwd().parent / "obj_out"))
            self.var_output_dir.set(str(Path.cwd().parent / "tool_output"))
            self.var_drc_input_dir.set(str(Path.cwd().parent / "tool_output"))
            self.var_drc_output_dir.set(str(Path.cwd().parent / "tool_output_drc"))
            self.var_meshlab_path.set(DEFAULT_MESHLAB)

    def on_close(self):
        self.save_settings()
        self.root.destroy()

    def _write_drc_data_json(self, obj_dir: Path, drc_output_dir: Path):
        source_json = next((path for path in [obj_dir / "data.json", obj_dir.parent / "data.json"] if path.exists()), None)
        if not source_json:
            return
        data = json.loads(source_json.read_text(encoding="utf-8"))
        data["order_id"] = drc_output_dir.name
        if isinstance(data.get("patient"), dict):
            data["patient"]["name"] = drc_output_dir.name
        result_urls = []
        for item in data.get("result_urls", []):
            drc_path = drc_output_dir / f"{item['filename']}.drc"
            if not drc_path.exists():
                continue
            new_item = dict(item)
            new_item["mime_type"] = "application/octet-stream"
            new_item["extension"] = "drc"
            new_item["directory"] = drc_output_dir.name
            new_item["url"] = f"./{drc_path.name}"
            new_item["size"] = drc_path.stat().st_size
            result_urls.append(new_item)
        data["result_urls"] = result_urls
        (drc_output_dir / "data.json").write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def simplify_obj(source_path: Path, output_path: Path, params: dict, target_ratio: float):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    ms = ml.MeshSet()
    ms.load_new_mesh(str(source_path))
    mesh = ms.current_mesh()
    original_faces = int(mesh.face_number())
    original_vertices = int(mesh.vertex_number())
    target_faces = max(1, int(original_faces * target_ratio))
    ms.meshing_decimation_quadric_edge_collapse(
        targetfacenum=target_faces,
        targetperc=0.0,
        qualitythr=float(params["qualitythr"]),
        preserveboundary=bool(params["preserveboundary"]),
        boundaryweight=1.0,
        preservenormal=bool(params["preservenormal"]),
        preservetopology=bool(params["preservetopology"]),
        optimalplacement=True,
        planarquadric=False,
        planarweight=0.001,
        qualityweight=False,
        autoclean=True,
        selected=False,
    )
    current = ms.current_mesh()
    ms.save_current_mesh(str(output_path))
    output_size = output_path.stat().st_size
    return {
        "file": source_path.name,
        "output": str(output_path),
        "target_faces": int(target_faces),
        "orig_faces": original_faces,
        "new_faces": int(current.face_number()),
        "orig_vertices": original_vertices,
        "new_vertices": int(current.vertex_number()),
        "output_size": output_size,
        "output_size_mb": output_size / 1024 / 1024,
    }


def write_output_data_json(source_dir: Path, output_dir: Path):
    candidates = [source_dir / "data.json", source_dir.parent / "data.json"]
    source_json = next((path for path in candidates if path.exists()), None)
    if not source_json:
        return
    data = json.loads(source_json.read_text(encoding="utf-8"))
    data["order_id"] = output_dir.name
    if isinstance(data.get("patient"), dict):
        data["patient"]["name"] = output_dir.name
    result_urls = []
    for item in data.get("result_urls", []):
        obj_path = output_dir / f"{item['filename']}.obj"
        if not obj_path.exists():
            continue
        new_item = dict(item)
        new_item["mime_type"] = "model/obj"
        new_item["extension"] = "obj"
        new_item["directory"] = output_dir.name
        new_item["url"] = f"./{obj_path.name}"
        new_item["size"] = obj_path.stat().st_size
        result_urls.append(new_item)
    data["result_urls"] = result_urls
    (output_dir / "data.json").write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def create_output_zip(output_dir: Path) -> Path:
    zip_path = output_dir.parent / f"{output_dir.name}.zip"
    if zip_path.exists():
        zip_path.unlink()

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(output_dir.iterdir()):
            if not path.is_file():
                continue
            if path.suffix.lower() not in {".obj", ".drc", ".json"}:
                continue
            if path.name.endswith("_preview.obj"):
                continue
            zf.write(path, arcname=path.name)
    return zip_path


def _powershell_exe() -> str | None:
    return shutil.which("pwsh") or shutil.which("powershell")


def try_copy_zip_to_pico_ords(zip_path: Path) -> str:
    ps = _powershell_exe()
    if not ps:
        return "未找到 PowerShell"

    zip_str = str(zip_path).replace("'", "''")
    zip_name = zip_path.name.replace("'", "''")
    script = f"""
$shell = New-Object -ComObject Shell.Application
$current = $shell.Namespace(17).Items() | Where-Object {{ $_.Name -eq 'PICO 4 Ultra' }} | Select-Object -First 1
if (-not $current) {{ Write-Output 'DEVICE_NOT_FOUND'; exit 0 }}
$segments = @('内部共享存储空间','Android','data','com.zcwl.MRSystem','files','Ords')
foreach ($seg in $segments) {{
  $folder = $current.GetFolder
  if (-not $folder) {{ Write-Output 'PATH_NOT_FOUND'; exit 0 }}
  $next = $folder.Items() | Where-Object {{ $_.Name -eq $seg }} | Select-Object -First 1
  if (-not $next) {{ Write-Output 'PATH_NOT_FOUND'; exit 0 }}
  $current = $next
}}
$target = $current.GetFolder
if (-not $target) {{ Write-Output 'PATH_NOT_FOUND'; exit 0 }}
$target.CopyHere('{zip_str}', 16)
$ok = $false
for ($i = 0; $i -lt 30; $i++) {{
  Start-Sleep -Milliseconds 500
  $item = $target.Items() | Where-Object {{ $_.Name -eq '{zip_name}' }} | Select-Object -First 1
  if ($item) {{ $ok = $true; break }}
}}
if ($ok) {{ Write-Output 'COPIED' }} else {{ Write-Output 'COPY_SENT' }}
"""
    try:
        result = subprocess.run([ps, "-NoProfile", "-Command", script], capture_output=True, text=True, timeout=45)
        status = (result.stdout or "").strip().splitlines()
        return status[-1] if status else (result.stderr.strip() or "UNKNOWN")
    except Exception as exc:
        return f"复制失败: {exc}"


def main():
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    MeshOptApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
