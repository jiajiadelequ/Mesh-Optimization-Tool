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
MAX_RECENT_PRESETS = 10


@dataclass
class ModelMeta:
    path: Path
    display_name: str
    group_name: str
    source_size: int


def list_obj_files(directory: Path):
    return sorted(
        (path for path in directory.iterdir() if path.is_file() and path.suffix.lower() == ".obj"),
        key=lambda p: p.stat().st_size,
        reverse=True,
    )


class ModelRow:
    def __init__(self, meta: ModelMeta):
        self.meta = meta
        self.selected = True
        self.ratio = "0.10"
        self.quality = "0.5"
        self.preserve_boundary = True
        self.preserve_normal = True
        self.preserve_topology = True
        self.result_text = "未处理"

    def reset_defaults(self):
        self.ratio = "0.10"
        self.quality = "0.5"
        self.preserve_boundary = True
        self.preserve_normal = True
        self.preserve_topology = True
        self.result_text = "未处理"

    def current_params(self):
        ratio = float(self.ratio.strip())
        quality = float(self.quality.strip())
        return {
            "ratio": ratio,
            "qualitythr": quality,
            "preserveboundary": self.preserve_boundary,
            "preservenormal": self.preserve_normal,
            "preservetopology": self.preserve_topology,
        }

    def export_state(self):
        return {
            "selected": self.selected,
            "ratio": self.ratio.strip(),
            "quality": self.quality.strip(),
            "preserve_boundary": self.preserve_boundary,
            "preserve_normal": self.preserve_normal,
            "preserve_topology": self.preserve_topology,
        }

    def apply_state(self, payload: dict):
        self.selected = bool(payload.get("selected", True))
        self.ratio = str(payload.get("ratio", "0.10"))
        self.quality = str(payload.get("quality", "0.5"))
        self.preserve_boundary = bool(payload.get("preserve_boundary", True))
        self.preserve_normal = bool(payload.get("preserve_normal", True))
        self.preserve_topology = bool(payload.get("preserve_topology", True))


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
        self.var_editor_selected = tk.BooleanVar(value=True)
        self.var_editor_ratio = tk.StringVar(value="0.10")
        self.var_editor_quality = tk.StringVar(value="0.5")
        self.var_editor_preserve_boundary = tk.BooleanVar(value=True)
        self.var_editor_preserve_normal = tk.BooleanVar(value=True)
        self.var_editor_preserve_topology = tk.BooleanVar(value=True)
        self.var_editor_name = tk.StringVar(value="未选择模型")
        self.var_editor_group = tk.StringVar(value="-")
        self.var_editor_size = tk.StringVar(value="-")
        self.var_editor_result = tk.StringVar(value="结果: 未处理")
        self.recent_profile_paths = []
        self.pending_profile_payload = None
        self.active_row = None

        self._load_settings()
        self._build_menu()
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
        ttk.Button(controls, text="保存参数JSON", command=self.save_profile_as).grid(row=0, column=5, padx=(8, 0))
        ttk.Button(controls, text="加载参数JSON", command=self.load_profile_from_dialog).grid(row=0, column=6, padx=(8, 0))

        center = ttk.Panedwindow(self.root, orient="horizontal")
        center.grid(row=1, column=0, sticky="nsew", padx=10)

        left_panel = ttk.Frame(center, padding=(0, 0, 10, 0))
        left_panel.columnconfigure(0, weight=1)
        left_panel.rowconfigure(1, weight=1)
        center.add(left_panel, weight=3)

        list_actions = ttk.Frame(left_panel)
        list_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(list_actions, text="全选", command=self.select_all_rows).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(list_actions, text="全不选", command=self.clear_all_rows).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(list_actions, text="反选", command=self.invert_row_selection).grid(row=0, column=2, padx=(0, 6))

        columns = ("selected", "name", "group", "size", "result")
        self.model_tree = ttk.Treeview(left_panel, columns=columns, show="headings", selectmode="browse")
        self.model_tree.heading("selected", text="处理")
        self.model_tree.heading("name", text="模型")
        self.model_tree.heading("group", text="分组")
        self.model_tree.heading("size", text="大小(MB)")
        self.model_tree.heading("result", text="结果")
        self.model_tree.column("selected", width=56, anchor="center", stretch=False)
        self.model_tree.column("name", width=240, anchor="w")
        self.model_tree.column("group", width=120, anchor="w")
        self.model_tree.column("size", width=90, anchor="e", stretch=False)
        self.model_tree.column("result", width=220, anchor="w")
        self.model_tree.grid(row=1, column=0, sticky="nsew")
        self.model_tree.bind("<<TreeviewSelect>>", self._on_model_tree_select)

        list_scroll = ttk.Scrollbar(left_panel, orient="vertical", command=self.model_tree.yview)
        self.model_tree.configure(yscrollcommand=list_scroll.set)
        list_scroll.grid(row=1, column=1, sticky="ns")

        right_panel = ttk.LabelFrame(center, text="当前模型参数", padding=10)
        right_panel.columnconfigure(1, weight=1)
        center.add(right_panel, weight=2)

        ttk.Label(right_panel, text="模型").grid(row=0, column=0, sticky="w")
        ttk.Label(right_panel, textvariable=self.var_editor_name).grid(row=0, column=1, sticky="w")
        ttk.Label(right_panel, text="分组").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Label(right_panel, textvariable=self.var_editor_group).grid(row=1, column=1, sticky="w", pady=(6, 0))
        ttk.Label(right_panel, text="大小").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Label(right_panel, textvariable=self.var_editor_size).grid(row=2, column=1, sticky="w", pady=(6, 0))

        param_box = ttk.LabelFrame(right_panel, text="减面参数", padding=10)
        param_box.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Checkbutton(param_box, text="纳入批量处理", variable=self.var_editor_selected, command=self._on_editor_selection_change).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(param_box, text="比例").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(param_box, textvariable=self.var_editor_ratio, width=10).grid(row=1, column=1, sticky="w", pady=(10, 0))
        ttk.Label(param_box, text="Quality").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(param_box, textvariable=self.var_editor_quality, width=10).grid(row=2, column=1, sticky="w", pady=(8, 0))
        ttk.Checkbutton(param_box, text="Boundary", variable=self.var_editor_preserve_boundary).grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Checkbutton(param_box, text="Normal", variable=self.var_editor_preserve_normal).grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Checkbutton(param_box, text="Topology", variable=self.var_editor_preserve_topology).grid(row=5, column=0, columnspan=2, sticky="w", pady=(4, 0))

        editor_actions = ttk.Frame(right_panel)
        editor_actions.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(editor_actions, text="应用到当前模型", command=self.apply_editor_to_active_row).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(editor_actions, text="应用到已勾选", command=self.apply_editor_to_selected_rows).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(editor_actions, text="重置当前参数", command=self.reset_active_row_defaults).grid(row=0, column=2)

        preview_actions = ttk.Frame(right_panel)
        preview_actions.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(preview_actions, text="减面当前模型", command=self.start_active_row_job).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(preview_actions, text="预览当前模型", command=self.preview_active_row_job).grid(row=0, column=1)

        ttk.Label(right_panel, textvariable=self.var_editor_result, wraplength=320, justify="left").grid(row=6, column=0, columnspan=2, sticky="w", pady=(12, 0))

        bottom = ttk.Frame(self.root, padding=10)
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.var_status).grid(row=0, column=0, sticky="w")

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="加载参数 JSON...", command=self.load_profile_from_dialog)
        file_menu.add_command(label="保存参数 JSON...", command=self.save_profile_as)
        file_menu.add_separator()
        self.recent_menu = tk.Menu(file_menu, tearoff=False)
        file_menu.add_cascade(label="最近打开的模型", menu=self.recent_menu)
        menubar.add_cascade(label="文件", menu=file_menu)
        self.root.config(menu=menubar)
        self._refresh_recent_menu()

    def _refresh_recent_menu(self):
        self.recent_menu.delete(0, "end")
        valid_paths = []
        for raw_path in self.recent_profile_paths:
            if not raw_path:
                continue
            path = Path(raw_path)
            if path.exists() and path.suffix.lower() == ".json":
                valid_paths.append(str(path))
        self.recent_profile_paths = valid_paths[:MAX_RECENT_PRESETS]
        if not self.recent_profile_paths:
            self.recent_menu.add_command(label="暂无", state="disabled")
            return
        for raw_path in self.recent_profile_paths:
            path = Path(raw_path)
            self.recent_menu.add_command(label=path.name, command=lambda p=raw_path: self.load_profile_file(Path(p)))

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

    def _row_tree_values(self, row: ModelRow):
        return (
            "是" if row.selected else "否",
            row.meta.display_name,
            row.meta.group_name or "-",
            f"{row.meta.source_size / 1024 / 1024:.2f}",
            row.result_text,
        )

    def _refresh_model_tree(self):
        if not hasattr(self, "model_tree"):
            return
        existing = set(self.model_tree.get_children())
        active_id = self.active_row.meta.path.name if self.active_row else None
        for row in self.rows:
            item_id = row.meta.path.name
            values = self._row_tree_values(row)
            if item_id in existing:
                self.model_tree.item(item_id, values=values)
            else:
                self.model_tree.insert("", "end", iid=item_id, values=values)
        for item_id in existing - {row.meta.path.name for row in self.rows}:
            self.model_tree.delete(item_id)
        if active_id and active_id in self.model_tree.get_children():
            self.model_tree.selection_set(active_id)

    def _load_editor_from_row(self, row: ModelRow | None):
        if row is None:
            self.var_editor_name.set("未选择模型")
            self.var_editor_group.set("-")
            self.var_editor_size.set("-")
            self.var_editor_selected.set(True)
            self.var_editor_ratio.set("0.10")
            self.var_editor_quality.set("0.5")
            self.var_editor_preserve_boundary.set(True)
            self.var_editor_preserve_normal.set(True)
            self.var_editor_preserve_topology.set(True)
            self.var_editor_result.set("结果: 未处理")
            return
        self.var_editor_name.set(f"{row.meta.display_name} [{row.meta.path.name}]")
        self.var_editor_group.set(row.meta.group_name or "-")
        self.var_editor_size.set(f"{row.meta.source_size / 1024 / 1024:.2f} MB")
        self.var_editor_selected.set(row.selected)
        self.var_editor_ratio.set(row.ratio)
        self.var_editor_quality.set(row.quality)
        self.var_editor_preserve_boundary.set(row.preserve_boundary)
        self.var_editor_preserve_normal.set(row.preserve_normal)
        self.var_editor_preserve_topology.set(row.preserve_topology)
        self.var_editor_result.set(f"结果: {row.result_text}")

    def _save_editor_to_active_row(self):
        if self.active_row is None:
            return
        self.active_row.selected = self.var_editor_selected.get()
        self.active_row.ratio = self.var_editor_ratio.get().strip()
        self.active_row.quality = self.var_editor_quality.get().strip()
        self.active_row.preserve_boundary = self.var_editor_preserve_boundary.get()
        self.active_row.preserve_normal = self.var_editor_preserve_normal.get()
        self.active_row.preserve_topology = self.var_editor_preserve_topology.get()
        self._refresh_model_tree()
        self._sync_select_all_state()

    def _set_active_row(self, row: ModelRow | None):
        self._save_editor_to_active_row()
        self.active_row = row
        self._load_editor_from_row(row)
        if row is not None and hasattr(self, "model_tree"):
            self.model_tree.selection_set(row.meta.path.name)

    def _on_model_tree_select(self, _event=None):
        selection = self.model_tree.selection()
        if not selection:
            return
        selected_id = selection[0]
        row = next((item for item in self.rows if item.meta.path.name == selected_id), None)
        if row is not None and row is not self.active_row:
            self._set_active_row(row)

    def _on_editor_selection_change(self):
        if self.active_row is None:
            return
        self.active_row.selected = self.var_editor_selected.get()
        self._refresh_model_tree()
        self._sync_select_all_state()

    def select_all_rows(self):
        for row in self.rows:
            row.selected = True
        self.var_select_all.set(True)
        self._refresh_model_tree()
        self._load_editor_from_row(self.active_row)

    def clear_all_rows(self):
        for row in self.rows:
            row.selected = False
        self.var_select_all.set(False)
        self._refresh_model_tree()
        self._load_editor_from_row(self.active_row)

    def invert_row_selection(self):
        for row in self.rows:
            row.selected = not row.selected
        self._sync_select_all_state()
        self._refresh_model_tree()
        self._load_editor_from_row(self.active_row)

    def apply_editor_to_active_row(self):
        self._save_editor_to_active_row()
        self.var_status.set("已应用当前参数到当前模型")

    def apply_editor_to_selected_rows(self):
        if not self.rows:
            return
        self._save_editor_to_active_row()
        selected_rows = [row for row in self.rows if row.selected]
        if not selected_rows:
            messagebox.showwarning(APP_TITLE, "请先勾选至少一个模型")
            return
        template = {
            "ratio": self.var_editor_ratio.get().strip(),
            "quality": self.var_editor_quality.get().strip(),
            "preserve_boundary": self.var_editor_preserve_boundary.get(),
            "preserve_normal": self.var_editor_preserve_normal.get(),
            "preserve_topology": self.var_editor_preserve_topology.get(),
        }
        for row in selected_rows:
            row.ratio = template["ratio"]
            row.quality = template["quality"]
            row.preserve_boundary = template["preserve_boundary"]
            row.preserve_normal = template["preserve_normal"]
            row.preserve_topology = template["preserve_topology"]
        self._refresh_model_tree()
        self._load_editor_from_row(self.active_row)
        self.var_status.set(f"已应用当前参数到 {len(selected_rows)} 个已勾选模型")

    def reset_active_row_defaults(self):
        if self.active_row is None:
            messagebox.showwarning(APP_TITLE, "请先选择一个模型")
            return
        self.active_row.reset_defaults()
        self._load_editor_from_row(self.active_row)
        self._refresh_model_tree()

    def start_active_row_job(self):
        if self.active_row is None:
            messagebox.showwarning(APP_TITLE, "请先选择一个模型")
            return
        self._save_editor_to_active_row()
        self.start_single_job(self.active_row, preview=False)

    def preview_active_row_job(self):
        if self.active_row is None:
            messagebox.showwarning(APP_TITLE, "请先选择一个模型")
            return
        self._save_editor_to_active_row()
        self.start_single_job(self.active_row, preview=True)

    def _capture_profile_payload(self):
        self._save_editor_to_active_row()
        return {
            "version": 1,
            "app": APP_TITLE,
            "source_dir": self.var_source_dir.get(),
            "output_dir": self.var_output_dir.get(),
            "drc_input_dir": self.var_drc_input_dir.get(),
            "drc_output_dir": self.var_drc_output_dir.get(),
            "meshlab_path": self.var_meshlab_path.get(),
            "generate_data_json": self.var_generate_data_json.get(),
            "select_all": self.var_select_all.get(),
            "models": {row.meta.path.name: row.export_state() for row in self.rows},
        }

    def save_profile_as(self):
        initial_dir = self._resolve_initial_dir(self.var_source_dir.get() or Path.cwd())
        path = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
            title="保存参数 JSON",
        )
        if not path:
            return
        profile_path = Path(path)
        profile_path.write_text(json.dumps(self._capture_profile_payload(), indent=2, ensure_ascii=False), encoding="utf-8")
        self._remember_recent_profile(profile_path)
        self.var_status.set(f"参数已保存: {profile_path}")

    def load_profile_from_dialog(self):
        initial_dir = self._resolve_initial_dir(self.var_source_dir.get() or Path.cwd())
        path = filedialog.askopenfilename(
            initialdir=initial_dir,
            filetypes=[("JSON Files", "*.json")],
            title="加载参数 JSON",
        )
        if path:
            self.load_profile_file(Path(path))

    def load_profile_file(self, profile_path: Path):
        try:
            payload = json.loads(profile_path.read_text(encoding="utf-8"))
            if not isinstance(payload, dict) or payload.get("app") != APP_TITLE or "models" not in payload:
                raise ValueError("这不是减面工具的参数 JSON")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"加载参数 JSON 失败: {exc}")
            return
        self._remember_recent_profile(profile_path)
        self._apply_profile_payload(payload, profile_path)

    def _remember_recent_profile(self, profile_path: Path):
        normalized = str(profile_path)
        self.recent_profile_paths = [item for item in self.recent_profile_paths if item != normalized]
        self.recent_profile_paths.insert(0, normalized)
        self.recent_profile_paths = self.recent_profile_paths[:MAX_RECENT_PRESETS]
        self._refresh_recent_menu()
        self.save_settings()

    def _apply_profile_payload(self, payload: dict, profile_path: Path | None = None):
        self.var_source_dir.set(str(payload.get("source_dir", self.var_source_dir.get())))
        self.var_output_dir.set(str(payload.get("output_dir", self.var_output_dir.get())))
        self.var_drc_input_dir.set(str(payload.get("drc_input_dir", self.var_drc_input_dir.get())))
        self.var_drc_output_dir.set(str(payload.get("drc_output_dir", self.var_drc_output_dir.get())))
        self.var_meshlab_path.set(str(payload.get("meshlab_path", self.var_meshlab_path.get())))
        self.var_generate_data_json.set(bool(payload.get("generate_data_json", self.var_generate_data_json.get())))
        self.var_select_all.set(bool(payload.get("select_all", self.var_select_all.get())))
        self.pending_profile_payload = payload
        self.save_settings()
        self.load_models_async()
        suffix = f": {profile_path}" if profile_path else ""
        self.var_status.set(f"正在加载参数{suffix}")

    def _apply_profile_to_rows(self, payload: dict):
        model_state = payload.get("models", {})
        if not isinstance(model_state, dict):
            model_state = {}
        for row in self.rows:
            state = model_state.get(row.meta.path.name) or model_state.get(row.meta.path.stem)
            if state:
                row.apply_state(state)
        self._sync_select_all_state()
        self._refresh_model_tree()
        self._load_editor_from_row(self.active_row)

    def toggle_select_all(self):
        checked = self.var_select_all.get()
        for row in self.rows:
            row.selected = checked
        self._refresh_model_tree()
        self._load_editor_from_row(self.active_row)

    def _sync_select_all_state(self):
        self.var_select_all.set(bool(self.rows) and all(row.selected for row in self.rows))

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
        obj_paths = list_obj_files(source_dir)
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
        self._save_editor_to_active_row()
        selected_rows = [row for row in self.rows if row.selected]
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
            obj_files = list_obj_files(obj_dir)
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
        previous_active_name = self.active_row.meta.path.name if self.active_row else None
        self.rows.clear()

        for index, item in enumerate(payload["rows"]):
            if isinstance(item, ModelMeta):
                row = ModelRow(item)
                self.rows.append(row)

        self.load_running = False
        self._refresh_model_tree()
        if self.rows:
            active_row = next((row for row in self.rows if row.meta.path.name == previous_active_name), self.rows[0])
            self.active_row = None
            self._set_active_row(active_row)
        else:
            self.active_row = None
            self._load_editor_from_row(None)
        self._sync_select_all_state()
        self.var_status.set(payload["status"])
        if self.pending_profile_payload:
            self._apply_profile_to_rows(self.pending_profile_payload)
            self.pending_profile_payload = None
            self.var_status.set(f"{payload['status']} 已应用参数 JSON。")
        self.save_settings()

    def _update_row_result(self, row: ModelRow, result: dict):
        row.result_text = (
            f"target={result['target_faces']}, "
            f"actual={result['new_faces']} faces, "
            f"size={result['output_size_mb']:.2f} MB"
        )
        self._refresh_model_tree()
        if row is self.active_row:
            self._load_editor_from_row(row)

    def save_settings(self):
        payload = {
            "source_dir": self.var_source_dir.get(),
            "output_dir": self.var_output_dir.get(),
            "drc_input_dir": self.var_drc_input_dir.get(),
            "drc_output_dir": self.var_drc_output_dir.get(),
            "meshlab_path": self.var_meshlab_path.get(),
            "generate_data_json": self.var_generate_data_json.get(),
            "recent_profile_paths": self.recent_profile_paths,
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
            self.recent_profile_paths = payload.get("recent_profile_paths", [])
        except Exception:
            self.var_source_dir.set(str(Path.cwd().parent / "obj_out"))
            self.var_output_dir.set(str(Path.cwd().parent / "tool_output"))
            self.var_drc_input_dir.set(str(Path.cwd().parent / "tool_output"))
            self.var_drc_output_dir.set(str(Path.cwd().parent / "tool_output_drc"))
            self.var_meshlab_path.set(DEFAULT_MESHLAB)
            self.recent_profile_paths = []

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
