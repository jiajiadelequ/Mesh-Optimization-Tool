# Mesh Batch Optimizer

Run `run_mesh_opt_tool.bat` to open the desktop tool.

Main files:

- `mesh_opt_tool.py`: Tkinter desktop app
- `fbx_pipeline.py`: Blender-backed model optimize/inspect bridge
- `blender_fbx_optimize.py`: Blender-side OBJ/FBX import/process/export script
- `fbx_test_cli.py`: Blender import/optimize/reimport validation entrypoint
- `run_mesh_opt_tool.bat`: launcher

Supported input:

- `.obj`
- `.fbx`

Notes:

- All model optimization now goes through Blender CLI.
- Processing requires a valid Blender executable path in the UI.
- Static meshes are supported in this round.
- Skinned meshes / armature-driven meshes are detected and skipped instead of being rewritten.
- Optimized OBJ files are exported as `*_optimized.obj` when previewing and original filename when batch outputting.
- Optimized FBX files are exported as `*_optimized.fbx`.

Blender 参数:

- `减面算法`：
  `通用减面（Collapse）` 适合大多数模型。
  `规整网格回退（Un-Subdivide）` 适合原本很规整、像细分过的网格。
  `平面合并（Planar Dissolve）` 适合建筑、机械、墙面这类比较平的模型。
- `保留比例`：只在 `Collapse` 里使用，减面后大概还保留多少面数。`0.10` 约等于保留 10%。
- `迭代次数`：只在 `Un-Subdivide` 里使用，数字越大，网格会一层层往回收。
- `平面合并角度`：只在 `Planar Dissolve` 里使用，数字越大，越容易把接近平面的碎面合并掉。
- `开启对称保护`：只在 `Collapse` 里使用，适合左右对称或前后对称的模型。
- `对称方向`：只在 `Collapse` 里使用，指定对称按 `X / Y / Z` 哪个方向判断。
- `保持三角面`：只在 `Collapse` 里使用，通常更适合游戏引擎和跨软件导入。

Test entry:

```powershell
python fbx_test_cli.py --blender "F:\Blender\blender.exe" --source "F:\sample\simple.obj" --source "F:\sample\multi_mesh_multi_mat.fbx"
```

The tool reads source files in read-only mode and writes optimized files only to the output directory you choose.
