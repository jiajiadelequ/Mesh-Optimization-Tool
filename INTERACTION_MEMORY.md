# Interaction Memory

Updated: 2026-04-17

## User Intent

- The user asked to save interaction memory in the current directory.
- Recent technical focus was troubleshooting why the mesh optimization tool did not show OBJ files from `F:/DicomDatas/A/model`.

## Relevant Project Context

- Workspace: `F:\mesh_opt_tool`
- Main script: `F:\mesh_opt_tool\mesh_opt_tool.py`
- Settings file: `F:\mesh_opt_tool\mesh_opt_tool_settings.json`
- Data directory discussed: `F:\DicomDatas\A\model`

## Prior Investigation Summary

- `mesh_opt_tool.py` defines `list_obj_files(directory)` and only includes files in the selected directory whose suffix is `.obj`.
- The code path for model loading is `load_models_async()` -> `_load_models_worker()` -> `list_obj_files(source_dir)`.
- UI completion path is `_poll_queue()` -> `_apply_loaded_models(payload)`.
- A direct filesystem check against `F:\DicomDatas\A\model` showed:
  - the directory exists
  - it is a directory
  - it contains 62 `.obj` files
- That means the current source code should detect OBJ files in that folder if the running process is actually using this source and the runtime source directory value is correct.

## Working Hypothesis

- If the app still shows `已加载0个模型` while `F:/DicomDatas/A/model` is selected, the most likely causes are:
  - the user is running a different copy / older build than the current source file
  - the runtime source directory value is being overwritten or differs from what is visible in the UI

## Suggested Next Debug Step

- Add temporary diagnostics to the app status/logging so it prints:
  - the exact resolved source directory used at runtime
  - whether that path exists
  - the counted number of `.obj` files before populating the UI

