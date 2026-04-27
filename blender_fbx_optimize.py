import json
import sys
import traceback
from pathlib import Path

import bpy


OBJ_SUFFIX = ".obj"
FBX_SUFFIX = ".fbx"


def report_and_exit(report_path: Path, payload: dict, exit_code: int = 0):
    report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    raise SystemExit(exit_code)


def load_config() -> tuple[dict, Path]:
    argv = sys.argv
    if "--" not in argv:
        raise RuntimeError("Missing config path")
    config_path = Path(argv[argv.index("--") + 1])
    payload = json.loads(config_path.read_text(encoding="utf-8"))
    return payload, Path(payload["report_path"])


def reset_scene():
    bpy.ops.wm.read_factory_settings(use_empty=True)


def import_model(path: Path):
    suffix = path.suffix.lower()
    if suffix == FBX_SUFFIX:
        bpy.ops.import_scene.fbx(filepath=str(path))
        return
    if suffix == OBJ_SUFFIX:
        if hasattr(bpy.ops.wm, "obj_import"):
            bpy.ops.wm.obj_import(
                filepath=str(path),
                forward_axis="NEGATIVE_Z",
                up_axis="Y",
                validate_meshes=True,
            )
        else:
            bpy.ops.import_scene.obj(
                filepath=str(path),
                axis_forward="-Z",
                axis_up="Y",
            )
        return
    raise ValueError(f"Unsupported input format: {path.suffix}")


def export_model(path: Path):
    suffix = path.suffix.lower()
    if suffix == FBX_SUFFIX:
        bpy.ops.export_scene.fbx(
            filepath=str(path),
            use_selection=False,
            add_leaf_bones=False,
            bake_anim=False,
            path_mode="RELATIVE",
            embed_textures=False,
            use_mesh_modifiers=True,
            mesh_smooth_type="FACE",
            axis_forward="-Z",
            axis_up="Y",
        )
        return
    if suffix == OBJ_SUFFIX:
        if hasattr(bpy.ops.wm, "obj_export"):
            bpy.ops.wm.obj_export(
                filepath=str(path),
                export_selected_objects=False,
                export_materials=True,
                export_uv=True,
                export_normals=True,
                export_triangulated_mesh=False,
                path_mode="RELATIVE",
                forward_axis="NEGATIVE_Z",
                up_axis="Y",
            )
        else:
            bpy.ops.export_scene.obj(
                filepath=str(path),
                use_selection=False,
                use_materials=True,
                use_uvs=True,
                use_normals=True,
                axis_forward="-Z",
                axis_up="Y",
                path_mode="RELATIVE",
            )
        return
    raise ValueError(f"Unsupported output format: {path.suffix}")


def is_skinned_object(obj) -> bool:
    if obj.find_armature() is not None:
        return True
    if any(mod.type == "ARMATURE" for mod in obj.modifiers):
        return True
    if obj.parent and obj.parent.type == "ARMATURE":
        return True
    return False


def collect_scene_stats():
    mesh_objects = [obj for obj in bpy.data.objects if obj.type == "MESH"]
    unique_meshes = {obj.data for obj in mesh_objects}
    triangles = 0
    vertices = 0
    material_slots = 0
    skinned = 0
    for obj in mesh_objects:
        mesh = obj.data
        mesh.calc_loop_triangles()
        triangles += len(mesh.loop_triangles)
        vertices += len(mesh.vertices)
        material_slots += len(mesh.materials)
        if is_skinned_object(obj):
            skinned += 1
    return {
        "object_count": len(bpy.data.objects),
        "renderer_count": len(mesh_objects),
        "mesh_count": len(unique_meshes),
        "triangle_count": triangles,
        "vertex_count": vertices,
        "material_slot_count": material_slots,
        "skinned_mesh_count": skinned,
    }


def apply_decimate(obj, params: dict):
    mesh = obj.data
    mesh.calc_loop_triangles()
    before_triangles = len(mesh.loop_triangles)
    before_vertices = len(mesh.vertices)
    algorithm = str(params.get("algorithm", "COLLAPSE")).upper()
    modifier = obj.modifiers.new(name="MeshOptDecimate", type="DECIMATE")
    modifier.decimate_type = algorithm
    ratio = None
    iterations = None
    angle_limit = None
    use_symmetry = bool(params.get("use_symmetry", False))
    symmetry_axis = str(params.get("symmetry_axis", "X")).upper()
    triangulate = bool(params.get("triangulate", True))
    if algorithm == "COLLAPSE":
        ratio = max(0.0, min(1.0, float(params.get("ratio", 0.1))))
        modifier.ratio = ratio
        if hasattr(modifier, "use_symmetry"):
            modifier.use_symmetry = use_symmetry
        if hasattr(modifier, "symmetry_axis") and symmetry_axis in {"X", "Y", "Z"}:
            modifier.symmetry_axis = symmetry_axis
        modifier.use_collapse_triangulate = triangulate
    elif algorithm == "UNSUBDIV":
        iterations = max(1, min(32, int(params.get("iterations", 2))))
        modifier.iterations = iterations
    elif algorithm == "DISSOLVE":
        angle_limit = float(params.get("angle_limit", 5.0))
        modifier.angle_limit = angle_limit * 3.141592653589793 / 180.0
        if hasattr(modifier, "delimit"):
            modifier.delimit = {"NORMAL", "MATERIAL", "SEAM", "SHARP", "UV"}
    else:
        raise ValueError(f"Unsupported decimate algorithm: {algorithm}")
    bpy.context.view_layer.objects.active = obj
    bpy.ops.object.modifier_apply(modifier=modifier.name)
    obj.data.calc_loop_triangles()
    return {
        "algorithm": algorithm,
        "triangle_count_before": before_triangles,
        "triangle_count_after": len(obj.data.loop_triangles),
        "vertex_count_before": before_vertices,
        "vertex_count_after": len(obj.data.vertices),
        "material_slots": len(obj.data.materials),
        "ratio": ratio,
        "iterations": iterations,
        "angle_limit": angle_limit,
        "use_symmetry": use_symmetry,
        "symmetry_axis": symmetry_axis,
        "triangulate": triangulate,
    }


def optimize_scene(config: dict):
    source_path = Path(config["source_path"])
    output_path = Path(config["output_path"])
    params = dict(config.get("params", {}))

    reset_scene()
    import_model(source_path)
    before_stats = collect_scene_stats()

    bpy.ops.object.select_all(action="DESELECT")
    processed = []
    skipped = []
    mesh_objects = [obj for obj in bpy.data.objects if obj.type == "MESH"]
    for obj in mesh_objects:
        if is_skinned_object(obj):
            skipped.append({"object": obj.name, "reason": "skinned_mesh"})
            continue
        if obj.data.users > 1:
            obj.data = obj.data.copy()
        stats = apply_decimate(obj, params)
        processed.append({"object": obj.name, **stats})

    output_path.parent.mkdir(parents=True, exist_ok=True)
    export_model(output_path)
    after_stats = collect_scene_stats()
    return {
        "ok": True,
        "mode": "optimize",
        "input": str(source_path),
        "output": str(output_path),
        "before": before_stats,
        "after": after_stats,
        "processed_meshes": processed,
        "skipped_meshes": skipped,
    }


def inspect_scene(config: dict):
    source_path = Path(config["source_path"])
    reset_scene()
    import_model(source_path)
    return {
        "ok": True,
        "mode": "inspect",
        "input": str(source_path),
        "stats": collect_scene_stats(),
    }


def main():
    try:
        config, report_path = load_config()
        mode = config["mode"]
        if mode == "inspect":
            payload = inspect_scene(config)
        elif mode == "optimize":
            payload = optimize_scene(config)
        else:
            raise ValueError(f"Unsupported mode: {mode}")
        report_and_exit(report_path, payload, 0)
    except Exception as exc:
        report_path = Path.cwd() / "mesh_opt_blender_error_report.json"
        if "report_path" in locals():
            report_path = locals()["report_path"]
        report_and_exit(
            report_path,
            {"ok": False, "error": str(exc), "traceback": traceback.format_exc()},
            1,
        )


if __name__ == "__main__":
    main()
