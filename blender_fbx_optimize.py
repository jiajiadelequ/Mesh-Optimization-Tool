import json
import sys
import traceback
from pathlib import Path

import bpy
from mathutils import Matrix, Vector


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


def apply_remesh(obj, params: dict):
    mesh = obj.data
    mesh.calc_loop_triangles()
    before_triangles = len(mesh.loop_triangles)
    before_vertices = len(mesh.vertices)
    voxel_size = max(0.0001, float(params.get("voxel_size", 0.05)))
    modifier = obj.modifiers.new(name="MeshOptRemesh", type="REMESH")
    modifier.mode = "VOXEL"
    if hasattr(modifier, "voxel_size"):
        modifier.voxel_size = voxel_size
    if hasattr(modifier, "adaptivity"):
        modifier.adaptivity = 0.0
    if hasattr(modifier, "use_remove_disconnected"):
        modifier.use_remove_disconnected = True
    if hasattr(modifier, "use_smooth_shade"):
        modifier.use_smooth_shade = False
    bpy.context.view_layer.objects.active = obj
    bpy.ops.object.modifier_apply(modifier=modifier.name)
    obj.data.calc_loop_triangles()
    return {
        "algorithm": "REMESH",
        "triangle_count_before": before_triangles,
        "triangle_count_after": len(obj.data.loop_triangles),
        "vertex_count_before": before_vertices,
        "vertex_count_after": len(obj.data.vertices),
        "material_slots": len(obj.data.materials),
        "ratio": None,
        "iterations": None,
        "angle_limit": None,
        "voxel_size": voxel_size,
        "use_symmetry": False,
        "symmetry_axis": None,
        "triangulate": None,
    }


def create_box_proxy_from_bounds(source_obj, local_center: Vector, local_size: Vector, name: str):
    bpy.ops.mesh.primitive_cube_add(size=2.0, location=(0.0, 0.0, 0.0))
    proxy_obj = bpy.context.active_object
    proxy_obj.name = name
    half_extents = Vector(
        (
            max(local_size.x * 0.5, 0.0005),
            max(local_size.y * 0.5, 0.0005),
            max(local_size.z * 0.5, 0.0005),
        )
    )
    proxy_obj.matrix_world = (
        source_obj.matrix_world
        @ Matrix.Translation(local_center)
        @ Matrix.Diagonal((half_extents.x, half_extents.y, half_extents.z, 1.0))
    )
    return proxy_obj


def iter_connected_component_bounds(source_obj):
    mesh = source_obj.data
    polygons = mesh.polygons
    vertices = mesh.vertices
    vert_to_polys = [set() for _ in range(len(vertices))]
    coord_to_polys = {}
    for poly_index, poly in enumerate(polygons):
        for vertex_index in poly.vertices:
            vert_to_polys[vertex_index].add(poly_index)
            coord_key = tuple(round(value, 6) for value in vertices[vertex_index].co)
            coord_to_polys.setdefault(coord_key, set()).add(poly_index)

    remaining = set(range(len(polygons)))
    while remaining:
        start = remaining.pop()
        stack = [start]
        component_polys = {start}
        component_verts = set(polygons[start].vertices)
        while stack:
            poly_index = stack.pop()
            for vertex_index in polygons[poly_index].vertices:
                neighbor_set = set(vert_to_polys[vertex_index])
                coord_key = tuple(round(value, 6) for value in vertices[vertex_index].co)
                neighbor_set.update(coord_to_polys.get(coord_key, set()))
                for neighbor_poly_index in neighbor_set:
                    if neighbor_poly_index in remaining:
                        remaining.remove(neighbor_poly_index)
                        component_polys.add(neighbor_poly_index)
                        stack.append(neighbor_poly_index)
                        component_verts.update(polygons[neighbor_poly_index].vertices)
        points = [vertices[index].co.copy() for index in component_verts]
        min_corner = Vector(
            (
                min(point.x for point in points),
                min(point.y for point in points),
                min(point.z for point in points),
            )
        )
        max_corner = Vector(
            (
                max(point.x for point in points),
                max(point.y for point in points),
                max(point.z for point in points),
            )
        )
        local_size = max_corner - min_corner
        yield {
            "poly_count": len(component_polys),
            "vert_count": len(component_verts),
            "min_corner": min_corner,
            "max_corner": max_corner,
            "local_center": (min_corner + max_corner) * 0.5,
            "local_size": local_size,
            "max_dim": max(local_size.x, local_size.y, local_size.z),
        }


def join_objects(objects, name: str):
    if not objects:
        return None
    bpy.ops.object.select_all(action="DESELECT")
    for obj in objects:
        obj.select_set(True)
    bpy.context.view_layer.objects.active = objects[0]
    bpy.ops.object.join()
    merged = bpy.context.view_layer.objects.active
    merged.name = name
    if merged.data:
        merged.data.name = f"{name}_mesh"
    return merged


def build_box_proxy_scene(params: dict):
    min_size = max(0.0, float(params.get("min_size", 0.0)))
    source_mesh_objects = [obj for obj in bpy.data.objects if obj.type == "MESH"]
    processed = []
    skipped = []
    proxies = []

    for obj in source_mesh_objects:
        if is_skinned_object(obj):
            skipped.append({"object": obj.name, "reason": "skinned_mesh"})
            continue
        obj.data.calc_loop_triangles()
        component_index = 0
        for component in iter_connected_component_bounds(obj):
            component_index += 1
            if component["max_dim"] < min_size:
                skipped.append({"object": f"{obj.name}#{component_index}", "reason": "below_min_size"})
                continue
            proxy_obj = create_box_proxy_from_bounds(
                obj,
                component["local_center"],
                component["local_size"],
                f"{obj.name}_box_{component_index}",
            )
            proxies.append(proxy_obj)
            processed.append(
                {
                    "object": f"{obj.name}#{component_index}",
                    "triangle_count_before": component["poly_count"],
                    "vertex_count_before": component["vert_count"],
                    "algorithm": "BOX_PROXY",
                    "triangle_count_after": 12,
                    "vertex_count_after": 8,
                    "min_size": min_size,
                    "max_dim": round(component["max_dim"], 5),
                }
            )

    for obj in source_mesh_objects:
        if obj.name in bpy.data.objects:
            bpy.data.objects.remove(obj, do_unlink=True)

    merged_proxy = join_objects(proxies, "BoxProxy")
    if merged_proxy is None:
        raise ValueError("没有生成任何 Box Proxy，请检查模型是否全被过滤掉了")

    merged_proxy.data.calc_loop_triangles()
    return {
        "processed_meshes": processed,
        "skipped_meshes": skipped,
        "proxy_box_count": len(proxies),
        "proxy_triangle_count": len(merged_proxy.data.loop_triangles),
        "proxy_vertex_count": len(merged_proxy.data.vertices),
    }


def optimize_scene(config: dict):
    source_path = Path(config["source_path"])
    output_path = Path(config["output_path"])
    params = dict(config.get("params", {}))

    reset_scene()
    import_model(source_path)
    before_stats = collect_scene_stats()

    bpy.ops.object.select_all(action="DESELECT")
    algorithm = str(params.get("algorithm", "COLLAPSE")).upper()
    processed = []
    skipped = []
    extra_result = {}
    if algorithm == "BOX_PROXY":
        proxy_result = build_box_proxy_scene(params)
        processed = proxy_result["processed_meshes"]
        skipped = proxy_result["skipped_meshes"]
        extra_result = {
            "proxy_box_count": proxy_result["proxy_box_count"],
            "proxy_triangle_count": proxy_result["proxy_triangle_count"],
            "proxy_vertex_count": proxy_result["proxy_vertex_count"],
        }
    else:
        mesh_objects = [obj for obj in bpy.data.objects if obj.type == "MESH"]
        for obj in mesh_objects:
            if is_skinned_object(obj):
                skipped.append({"object": obj.name, "reason": "skinned_mesh"})
                continue
            if obj.data.users > 1:
                obj.data = obj.data.copy()
            if algorithm == "REMESH":
                stats = apply_remesh(obj, params)
            else:
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
        **extra_result,
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
