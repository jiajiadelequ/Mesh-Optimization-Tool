import argparse
import json
from pathlib import Path

from fbx_pipeline import inspect_model, optimize_model


def main():
    parser = argparse.ArgumentParser(description="Run a Blender import/optimize/reimport validation flow.")
    parser.add_argument("--source", required=True, action="append", help="Model file to test. Repeat for multiple files.")
    parser.add_argument("--blender", required=True, help="Path to blender executable.")
    parser.add_argument("--output-dir", default="", help="Optional output directory. Defaults to each FBX file's parent.")
    parser.add_argument("--algorithm", choices=("COLLAPSE", "UNSUBDIV", "DISSOLVE"), default="COLLAPSE")
    parser.add_argument("--ratio", type=float, default=0.1)
    parser.add_argument("--iterations", type=int, default=2)
    parser.add_argument("--angle-limit", type=float, default=5.0)
    parser.add_argument("--use-symmetry", action="store_true")
    parser.add_argument("--symmetry-axis", choices=("X", "Y", "Z"), default="X")
    parser.add_argument("--no-triangulate", action="store_true")
    args = parser.parse_args()

    blender_path = Path(args.blender)
    output_dir = Path(args.output_dir) if args.output_dir else None
    params = {"algorithm": args.algorithm}
    if args.algorithm == "COLLAPSE":
        params.update(
            {
                "ratio": args.ratio,
                "use_symmetry": args.use_symmetry,
                "symmetry_axis": args.symmetry_axis,
                "triangulate": not args.no_triangulate,
            }
        )
    elif args.algorithm == "UNSUBDIV":
        params["iterations"] = args.iterations
    else:
        params["angle_limit"] = args.angle_limit

    reports = []
    for raw_source in args.source:
        source_path = Path(raw_source)
        if not source_path.exists():
            raise SystemExit(f"源文件不存在: {source_path}")
        target_dir = output_dir or source_path.parent
        target_dir.mkdir(parents=True, exist_ok=True)
        output_path = target_dir / f"{source_path.stem}_optimized{source_path.suffix.lower()}"
        before = inspect_model(source_path, blender_path)
        optimize = optimize_model(source_path, output_path, blender_path, params)
        after = inspect_model(output_path, blender_path)
        reports.append(
            {
                "source": str(source_path),
                "output": str(output_path),
                "before": before.get("stats", {}),
                "optimize": optimize,
                "after_reimport": after.get("stats", {}),
            }
        )

    print(json.dumps({"ok": True, "reports": reports}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
