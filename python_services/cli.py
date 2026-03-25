from __future__ import annotations

import json
import sys
from pathlib import Path

from bootstrap import bootstrap_vendor

bootstrap_vendor()

from processors.contracts import ProcessRequest
from processors.dispatch import dispatch


def main() -> int:
    if len(sys.argv) < 2:
        print("ERROR|Missing manifest path")
        return 1

    manifest_path = Path(sys.argv[1]).resolve()
    if not manifest_path.is_file():
        print(f"ERROR|Manifest not found: {manifest_path}")
        return 1

    payload = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    processor = payload.get("processor")
    output_path = payload.get("output_path")
    if not output_path:
        print("ERROR|Manifest missing output_path")
        return 1

    if processor:
        request = ProcessRequest(
            input_paths=[Path(path).resolve() for path in payload.get("input_paths", [])],
            output_path=Path(output_path).resolve(),
            template_path=Path(payload["template_path"]).resolve() if payload.get("template_path") else None,
            options=payload.get("options", {}),
        )
        try:
            result = dispatch(str(processor), request)
        except Exception as exc:
            print(json.dumps({"success": False, "error": str(exc)}, ensure_ascii=True))
            return 1

        print(
            json.dumps(
                {
                    "success": result.success,
                    "label": result.label,
                    "output_path": str(result.output_path),
                    "metadata": result.metadata,
                },
                ensure_ascii=True,
            )
        )
        return 0

    print(f"OUTPUT|{Path(output_path).name}|python scaffold placeholder")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
