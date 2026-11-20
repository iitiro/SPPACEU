from pathlib import Path
import json
import sys
import types
import uuid

ROOT = Path(__file__).resolve().parents[1]


def notebook_namespace(filename, as_main=False):
    name = "release_notebook_" + uuid.uuid4().hex
    module = types.ModuleType(name)
    module.__file__ = str(ROOT / filename)
    sys.modules[name] = module
    if as_main:
        module.__dict__["__name__"] = "__main__"
    # Default execution exercises definitions without live entry points.
    for index, cell in enumerate(json.loads((ROOT / filename).read_text())["cells"]):
        if cell["cell_type"] == "code":
            exec(compile("".join(cell["source"]), f"{filename}:cell{index}", "exec"), module.__dict__)
    return module.__dict__
