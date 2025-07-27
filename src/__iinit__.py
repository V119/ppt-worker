
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent.absolute()
template_dir = PROJECT_ROOT / "data" / "template"
output_dir = PROJECT_ROOT / "data" / "output"
output_dir.mkdir(parents=True, exist_ok=True)