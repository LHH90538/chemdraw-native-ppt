"""Create a standalone ChemDraw-PPT project. Python standard library only."""
import argparse
import json
from pathlib import Path
import shutil


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--destination', required=True, type=Path)
    ap.add_argument('--mode', choices=['empty', 'example'], default='empty')
    a = ap.parse_args()
    target = a.destination.resolve()
    if target.exists() and any(target.iterdir()):
        ap.error('Destination must be new or empty; existing projects are not overwritten.')
    skill = Path(__file__).resolve().parents[1]
    target.mkdir(parents=True, exist_ok=True)
    if a.mode == 'example':
        shutil.copytree(skill / 'assets' / 'synthetic-example', target, dirs_exist_ok=True)
    else:
        (target / 'molecules').mkdir()
        project = {
            'name': 'Chemical_Figure', 'canvasWidth': 960, 'canvasHeight': 540,
            'slideWidth': 960, 'slideHeight': 540, 'slideCount': 1,
            'template': None, 'molecules': [], 'elements': [], 'textEdits': [],
            'rasterExceptions': [],
        }
        (target / 'project.json').write_text(json.dumps(project, indent=2), encoding='utf-8')
    shutil.copy2(skill / 'scripts' / 'build.ps1', target / 'build.ps1')
    (target / 'start.cmd').write_text(
        '@echo off\ncd /d "%~dp0"\npowershell.exe -NoProfile -STA -File "%~dp0build.ps1"\n'
        'echo.\npause\n', encoding='ascii', newline='\r\n')
    print(str(target))


if __name__ == '__main__':
    main()
