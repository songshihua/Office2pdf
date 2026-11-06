"""Dependency-free checks for source layout and local Markdown links."""
from pathlib import Path
from urllib.parse import unquote
import ast
import re

ROOT = Path(__file__).resolve().parents[1]
REQUIRED = ('README.md', 'LICENSE', 'THIRD_PARTY_NOTICES.md', 'WordToPDF.spec',
            'requirements.txt', 'requirements-dev.txt', 'src/word_pdf_app.py',
            'src/modern_ui.py', 'src/converter_engine.py',
            'docs/assets/app-preview.png', 'docs/assets/wps-export-before.png',
            'docs/assets/vector-export-after.png', 'examples/vector-output-demo.pdf')

def main():
    failures = []
    for name in REQUIRED:
        if not (ROOT / name).is_file():failures.append('Missing: ' + name)
    for path in sorted(ROOT.rglob('*.md')):
        if any(part in ('.venv', '_local', 'local-backups', 'build', 'dist') for part in path.parts):continue
        text = path.read_text(encoding='utf-8')
        for link in re.findall(r'!?\[[^\]]*\]\(([^)]+)\)', text):
            link = link.split(' "', 1)[0].strip('<>')
            if link.startswith(('http:', 'https:', '#', 'mailto:', '../../issues')):continue
            target = (path.parent / unquote(link.split('#',1)[0])).resolve()
            if not target.exists():failures.append(f'Broken link in {path.relative_to(ROOT)}: {link}')
    for folder in ('src', 'scripts'):
        for path in (ROOT / folder).glob('*.py'):
            try:ast.parse(path.read_text(encoding='utf-8'), filename=str(path))
            except SyntaxError as exc:failures.append(str(exc))
    for path in (ROOT / 'docs/assets').glob('*.png'):
        if path.read_bytes()[:8] != b'\x89PNG\r\n\x1a\n':failures.append('Invalid PNG: '+path.name)
    example = ROOT / 'examples/vector-output-demo.pdf'
    if example.exists() and not example.read_bytes().startswith(b'%PDF-'):failures.append('Invalid example PDF')
    if failures:raise SystemExit('\n'.join(failures))
    print('Project layout, Markdown links, Python syntax and example assets: OK')

if __name__ == '__main__':
    main()
