"""Package a built EXE with its corresponding source and documentation."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import argparse

ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIRS = ('src', 'scripts', 'docs', 'examples', 'licenses', '.github')
SOURCE_FILES = ('README.md', 'LICENSE', 'THIRD_PARTY_NOTICES.md', 'CONTRIBUTING.md',
                'CHANGELOG.md', 'requirements.txt', 'requirements-dev.txt',
                'WordToPDF.spec', '.gitignore', '.gitattributes')

def source_files():
    for name in SOURCE_FILES:
        path = ROOT / name
        if not path.is_file():
            raise FileNotFoundError(path)
        yield path
    for name in SOURCE_DIRS:
        for path in sorted((ROOT / name).rglob('*')):
            if path.is_file() and '__pycache__' not in path.parts and path.suffix not in ('.pyc', '.pyo'):
                yield path

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--exe', type=Path, help='Built executable; omit only with --source-only.')
    parser.add_argument('--source-only', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if not args.source_only and (args.exe is None or not args.exe.is_file()):
        parser.error('--exe must point to an existing executable.')
    output = args.output.resolve();output.parent.mkdir(parents=True, exist_ok=True)
    files = list(source_files())
    with ZipFile(output, 'w', ZIP_DEFLATED, compresslevel=6) as archive:
        for path in files:
            archive.write(path, Path('word-to-pdf') / path.relative_to(ROOT))
        if not args.source_only:
            archive.write(args.exe, 'word-to-pdf/Word转PDF.exe')
    print(output)

if __name__ == '__main__':
    main()
