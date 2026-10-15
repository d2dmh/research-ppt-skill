"""Render a PPTX to per-slide PNGs plus a contact sheet.

Engine "auto" uses Microsoft PowerPoint on Windows when available (closest to what the
audience sees) and otherwise LibreOffice with assets/fonts/fonts.conf mapping
微软雅黑 → Noto Sans CJK / 思源黑体, Arial → Liberation Sans. LibreOffice line breaks and
glyph widths differ slightly from PowerPoint; confirm final typography in PowerPoint.

    python scripts/render.py --input deck.pptx --directory work/render [--slides 3 5] [--engine auto]
"""
import argparse
import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def powerpoint(source, destination, indices, width):
    import win32com.client
    app = win32com.client.DispatchEx('PowerPoint.Application')
    deck = None
    try:
        deck = app.Presentations.Open(str(source), ReadOnly=True, Untitled=False, WithWindow=False)
        count = deck.Slides.Count
        indices = indices or list(range(1, count + 1))
        for i in indices:
            deck.Slides(i).Export(str(destination / f'slide-{i:02d}.png'), 'PNG', width, round(width * 9 / 16))
    finally:
        if deck is not None:
            deck.Close()
        app.Quit()
    return indices


def libreoffice(source, destination, indices, width):
    office = shutil.which('soffice') or shutil.which('libreoffice')
    if not office:
        for candidate in ('/Applications/LibreOffice.app/Contents/MacOS/soffice',
                          r'C:\Program Files\LibreOffice\program\soffice.exe'):
            if Path(candidate).exists():
                office = candidate
    if not office:
        raise SystemExit('No renderer: install Microsoft PowerPoint (Windows) or LibreOffice.')
    env = dict(os.environ)
    if platform.system() != 'Windows':
        env['FONTCONFIG_FILE'] = str(ROOT / 'assets' / 'fonts' / 'fonts.conf')
    with tempfile.TemporaryDirectory() as work:
        work = Path(work)
        subprocess.run([office, '--headless', f'-env:UserInstallation=file://{work / "profile"}',
                        '--convert-to', 'pdf', '--outdir', str(work), str(source)],
                       check=True, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
        pdf = work / (source.stem + '.pdf')
        dpi = round(width / 13.333)
        if shutil.which('pdftoppm'):
            subprocess.run(['pdftoppm', '-r', str(dpi), '-png', str(pdf), str(work / 'p')], check=True)
            pages = sorted(work.glob('p-*.png'), key=lambda p: int(p.stem.split('-')[-1]))
        else:
            import fitz  # PyMuPDF fallback
            doc = fitz.open(pdf)
            pages = []
            for n, page in enumerate(doc, 1):
                out = work / f'p-{n}.png'
                page.get_pixmap(dpi=dpi).save(out)
                pages.append(out)
        indices = indices or list(range(1, len(pages) + 1))
        for i in indices:
            shutil.copy(pages[i - 1], destination / f'slide-{i:02d}.png')
        shutil.copy(pdf, destination / (source.stem + '.pdf'))
    return indices


def contact_sheet(destination, indices):
    from PIL import Image, ImageDraw
    cols = 3 if len(indices) > 4 else 2
    tile_w, tile_h, pad = 520, 293, 14
    rows = (len(indices) + cols - 1) // cols
    sheet = Image.new('RGB', (cols * (tile_w + pad) + pad, rows * (tile_h + pad + 18) + pad), '#d7dbe0')
    draw = ImageDraw.Draw(sheet)
    for n, number in enumerate(indices):
        x, y = pad + (n % cols) * (tile_w + pad), pad + (n // cols) * (tile_h + pad + 18)
        draw.text((x, y), f'{number}', fill='#111111')
        with Image.open(destination / f'slide-{number:02d}.png') as frame:
            sheet.paste(frame.convert('RGB').resize((tile_w, tile_h)), (x, y + 16))
    sheet.save(destination / 'contact-sheet.png')


def render(source, destination, indices=None, engine='auto', width=1600):
    source, destination = Path(source).resolve(), Path(destination).resolve()
    destination.mkdir(parents=True, exist_ok=True)
    if engine == 'auto':
        engine = 'libreoffice'
        if platform.system() == 'Windows':
            try:
                import win32com.client  # noqa: F401
                engine = 'powerpoint'
            except ImportError:
                pass
    indices = (powerpoint if engine == 'powerpoint' else libreoffice)(source, destination, indices, width)
    contact_sheet(destination, indices)
    print(f'Rendered {len(indices)} slides with {engine} -> {destination}')
    if engine == 'libreoffice':
        print('Note: LibreOffice preview; fonts are substituted, confirm final line breaks in PowerPoint.')
    return engine


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, required=True)
    parser.add_argument('--directory', type=Path, required=True)
    parser.add_argument('--slides', type=int, nargs='+')
    parser.add_argument('--engine', choices=['auto', 'powerpoint', 'libreoffice'], default='auto')
    parser.add_argument('--width', type=int, default=1600)
    args = parser.parse_args()
    render(args.input, args.directory, args.slides, args.engine, args.width)
