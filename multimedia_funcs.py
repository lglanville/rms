from pathlib import Path
import os
import re
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor

def id_transform(number):
    idents = re.findall(r'(\d{4})[,_.-](\d{2,4})[,_.-](\d{1,5})', number)
    if len(idents) == 1:
        pre = idents[0][0]
        mid = '0' * (4 - len(idents[0][1])) + idents[0][1]
        suf = '0' * (5 - len(idents[0][2])) + idents[0][2]
        identifier = pre + '.' + mid + '.' + suf
        return identifier


def move_image(fpath, base_dir, ident):
    fpath = Path(fpath)
    newpath = Path(base_dir, *ident.split('.'), fpath.suffix.upper().strip('.'))
    if not newpath.exists():
        newpath.mkdir(parents=True)
    page = 1
    new_fname = '-'.join(ident.split('.'))+'-'+'0'*(5-len(str(page)))+str(page)+fpath.suffix
    new_dest = newpath / new_fname
    while new_dest.exists():
        page += 1
        new_fname = '-'.join(ident.split('.'))+'-'+'0'*(5-len(str(page)))+str(page)+fpath.suffix
        new_dest = newpath / new_fname
    shutil.copy2(fpath, new_dest)
    print(fpath.name, '->', new_fname)
    return new_dest


def create_jpeg(fpath, outdir, dim='2048x2048^>'):
    fpath = Path(fpath)
    outfile = Path(outdir, fpath.stem+'.jpg')
    if not outfile.exists():
        print("jpegging", fpath.name, "->", outfile.name)
        subprocess.run(
            [
                'magick', 'convert', str(fpath), '-resize', dim,
                '-quality', '65', '-depth', '8', '-unsharp',
                '1.5x1+0.7+0.02', str(outfile)])
    return outfile


def bulk_jpeg(in_dir, out_dir, lformat=False, exts=['.tif', '.tiff']):
    f = []
    with ThreadPoolExecutor() as ex:
        for root, _, files in os.walk(in_dir):
            for file in files:
                fpath = Path(root, file)
                if fpath.suffix.lower() in exts:
                    print('jpegging', fpath)
                    f.append(ex.submit(jpeg, *(fpath, out_dir, lformat)))
    results = [future.result() for future in f]
    return results


def find_asset_folder(ident):
    base = r"\\research-cifs.unimelb.edu.au\9730-UniversityArchive-Shared\Digitised_Holdings\Registered"
    try:
        pre, mid, suf = ident.split('.')
        folder = Path(base, pre, mid, suf)
        if folder.exists():
            return folder
        else:
            folder = Path(base, pre, mid, suf[1:])
            if folder.exists():
                return folder
    except ValueError as e:
        print(e)


def find_assets(folder, exts=('.tif', '.tiff')):
    if folder is None:
        return None
    for root, _, files in os.walk(folder):
        for file in files:
            fpath = Path(root, file)
            if fpath.suffix.lower() in exts:
                yield fpath