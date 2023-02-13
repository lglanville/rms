from pathlib import Path
import os
import re
import shutil
import subprocess

def create_jpeg(fpath, outdir):
    fpath = Path(fpath)
    outfile = Path(outdir, fpath.stem+'.jpg')
    if not outfile.exists():
        print("jpegging", fpath.name, "->", outfile.name)
        subprocess.run(
            [
                'magick', 'convert', str(fpath), '-resize', '2048x2048^>',
                '-quality', '65', '-depth', '8', '-unsharp',
                '1.5x1+0.7+0.02', str(outfile)])
    return outfile

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

def find_assets(folder):
    if folder is None:
        return None
    for root, _, files in os.walk(folder):
        for file in files:
            fpath = Path(root, file)
            if fpath.suffix.lower() in ('.tif', '.tiff'):
                yield fpath