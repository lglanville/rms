from pathlib import Path
import os
import re
import shutil
import subprocess
import sys
import csv
from emu_xml_parser import record
from PIL import Image

BASE_DIR = "\\research-cifs.unimelb.edu.au\9730-UniversityArchive-Shared\Registered"

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


def find_asset(folder, index):
    if folder is None:
        return None
    tifs = []
    for root, _, files in os.walk(folder):
        for file in files:
            fpath = Path(root, file)
            if fpath.suffix.lower() in ('.tif', '.tiff'):
                tifs.append(fpath)
    try:
        return tifs[index]
    except IndexError as e:
        print(e)


def make_assets(indir, outdir):
    tifs = []
    for root, _, files in os.walk(indir):
        for file in files:
            fpath = Path(root, file)
            if fpath.suffix.lower() in ('.tif', '.tiff', '.dng'):
                jpeg = create_jpeg(fpath, outdir)
                yield jpeg

def replace_jpegs(input_xml, out_dir):
    for r in record.parse_xml(input_xml):
        ident = r['EADUnitID']
        row = dict(EADUnitID=r['EADUnitID'], EADUnitTitle=r['EADUnitTitle'])
        for page, multimedia in enumerate(r.get("MulMultiMediaRef_tab")):
            row['page'] = page
            row['MulMultiMediaRef_tab.irn'] = multimedia.get("irn")
            size = (
                int(multimedia.get("ChaImageWidth")),
                int(multimedia.get("ChaImageHeight"))
                    )
            row['size'] = "{}x{}".format(*size)
            if max(size) < 2000:
                print(ident, "page", page, "is insufficient size:", size)
                folder = find_asset_folder(ident)
                row['TIF_folder'] = folder
                tif = find_asset(folder, page)
                row['TIF'] = tif
                if tif is not None:
                    with Image.open(tif) as im:
                        row['TIF_size'] = "{}x{}".format(*size)
                        if size != im.size:
                            print(im.size)
                            jpeg = create_jpeg(tif, out_dir)
                            row['status'] = "JPEG replaced"
                            row['Multimedia'] = jpeg
                        else:
                            row['status'] = "TIF is insufficient quality"
                else:
                    yield(dict(EADUnitID=ident, page=page, status="Poor quality, no Tif found"))
            else:
                row['status'] = 'OK'
            yield row

def main(input_xml, out_dir):
    csv_file = Path(out_dir, "jpeg_replacements.csv")
    csv_log = Path(out_dir, "multimedia_log.csv")
    with csv_file.open('w', encoding='utf-8-sig', newline='') as f, csv_log.open('w', encoding='utf-8-sig', newline='') as g:
        fieldnames = [
            'EADUnitID', 'EADUnitTitle', 'status', 'page', 'MulMultiMediaRef_tab.irn',
            'size', 'TIF_folder', 'TIF', 'TIF_size', 'status', 'Multimedia']
        multi_writer = csv.DictWriter(f, fieldnames=['irn', 'Multimedia'])
        log_writer = csv.DictWriter(g, fieldnames=fieldnames)
        multi_writer.writeheader()
        log_writer.writeheader()
        for row in replace_jpegs(input_xml, out_dir):
            log_writer.writerow(row)
            if row.get('Multimedia') is not None:
                jpeg_row = dict(irn=row.get('MulMultiMediaRef_tab.irn'), Multimedia=row.get('MulMultiMediaRef_tab.irn'))
                multi_writer.writerow(jpeg_row)

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
