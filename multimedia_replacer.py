from pathlib import Path
import os
import re
import shutil
import subprocess
import sys
import csv
from emu_xml_parser import record
from PIL import Image
import multimedia_funcs
import argparse

Image.MAX_IMAGE_PIXELS = 933120000
MIN_SIZE = 2048

def replace_jpegs(input_xml, out_dir, min_size=MIN_SIZE):
    for r in record.parse_xml(input_xml):
        ident = r['EADUnitID']
        row = dict(EADUnitID=r['EADUnitID'], EADUnitTitle=r['EADUnitTitle'])
        for page, multimedia in enumerate(r.get("MulMultiMediaRef_tab")):
            row['page'] = page
            row['MulMultiMediaRef_tab.irn'] = multimedia.get("irn")
            try:
                size = (
                    int(multimedia.get("ChaImageWidth")),
                    int(multimedia.get("ChaImageHeight"))
                        )
                row['size'] = "{}x{}".format(*size)
            except TypeError as e:
                print(e, ident)
                size = (0, 0)
                row['size'] = "Not in EMu"
            if max(size) < min_size:
                print(ident, "page", page, "is insufficient size:", size)
                folder = multimedia_funcs.find_asset_folder(ident)
                row['TIF_folder'] = folder
                tifs = multimedia_funcs.find_assets(folder)
                try:
                    tif = list(tifs)[page]
                    row['TIF'] = tif
                    with Image.open(tif) as im:
                        row['TIF_size'] = "{}x{}".format(*im.size)
                        tif_size = im.size
                    if max(size) < max(tif_size):
                        print(im.size)
                        jpeg = multimedia_funcs.create_jpeg(tif, out_dir, dim=f"{min_size}x{min_size}^")
                        row['status'] = "JPEG replaced"
                        row['Multimedia'] = jpeg
                    else:
                         row['status'] = "TIF is insufficient quality"
                except IndexError:
                    yield(dict(EADUnitID=ident, page=page, status="Poor quality, no Tif found"))
            else:
                row['status'] = 'OK'
            yield row


def main(input_xml, out_dir, min_size=MIN_SIZE):
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
        for row in replace_jpegs(input_xml, out_dir, min_size=min_size):
            log_writer.writerow(row)
            if row.get('Multimedia') is not None:
                jpeg_row = dict(irn=row.get('MulMultiMediaRef_tab.irn'), Multimedia=row.get('Multimedia'))
                multi_writer.writerow(jpeg_row)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Replace multimedia images with better quality versions if found')
    parser.add_argument(
        'input', metavar='i', help='EMu catalogue xml')
    parser.add_argument(
        'output', help='directory for multimedia assets and output sheets')
    parser.add_argument(
        '--minimum', '-m', type=int, default=MIN_SIZE, help='minimum dimension')
    
    args = parser.parse_args()
    main(args.input, args.output, min_size=args.minimum)