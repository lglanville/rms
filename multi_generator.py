import sys
import csv
from pathlib import Path
import argparse
from emu_xml_parser import record
import multimedia_funcs
MIN_SIZE = 2048


def replace_jpegs(input_xml, out_dir, to_jpeg=False, min_size=MIN_SIZE):
    for r in record.parse_xml(input_xml):
        ident = r['EADUnitID']
        row = dict(EADUnitID=r['EADUnitID'])
        folder = multimedia_funcs.find_asset_folder(ident)
        tifs = list(multimedia_funcs.find_assets(folder))
        pdf = list(multimedia_funcs.find_assets(folder, exts=('.pdf')))
        multi = r.get("MulMultiMediaRef_tab")
        if multi is None:
            multi = []
            format = None
        else:
            format = multi[0].get('MulMimeFormat')
        if len(tifs) > len(multi) and format != 'pdf':
            print("Found additional assets for", ident)
            if pdf != []:
                row['MulMultiMediaRef_tab(+).Multimedia'] = pdf[0]
                row['MulMultiMediaRef_tab(+).DetSource'] = ident
                row['MulMultiMediaRef_tab(+).MulTitle'] = r['EADUnitTitle']
                yield row
            else:
                if to_jpeg:
                    start = 0
                else:
                    start = len(multi)
                for tif in tifs[start:]:
                    jpeg = multimedia_funcs.create_jpeg(tif, out_dir, dim=f'[min_size}x{min_size}^')
                    row['MulMultiMediaRef_tab(+).Multimedia'] = jpeg
                    row['MulMultiMediaRef_tab(+).DetSource'] = ident
                    row['MulMultiMediaRef_tab(+).MulTitle'] = r['EADUnitTitle']
                    yield row


def main(input_xml, out_dir, to_jpeg=False, min_size=MIN_SIZE):
    csv_file = Path(out_dir, "jpeg_replacements.csv")
    with csv_file.open('w', encoding='utf-8-sig', newline='') as f:
        fieldnames = [
            'EADUnitID', 'AdmPublishWebNoPassword', 'SecDepartment_tab(2)', 
            'MulMultiMediaRef_tab(+).Multimedia', 'MulMultiMediaRef_tab(+).DetSource', 
            'MulMultiMediaRef_tab(+).MulTitle',
            'MulMultiMediaRef_tab(+).AdmPublishWebNoPassword']
        multi_writer = csv.DictWriter(f, fieldnames=fieldnames)
        multi_writer.writeheader()
        for row in replace_jpegs(input_xml, out_dir, to_jpeg=to_jpeg, min_size=min_size):
            multi_writer.writerow(row)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Replace multimedia images with better quality versions if found')
    parser.add_argument(
        'input', metavar='i', help='EMu catalogue xml')
    parser.add_argument(
        'output', help='directory for multimedia assets and output sheets')
    parser.add_argument(--to_jpeg, '-j', action="store_true", help='if a pdf asset, create new jpegs for all assets found')
    parser.add_argument(
        '--minimum', '-m', type=int, default=MIN_SIZE, help='minimum dimension')
    
    args = parser.parse_args()
    main(args.input, args.output, to_jpeg=args.to_jpeg, min_size=args.minimum)