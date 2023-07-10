import sys
import csv
from pathlib import Path
import argparse
from emu_xml_parser import record
import multimedia_funcs
MIN_SIZE = 2048


class multi_replacer(record):

    def find_assets(self):
        if self.get('EADUnitID') is not None:
            self.asset_folder = multimedia_funcs.find_asset_folder(self['EADUnitID'])
            self.tifs = list(multimedia_funcs.find_assets(self.asset_folder))
            self.pdf = list(multimedia_funcs.find_assets(self.asset_folder, exts=('.pdf')))
        else:
            raise ValueError('Record has no EADUnitID')
    
    def to_jpeg(self, out_dir, min_size=MIN_SIZE):
        self.find_assets()
        for tif in self.tifs:
            row = {}
            for k, v in self.items():
                if type(v) == str:
                    row[k] = v
            jpeg = multimedia_funcs.create_jpeg(tif, out_dir, dim=f'{min_size}x{min_size}^')
            row['MulMultiMediaRef_tab(+).Multimedia'] = jpeg
            row['MulMultiMediaRef_tab(+).DetSource'] = self['EADUnitID']
            row['MulMultiMediaRef_tab(+).MulTitle'] = self['EADUnitTitle']
            yield row
    

    def find_new_jpegs(self, out_dir, min_size=MIN_SIZE):
        self.find_assets()
        multi = self.get("MulMultiMediaRef_tab")
        if multi is None:
            multi = []
            format = None
        else:
            format = multi[0].get('MulMimeFormat')
        if len(self.tifs) > len(multi) and format != 'pdf':
            print("Found additional assets for", ident)
            row = {}
            if self.pdf != []:
                row['MulMultiMediaRef_tab(+).Multimedia'] = pdf[0]
                row['MulMultiMediaRef_tab(+).DetSource'] = ident
                row['MulMultiMediaRef_tab(+).MulTitle'] = r['EADUnitTitle']
                yield row
            else:
                for tif in tifs[len(multi):]:
                    jpeg = multimedia_funcs.create_jpeg(tif, out_dir, dim=f'{min_size}x{min_size}^')
                    row['MulMultiMediaRef_tab(+).Multimedia'] = jpeg
                    row['MulMultiMediaRef_tab(+).DetSource'] = self['EADUnitID']
                    row['MulMultiMediaRef_tab(+).MulTitle'] = r['EADUnitTitle']
                    yield row


def main(input_xml, out_dir, to_jpeg=False, min_size=MIN_SIZE):
    fieldnames = set()
    rows = []
    for x in multi_replacer.parse_xml(input_xml):
            if to_jpeg:
                for row in x.to_jpeg(out_dir, min_size=min_size):
                    fieldnames.update(row.keys())
                    rows.append(row)
            else:
                for row in x.find_new_jpegs(out_dir, min_size=min_size):
                    fieldnames.update(row.keys())
                    rows.append(row)
    csv_file = Path(out_dir, "jpeg_replacements.csv")
    with csv_file.open('w', encoding='utf-8-sig', newline='') as f:
        multi_writer = csv.DictWriter(f, fieldnames=list(fieldnames))
        multi_writer.writeheader()
        multi_writer.writerows(rows)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Replace multimedia images with better quality versions if found')
    parser.add_argument(
        'input', metavar='i', help='EMu catalogue xml')
    parser.add_argument(
        'output', help='directory for multimedia assets and output sheets')
    parser.add_argument('--to_jpeg', '-j', action="store_true", help='if a pdf asset, create new jpegs for all assets found')
    parser.add_argument(
        '--minimum', '-m', type=int, default=MIN_SIZE, help='minimum dimension')
    
    args = parser.parse_args()
    main(args.input, args.output, to_jpeg=args.to_jpeg, min_size=args.minimum)