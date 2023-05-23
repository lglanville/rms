import sys
import csv
from pathlib import Path
from emu_xml_parser import record
import multimedia_funcs

def replace_jpegs(input_xml, out_dir):
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
                for tif in tifs[len(multi):]:
                    jpeg = multimedia_funcs.create_jpeg(tif, out_dir, dim='4096x4096^>')
                    row['MulMultiMediaRef_tab(+).Multimedia'] = jpeg
                    row['MulMultiMediaRef_tab(+).DetSource'] = ident
                    row['MulMultiMediaRef_tab(+).MulTitle'] = r['EADUnitTitle']
                    yield row


def main(input_xml, out_dir):
    csv_file = Path(out_dir, "jpeg_replacements.csv")
    with csv_file.open('w', encoding='utf-8-sig', newline='') as f:
        fieldnames = [
            'EADUnitID', 'AdmPublishWebNoPassword', 'SecDepartment_tab(2)', 
            'MulMultiMediaRef_tab(+).Multimedia', 'MulMultiMediaRef_tab(+).DetSource', 
            'MulMultiMediaRef_tab(+).MulTitle',
            'MulMultiMediaRef_tab(+).AdmPublishWebNoPassword']
        multi_writer = csv.DictWriter(f, fieldnames=fieldnames)
        multi_writer.writeheader()
        for row in replace_jpegs(input_xml, out_dir):
            multi_writer.writerow(row)

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])