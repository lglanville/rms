import sys
import csv
import multimedia_funcs
BASE_DIR = r"\\research-cifs.unimelb.edu.au\9730-UniversityArchive-Shared\Digitised_Holdings\Registered"

def move_and_jpeg(csv_file, out_dir):
    with open(csv_file, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        yield reader.fieldnames
        for row in reader:
            tif = multimedia_funcs.move_image(row['fpath'], BASE_DIR, row['EADUnitID'])
            jpeg = create_jpeg(tif, out_dir)
            row['MulMultiMediaRef_tab(+).Multimedia'] = jpeg
            yield row

def main(in_file, out_file, out_dir):
    with open(out_file, 'w', encoding='utf-8-sig', newline='') as f:
        fieldnames = ["EADUnitID", "EADUnitTitle", "previous_id", "fpath", "MulMultiMediaRef_tab(+).Multimedia"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in move_and_jpeg(in_file, out_dir):
            writer.writerow(row)

if __name__ == '__main__':
    main(*sys.argv[1:])