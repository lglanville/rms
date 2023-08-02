from pathlib import Path
import argparse
import sys
import openpyxl
import multimedia_funcs

def main(workbookpath, min=None):
    dim = '2048x2048^>'
    if min is not None:
        dim = f'{min}x{min}^>'
    wb = openpyxl.open(workbookpath)
    workbookpath = Path(workbookpath)
    asset_dir = workbookpath.parent / workbookpath.stem
    asset_dir.mkdir(exist_ok=True)
    for ws in wb.worksheets:
        asset_col = None
        for cell in ws[1]:
            if cell.value == 'ASSETS':
                asset_col = cell.column - 1
            elif cell.value == 'Previous System ID':
                id_col = cell.column - 1
        for row in ws.iter_rows(min_row=2):
            assets = row[asset_col].value
            if assets is not None:
                jpegs = assets.split('|')
                ident = row[id_col].value
                print(ident, jpegs)
                asset_fol = multimedia_funcs.find_asset_folder(ident)
                tifs = list(multimedia_funcs.find_assets(asset_fol))
                if len(tifs) != len(jpegs):
                    print(f'Warning: {ident} has {len(jpegs)} in EMu and {len(tifs)} in storage')
                if bool(tifs):
                    j = []
                    for tif in tifs:
                        jpeg = multimedia_funcs.create_jpeg(tif, asset_dir, dim=dim)
                        j.append(jpeg.relative_to(asset_dir.parent).as_posix())
                    row[asset_col].value = '|'.join(j)
    p = Path(workbookpath)
    new_path = p.parent / (p.stem+'_new_assets.xlsx')
    print("Saving workbook to", new_path)
    wb.save(new_path)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu items into ReCollect data')
    parser.add_argument(
        'workbook', metavar='i', help='recollect workbook')
    parser.add_argument(
        '--min', '-m',
        help='minimum dimensions (in pixels) of jpegs')
    args = parser.parse_args()
    main(args.workbook, min=args.min)