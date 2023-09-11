from pathlib import Path
import argparse
import sys
import shutil
import openpyxl
import multimedia_funcs

def copy_asset(asset):
    p = Path(workbookpath.parent / a)
    if p.exists():
        shutil.copy2(p, asset_dir)
        att.append(p.relative_to(asset_dir.parent).as_posix())

def replace_assets(assets, asset_dir, min):
    jpegs = assets.split('|')
    ident = row[id_col].value
    asset_fol = multimedia_funcs.find_asset_folder(ident)
    tifs = list(multimedia_funcs.find_assets(asset_fol))
    if len(tifs) != len(jpegs):
        print(f'Warning: {ident} has {len(jpegs)} in EMu and {len(tifs)} in storage')
    j = []
    if bool(tifs):
        for tif in tifs:
            jpeg = multimedia_funcs.create_jpeg(tif, asset_dir, dim=min)
            j.append(jpeg.relative_to(asset_dir.parent).as_posix())
    else:
        for jpeg in jpegs:
            p = Path(workbookpath.parent / jpeg)
            if p.exists():
                shutil.copy2(p, asset_dir)
                j.append(p.relative_to(asset_dir.parent).as_posix())
            else:
                print("Warning:", jpeg, "not found")
    return j

def main(workbookpath, min=None):
    dim = '2048x2048^>'
    if min is not None:
        dim = f'{min}x{min}^>'
    wb = openpyxl.open(workbookpath)
    workbookpath = Path(workbookpath)
    asset_dir = workbookpath.parent / (workbookpath.stem + '_new_assets')
    asset_dir.mkdir(exist_ok=True)
    for ws in wb.worksheets:
        asset_col = None
        for cell in ws[1]:
            if cell.value == 'ASSETS':
                asset_col = cell.column - 1
            if cell.value == 'ATTACHMENTS':
                attach_col = cell.column - 1
            elif cell.value == 'Previous System ID':
                id_col = cell.column - 1
        for row in ws.iter_rows(min_row=2):
            assets = row[asset_col].value
            attachments = row[attach_col].value
            if bool(attachments):
                att = []
                attachments = attachments.split('|')
                for a in attachments:
                    p = Path(workbookpath.parent / a)
                    if p.exists():
                        shutil.copy2(p, asset_dir)
                        att.append(p.relative_to(asset_dir.parent).as_posix())
                row[attach_col].value = '|'.join(att)
            if assets is not None:
                jpegs = replace_assets(assets, asset_dir, min=min)
                row[asset_col].value = '|'.join(jpegs)
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