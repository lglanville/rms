import csv
import re
import argparse
import re
import shutil
from pathlib import Path
import metadata_funcs
import multimedia_funcs
import openpyxl

MAP = {
        'EADUnitTitle': 'NODE_TITLE',
        'EADUnitID': 'Previous System ID',
        'EADScopeAndContent': 'Scope and Content',
        'EADDimensions': 'Dimensions',
        'NotNotes': 'Internal Notes',
        'EADAccessRestrictions': 'Access Conditions',
        'EADUseRestrictions': 'Conditions of Use and Reproduction',
        'AdmPublishWebNoPassword': 'Publication Status',
        'EADGenreForm_tab': 'Genre/Form',
        'MulMultiMediaRef_tab.AdmPublishWebNoPassword': '#REDACT',
        'AssParentObjectRef.EADUnitID': 'Accession',
        'LocCurrentLocationRef.LocHolderName': 'Unit',
        'LocCurrentLocationRef.LocLocationCode': 'Location if Unenclosed',
        'EADSubject_tab': 'Subject',
        'EADGeographicName_tab': 'Subject (Place)',
        'EADExtent_tab': 'Extent',
        'EADPreviousID_tab': 'Other IDs',
        '###Dates': '###Dates',
        'm_creator': '###Provenance',
        'Instructions': 'Digitisation Notes',
        'Accession': 'Accession',
        'Series': 'Series',
        'Part of Item': 'Part of Item'
        }

def re_jpeg(ident, asset_dir, min=2048):
    asset_fol = multimedia_funcs.find_asset_folder(ident)
    tifs = list(multimedia_funcs.find_assets(asset_fol))
    if bool(tifs):
        for tif in tifs:
            jpeg = multimedia_funcs.create_jpeg(tif, asset_dir)
            yield jpeg.relative_to(asset_dir.parent).as_posix()

class Assets():
    def __init__(self, csv_sheet):
        self.asset_map = {}
        with open(csv_sheet, encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                id = row.get('EADUnitID')
                asset = Path(row.get('MulMultiMediaRef_tab(+).Multimedia'))
                assert asset.exists()
                if id in self.asset_map.keys():
                    self.asset_map[id].append(asset)
                else:
                    self.asset_map[id] = [asset]

    def get_assets(self, id, asset_dir):
        assets = self.asset_map.get(id)
        if assets is not None:
            for asset in assets:
                target = Path(asset_dir, asset.name)
                shutil.copy2(asset, target)
                yield target.relative_to(asset_dir.parent).as_posix()

class parent_map():
    def __init__(self):
        self.parent_map = []

    def update(self, ident):
        i = None
        choices = ('1', '2', '3')
        while i not in choices:
            i = input(f'Is {ident} an (1) accession, (2) series, or (3) item? ->')
        if i == '1':
            type = 'Accession'
        elif i == '2':
            type = 'Series'
        elif i == '3':
            type = 'Part of Item'
        title = input(f'Enter the title of {ident} ->')
        self.parent_map.append({'type': type, 'id': ident, 'title': title})

    def retrieve_parent(self, ident):
        p = [x for x in self.parent_map if x['id'] == ident]
        if bool(p):
            return p[0]['title'], p[0]['type']
        else:
            self.update(ident)
            return self.retrieve_parent(ident)

class csv_row(dict):
    def restructure(self):
        print(self)
        new_vals = csv_row()
        for k, v in self.items():
            if '(' in k:
                fieldname = re.sub('\(.+\)', '', k)
                if fieldname in new_vals.keys():
                    new_vals[fieldname].append(v)
                else:
                    new_vals[fieldname] = [v]
            else:
                new_vals[k] = v
        print(new_vals)
        return new_vals

    def convert_dates(self):
        string_date = self.pop('EADUnitDate')
        e_date = ''
        if 'EADUnitDateEarliest' in self.keys():
            e_date = self.pop('EADUnitDateEarliest')
        l_date = ''
        if 'EADUnitDateLatest' in self.keys():
            l_date = self.pop('EADUnitDateLatest')
        self['###Dates'] = metadata_funcs.format_date(string_date, e_date, l_date)
   
    
    def convert_holder_name(self):
        if 'LocCurrentLocationRef.LocHolderName' in self.keys():
            holder_name = self.pop('LocCurrentLocationRef.LocHolderName')
            if holder_name is not None:
                print(holder_name)
                repl = lambda x: x.group(1) +" " + "0"*(4-len(x.group(2)))+x.group(2)
                new_name = re.sub(r'(unit|album) (\d+)', repl, holder_name, flags=re.IGNORECASE)
                self['LocCurrentLocationRef.LocHolderName'] = new_name

    def get_parent(self, pm):
        parent_id = self.pop('AssParentObjectRef.EADUnitID')
        title, type = pm.retrieve_parent(parent_id)
        self[type] = title


    def convert_to_row(self, pm, access, request_type, provenance=None,accession=None, assets=None):
        self = self.restructure()
        self.convert_dates()
        self.convert_holder_name()
        self.get_parent(pm)
        self['Access Status'] = access
        self['Request Type'] = request_type
        if self.get('Accession') is None:
            self['Accession'] = accession
        if self.get('###Provenance') is None:
            self['###Provenance'] = provenance
        new_row = {}
        for k, v in self.items():
            if k in MAP.keys():
                new_row[MAP[k]] = v
            else:
                print(k,':', v, 'not mappable')
        new_row['Identifier'] = "UMA-ITE-" + new_row.get('Previous System ID').replace('.', '')
        if assets is not None:
            new_row['ASSETS'] = assets.get_assets(new_row.get('Previous System ID'))
        return new_row

def from_udc(recollect_sheet, multimedia_sheet, out_dir, regen_jpegs=False):
    out_dir = Path(out_dir)
    multimedia_sheet = Path(multimedia_sheet)
    asset_dir = Path(out_dir, multimedia_sheet.stem)
    asset_dir.mkdir(exist_ok=True)
    asset_map = Assets(multimedia_sheet)
    wb = openpyxl.Workbook()
    sheet = wb.active
    fieldnames = ['Node ID', 'Node Type', 'ASSETS', 'ATTACHMENTS']
    sheet.append(fieldnames)
    with open(recollect_sheet, encoding='utf-8') as f:
        reader = csv.DictReader(f)
        templates = metadata_funcs.template_handler()
        for row in reader:
            unit_id = row['Previous System ID']
            if regen_jpegs:
                assets = list(re_jpeg(unit_id, asset_dir))
            else:
                assets = list(asset_map.get_assets(unit_id, asset_dir))
            if bool(assets):
                sheet.append([row.get("Node ID"), row.get("Node Type"), '|'.join(assets)])
    fpath = Path(out_dir, f'{Path(multimedia_sheet).stem}.xlsx')
    wb.save(fpath)

def main(meta_sheet, out_dir, access=None, request=None, provenance=None,accession=None, multimedia_sheet=None, regen_jpegs=False):
    fields = check_headings(meta_sheet)
    if 'MulMultiMediaRef_tab(+).Multimedia' in fields:
        from_udc(meta_sheet, multimedia_sheet, out_dir, regen_jpegs=regen_jpegs)
    else:
        flip_metadata(meta_sheet, out_dir, access, request, provenance=None,accession=None)


def check_headings(csv_file):
    with open(csv_file, encoding='utf-8') as f:
        reader = csv.DictReader(f)
        return reader.fieldnames


def flip_metadata(spreadsheet, out_dir, access, request_type, template_name='item', provenance=None,accession=None, assets=None):
    if assets:
        assets = Assets(assets)
    pm = parent_map()
    with open(spreadsheet, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        templates = metadata_funcs.template_handler()
        for row in reader:
            row = csv_row(row)
            rc_row = row.convert_to_row(pm, access, request_type, assets=assets)
            templates.add_row(template_name, rc_row)
        templates.serialise(out_dir, sort_by='Identifier')

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Get EMu metadata into ReCollect data')
    parser.add_argument(
        'input_csv', metavar='i', help='EMu or Recollect csv')
    parser.add_argument(
        'output', help='directory for multimedia assets and output sheets')
    parser.add_argument(
        '--template', '-t', default='item',
        help='template for output')
    parser.add_argument(
        '--assets', '-a',
        help='Asset csv sheet')
    parser.add_argument(
        '--regen_jpegs', '-j', action='store_true',
        help='Regenerate jpegs from stored tifs')
    parser.add_argument(
        '--provenance', '-p',
        help='A provenance entity to assign to the whole sheet')
    parser.add_argument(
        '--accession',
        help='An accession to assign to the whole sheet')
    parser.add_argument(
        '--access', '-s', default="Open for public access",
        help='An access status')
    parser.add_argument(
        '--request', '-r', default="Request unit",
        help='a request type')

    args = parser.parse_args()
    main(
        args.input_csv, args.output, 
        multimedia_sheet=args.assets, 
        provenance=args.provenance,
        accession=args.accession,
        access=args.access,
        request=args.request,
        regen_jpegs=args.regen_jpegs)
