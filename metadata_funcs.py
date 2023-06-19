import csv
from pprint import pprint
from pathlib import Path
import shutil
import re
from uuid import uuid4
import unicodedata
import multimedia_funcs
import openpyxl
from pypdf import PdfWriter, PdfReader

TEMPLATE_DIR = Path(__file__).parent / "Templates"

def slugify(value, allow_unicode=False):
    """
    Taken from https://github.com/django/django/blob/master/django/utils/text.py
    Convert to ASCII if 'allow_unicode' is False. Convert spaces or repeated
    dashes to single dashes. Remove characters that aren't alphanumerics,
    underscores, or hyphens. Convert to lowercase. Also strip leading and
    trailing whitespace, dashes, and underscores.
    """
    value = str(value)
    if allow_unicode:
        value = unicodedata.normalize('NFKC', value)
    else:
        value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
    value = re.sub(r'[^\w\s-]', '', value.lower())
    return re.sub(r'[-\s]+', '-', value).strip('-_')

def order_values(row, fieldnames):
    """Order a dictionary to the exact format and column order required by ReCollect.
    Will throw an error if dictionary contains keys not in fieldnames"""
    if not all(item in fieldnames for item in row.keys()):
        pprint(row)
        wrong_keys = filter(None, [key if key not in fieldnames else None for key in row.keys()])
        raise ValueError("Row contains fields not in template", ",".join(wrong_keys))
    ordered_row = Row()
    for f in fieldnames:
        val = row.get(f)
        ordered_row[f] = val
    return ordered_row

def find_access(access_conditions):
    """Determine an access status from text in EADAccessRestrictions field text"""
    access_status = "Access not determined"
    acc_status = {
        'closed': 'Closed for public access',
        'open': 'Open for public access',
        'restricted': 'Access restrictions apply',
        'part restricted': 'Access restrictions apply for some items',
        'part-restricted': 'Access restrictions apply for some items'}
    access_pattern = re.compile(r"(access:? )?(open|closed|restricted|part restricted|part-restricted)", flags=re.IGNORECASE)
    access_status = "Access not determined"
    if access_conditions is not None:
        access_match = access_pattern.search(access_conditions)
        if access_match is not None:
            access_status = acc_status.get(access_match[2].lower())
            access_conditions = access_conditions.replace(access_match[0], '').strip('. ')
    return access_status, access_conditions


def get_access(record):
    """Traverse up the hierarchy to get the nearest parent record's access status"""
    access_status = "Access not determined"
    access_conditions = None
    for x in record.findall('EADAccessRestrictions'):
        access_status, access_conditions = find_access(x)
        if access_status != "Access not determined":
            break
    return access_status, access_conditions


def shorten_title(text):
    """Shorten title to 140 characters"""
    if len(str(text)) > 256:
        for count, x in enumerate(text[140:]):
            if x == ' ':
                break
        return text[:140+count] + '...', text
    else:
        return text, ''

def flatten_table(record, key):
    """Flatten a table field to a list"""
    terms = []
    term_tab = record.get(key)
    if term_tab is not None:
        for x in record[key]:
            t = key.replace('_tab', '')
            terms.append(x[t])
    return list(filter(None, terms))

def format_date(date_str, earliest, latest):
    """Format dates to strcuture required by ReCollect"""
    if date_str is None:
        date_str = ''
    if earliest is None:
        earliest = ''
    if latest is None:
        latest = ''
    if len(earliest) == 4:
        earliest = '01/01/' + earliest
    if len(latest) == 4:
        latest = '31/12/' + latest
    return date_str + ';' + earliest + ";" + latest

def concat_fields(*fields, sep='|'):
    if any(fields):
        fields = list(filter(None, fields))
        if fields != []:
            return sep.join(fields)

def get_multimedia(record):
    """Gat paths for multimedia from temp, or if a pdf from fileshare"""
    asset_data = {'ASSETS': [], '#REDACT': None}
    redacted = None
    multi = record.get('MulMultiMediaRef_tab')
    id = record.get('EADUnitID')
    if multi is not None:
        redact = set([x['AdmPublishWebNoPassword'].lower() for x in multi])
        if len(redact) == 2: 
            print('Warning: multiple publishing permissions for record', id)
        if 'no' in redact:
            asset_data['#REDACT'] = 'Yes'
        for m in multi:
            try:
                fpath = Path(m.get('Multimedia'))
                if fpath.suffix == '.pdf':
                    id = record.get('EADUnitID')
                    folder = multimedia_funcs.find_asset_folder(id)
                    pdfs = list(multimedia_funcs.find_assets(folder, exts=('.pdf')))
                    for pdf in pdfs:
                        print('Found fresh pdf in storage for', id)
                        asset_data['ASSETS'].append(pdf)
                    if pdfs == []:
                        print('No fresh pdfs found, using EMu version for', id)
                        asset_data['ASSETS'].append(fpath)
                else:
                    asset_data['ASSETS'].append(fpath)
            except Exception as e:
                print("Multimedia item not exported from EMu for", id)
    return asset_data       


class Row(dict):
    def copy_assets(self, asset_dir, key):
        if self.get(key) is not None:
            assets = []
            for fpath in self[key]:
                target = Path(asset_dir, fpath.name)
                if not target.exists():
                    shutil.copy2(fpath, target)
                assets.append(target.relative_to(asset_dir.parent).as_posix())
            self[key] = '|'.join(assets)

    def concat_pdfs(self, asset_dir):
        pdfs = list(filter(lambda x: x.suffix.lower() == '.pdf', self['ASSETS']))
        if len(pdfs) > 1:
            output = PdfWriter()
            outfile = Path(asset_dir, f"{self['Identifier']} record description list.pdf")
            for pdf in filter(lambda x: x.suffix.lower() == '.pdf', self['ASSETS']):
                input = PdfReader(open(pdf, "rb"))
                for i in input.pages:
                    output.add_page(i)
            print('concatenating', ', '.join([p.name for p in pdfs]), 'to', outfile.name)
            with open(outfile, 'wb') as outputStream:
                output.write(outputStream)
            self['ASSETS'].append(outfile)
            for x in pdfs:
                self['ASSETS'].remove(x)

    def normalise(self, asset_dir):
        """clean up a row so all values are output as strings, ints or None.
        Also copies any multimedia to central folder"""
        sheet_row = []
        if self.get('ASSETS') is not None:
            self.concat_pdfs(asset_dir)
        self.copy_assets(asset_dir, 'ASSETS')
        self.copy_assets(asset_dir, 'ATTACHMENTS')
        for key, value in self.items():
            if type(value) == list:
                value = list(filter(None, value)) # strip out empty rows
                if key.startswith('#'):
                    value = '#ng#'.join(value)
                else:
                    value = '|'.join(value)
            if not bool(value): # remove empty strings and lists
                value = None
            sheet_row.append(value)
        return sheet_row


class audit_log(dict):
    def __init__(self, csv_file):
        with open(csv_file, encoding='utf_16_le') as f:
            reader = csv.DictReader(f)
            self.fieldnames = reader.fieldnames
            for row in reader:
                if self.get(row['AudKey']) is None:
                    self[row['AudKey']] = [row]
                else:
                    self[row['AudKey']].append(row)

    def get_record_log(self, irn, outdir):
        outfile = Path(outdir, irn + '.csv')
        rows = self.get(irn)
        if rows is not None:
            with open(outfile, 'w', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=self.fieldnames)
                writer.writeheader()
                writer.writerows(rows)                
            return outfile

class template_handler(dict):
    def __init__(self):
        super().__init__()
        self.fieldnames = {}

    def add_template(self, template_name):
        template_name = template_name.lower()
        for template in TEMPLATE_DIR.iterdir():
            if template_name.lower() in str(template).lower():
                with template.open(encoding='utf-8') as f:
                    print('adding template', template_name)
                    reader = csv.DictReader(f)
                    self[template_name] = []
                    self.fieldnames[template_name] = reader.fieldnames
                break

    def add_row(self, template_name, row):
        template_name = template_name.lower()
        fieldnames = self.fieldnames[template_name]
        ordered_row = order_values(row, fieldnames)
        if ordered_row not in self[template_name]:
            self[template_name].append(ordered_row)
        else:
            print(ordered_row['NODE_TITLE'], 'is already in template', template_name)
    
    def chunk_rows(self, rows, rowlimit, sort_by):
        if sort_by is not None:
            rows = sorted(rows, key=lambda row: row[sort_by])
        for i in range(0, len(rows), rowlimit):
            yield rows[i:i + rowlimit]

    def serialise(self, out_dir, rowlimit=3000, sort_by=None, batch_id=None):
        if batch_id is None:
            batch_id = uuid4()
        for template, rows in self.items():
            print(f"{template}: {len(rows)}")
            for c, chunk in enumerate(self.chunk_rows(rows, rowlimit, sort_by), 1):
                batch_name = f'{template}_{batch_id}_{c}'
                print(f"{batch_name}: {len(chunk)} rows")
                asset_dir = Path(out_dir, batch_name)
                asset_dir.mkdir(exist_ok=True)
                wb = openpyxl.Workbook()
                sheet = wb.active
                sheet.append(self.fieldnames[template])
                for row in chunk:
                    try:
                        sheet.append(row.normalise(asset_dir))
                    except ValueError as e:
                        print(row)
                        print(e)
                fpath = Path(out_dir, f'{batch_name}.xlsx')
                wb.save(fpath)

