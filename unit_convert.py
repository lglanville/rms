import argparse
import csv
import re
import shutil
from pathlib import Path
import re
import metadata_funcs
from emu_xml_parser import record
import openpyxl


def convert_name(name):
    n = re.compile(r'Unit (\d{1,4})', flags=re.IGNORECASE)
    m = n.search(name)
    if m is not None:
        number= int(m[1])
        return name.replace(m[0], '') + 'Unit ' + '0'*(4-len(str(number)))+str(number)
    else:
        return name


def convert_holder(record, out_dir):
    row = {}
    row['NODE_TITLE'] = convert_name(record.get('LocHolderName'))
    row['Location Type'] = record.get('LocStorageType')
    row['EMu IRN'] = record.get('irn')
    row['Home Location'] = record['LocHolderLocationRef'].get('LocLocationCode')
    items = record.get('LocCurrentLocationRef')
    if items is not None:
        units = [i for i in items if str(i.get('EADLevelAttribute')).lower() == 'unit']
        if len(units) > 1:
            print('warning: multiple unit records for', record.get('LocHolderName'))
        elif len(units) == 1:
            unit = units[0]
            row['Contents'] = metadata_funcs.concat_fields(unit.get('EADUnitTitle'), unit.get('EADScopeAndContent'), sep='\n')
            if unit.get('EADUnitDate') is not None:
                row['Date Range'] = metadata_funcs.format_date(unit.get('EADUnitDate'), unit.get('EADUnitDateEarliest'), unit.get('EADUnitDateLatest'))
            parent = unit.get('AssParentObjectRef')
            if parent.get('EADLevelAttribute') == 'Series':
                row['Series'] = parent.get('EADUnitTitle')
            else:
                row['Accession'] = parent.get('EADUnitTitle')
    return row


def main(holder_xml, out_dir, template, log_file=None):
    if log_file is not None:
        audit_log = metadata_funcs.audit_log(log_file)
    templates = metadata_funcs.template_handler()
    templates.add_template('Unit', template)
    for r in record.parse_xml(holder_xml):
        row = convert_holder(r, out_dir)
        if log_file is not None:
            row['ATTACHMENTS'] = audit_log.get_record_log(r['irn'])
        templates.add_row('Unit', row)
    templates.serialise(out_dir, sort_by='NODE_TITLE')

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu units into ReCollect data')
    parser.add_argument(
        'input', metavar='i', help='EMu location xml')
    parser.add_argument(
        'output', help='directory for multimedia assets and output sheets')
    parser.add_argument(
        '--template', '-t',
        help='ReCollect csv template')
    parser.add_argument(
        '--audit', '-a',
        help='audit log export')


    args = parser.parse_args()
    main(args.input, args.output, args.template)