import argparse
import re
from pathlib import Path
from pprint import pprint
from tempfile import TemporaryDirectory
import metadata_funcs
from emu_xml_parser import record

def convert_agreement(record):
    row = {}
    title = record.get('MulTitle')
    if title is None:
        title = record.get('MulIdentifier')
    row['NODE_TITLE'] = title
    row['Notes'] = record.get('MulDescription')
    row['Type'] = record.get('DetResourceType')
    row['Donor'] = list(set(record.findall('NamCitedName')))
    row['EMu IRN'] = record.get('irn')
    fpath = Path(record.get('Multimedia'))
    row['ASSETS'] = [fpath]
    return row


def main(multimedia_xml, out_dir, log_file=None):
    if log_file is not None:
        audit_log = metadata_funcs.audit_log(log_file)
    templates = metadata_funcs.template_handler()
    template_name = 'Agreement'
    templates.add_template(template_name)
    with TemporaryDirectory(dir=out_dir) as t:
        for r in record.parse_xml(multimedia_xml):
            row = convert_agreement(r)
            xml_path = Path(t, metadata_funcs.slugify(row['NODE_TITLE']) + '.xml')
            record.serialise_to_xml('emultimedia', [r], xml_path)
            row['ATTACHMENTS'] = [xml_path]
            if log_file is not None:
                log = audit_log.get_record_log(r['irn'], t)
                if log is not None:
                    row['ATTACHMENTS'].append(log)
            templates.add_row(template_name, row)
        templates.serialise(out_dir)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu multimedia items into ReCollect agreements')
    parser.add_argument(
        'input', metavar='i', help='EMu multimedia xml')
    parser.add_argument(
        'output', help='directory for assets and output sheets')
    parser.add_argument(
        '--audit', '-a',
        help='audit log export')


    args = parser.parse_args()
    main(args.input, args.output, log_file=args.audit)