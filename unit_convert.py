import argparse
import csv
import re
import shutil
from pprint import pprint
from pathlib import Path
import re
import logging
from tempfile import TemporaryDirectory
import metadata_funcs
from emu_xml_parser import record
import openpyxl

logging.basicConfig(
    format=f'%(asctime)s %(levelname)s %(message)s', level=logging.INFO)
formatter = logging.Formatter(
    '%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
logger.propagate = False
ch = logging.StreamHandler()
ch.setFormatter(formatter)
ch.setLevel(logging.INFO)
logger.addHandler(ch)


class unit(record):

    def convert_name(self):
        name = self.get('LocHolderName')
        if name is not None:
            repl = lambda x: x.group(1) +" " + "0"*(4-len(x.group(2)))+x.group(2)
            return re.sub(r'(unit|album) (\d+)', repl, name, flags=re.IGNORECASE)

    def get_location(self):
        loc_name = self['LocHolderLocationRef'].get('LocLocationCode')
        if loc_name is not None:
            repl = lambda x: "0"*(2-len(x.group(0)))+x.group(0)
            return re.sub('\d+', repl, loc_name, flags=re.IGNORECASE)

    def publish_online(self):
        status = "Not for publication"
        p = list(set([x.lower() for x in self.findall('AdmPublishWebNoPassword')]))
        if len(p) == 1:
            if p[0] == 'yes':
                status = 'Public'
        return status

    def to_xml_file(self, out_dir):
        xml_path = Path(out_dir, metadata_funcs.slugify(self.convert_name()) + '.xml')
        self.serialise_to_xml('elocations', [self], xml_path)
        return xml_path

    def convert_to_row(self, out_dir):
        row = {}
        row['NODE_TITLE'] = self.convert_name()
        row['Unit Type'] = self.get('LocStorageType')
        row['EMu IRN'] = self.get('irn')
        row['Internal Notes'] = self.get('NotNotes')
        row['Home Location'] = self.get_location()
        items = self.get('LocCurrentLocationRef')
        row['Publication Status'] = self.publish_online()
        if 'LocCurrentLocationRef' in self.keys():
            t = lambda x: x['EADLevelAttribute'].lower() == 'unit'
            self['LocCurrentLocationRef'] = list(filter(t, self['LocCurrentLocationRef']))
            if len(self['LocCurrentLocationRef']) > 1:
                print('warning: multiple unit records for', self.get('LocHolderName'))
                pprint(self['LocCurrentLocationRef'])
            elif len(self['LocCurrentLocationRef']) == 1:
                unit = self['LocCurrentLocationRef'][0]
                row['Description'] = metadata_funcs.concat_fields(unit.get('EADUnitTitle'), unit.get('EADScopeAndContent'), sep='\n')
                if unit.get('EADUnitDate') is not None:
                    row['Date Range'] = metadata_funcs.format_date(unit.get('EADUnitDate'), unit.get('EADUnitDateEarliest'), unit.get('EADUnitDateLatest'))
                parent = unit.get('AssParentObjectRef')
                parent_name = f"[{parent.get('EADUnitID')}] {parent.get('EADUnitTitle')}"
                if parent.get('EADLevelAttribute') == 'Series':
                    row['Series'] = parent_name
                else:
                    row['Accession'] = parent_name
        return row


def main(holder_xml, out_dir, log_file=None):
    if log_file is not None:
        audit_log = metadata_funcs.audit_log(log_file)
    templates = metadata_funcs.template_handler()
    metadata_funcs.configlogfile(Path(out_dir, templates.batch_id + '.log'), logger)
    templates.add_template('unit')
    with TemporaryDirectory(dir=out_dir) as t:
        for u in unit.parse_xml(holder_xml):
            row = u.convert_to_row(out_dir)
            row['ATTACHMENTS'] = [u.to_xml_file(t)]
            if log_file is not None:
                log = audit_log.get_record_log(u['irn'], t)
                if log is not None:
                    row['ATTACHMENTS'].append(log)
            templates.add_row('unit', row)
        templates.serialise(out_dir, sort_by='NODE_TITLE', rowlimit=3000)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu units into ReCollect data')
    parser.add_argument(
        'input', metavar='i', help='EMu location xml')
    parser.add_argument(
        'output', help='directory for multimedia assets and output sheets')
    parser.add_argument(
        '--audit', '-a',
        help='audit log export')


    args = parser.parse_args()
    main(args.input, args.output, log_file=args.audit)