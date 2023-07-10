import argparse
import csv
import re
import shutil
from pathlib import Path
import re
import metadata_funcs
from emu_xml_parser import record
from tempfile import TemporaryDirectory
import openpyxl


def format_strdate(birth, death):
    if birth is not None:
        birth += '-'
        if death is not None:
            birth += death
    return birth

class agent(record):
    def convert_to_row(self, out_dir):
        row = {}
        row['NODE_TITLE'] = self.get('NamCitedName')
        row['Internal Notes'] = self.get('NotNotes')
        row['EMu IRN'] = self.get('irn')
        row['Legacy Data'] = self.get('AdmOriginalData')
        row['Agent Type'] = self.get('NamPartyType')
        if self.get('AdmPublishWebNoPassword').lower() == 'yes':
            row['Publication Status'] = 'Public'
        else:
            row['Publication Status'] = 'Not for publication'
        dates = metadata_funcs.format_date(format_strdate(self.get('BioBirthDate'), self.get('BioDeathDate')), self.get('BioBirthEarliestDate'), self.get('BioDeathLatestDate'))
        row['###Dates'] = dates
        row['History'] = '\n'.join(filter(None, (self.get('BioCommencementNotes'), self.get('HisBeginDateNotes'), self.get('HisEndDateNotes'))))
        row['Title'] = self.get('NamTitle')
        row['Given Name'] = self.get('NamFirst')
        row['Middle Name'] = self.get('NamMiddle')
        row['Family Name'] = self.get('NamLast')
        row['Suffix'] = self.get('NamSuffix')
        row['Other Names'] = []
        row['Other Names'].extend(metadata_funcs.flatten_table(self, 'NamOtherNames_tab'))
        row['Other Names'].extend(metadata_funcs.flatten_table(self, 'NamOrganisationOtherNames_tab'))
        row['Acronym'] = self.get('NamOrganisationAcronym')
        row['Place of Birth'] = self.get('BioBirthPlace')
        row['Place of Death'] = self.get('BioDeathPlace')
        row['Activities & Occupations'] = metadata_funcs.flatten_table(self, 'NamSpecialities_tab')
        row['Telephone (Business)'] = metadata_funcs.flatten_table(self, 'NamBusiness_tab')
        row['Telephone (Mobile)'] = self.get('NamMobile')
        row['Telephone (Home)'] = self.get('NamHome')
        row['Gender'] = self.get('NamSex')
        row['Website'] = self.get('AddWeb')
        row['Email'] = self.get('AddEmail')
    
        home_address = metadata_funcs.concat_fields(
            self.get('AddPhysStreet'),
            self.get('AddPhysCity'),
            self.get('AddPhysState'),
            self.get('AddPhysPost'),
            self.get('AddPhysCountry'))
        post_address = metadata_funcs.concat_fields(
            self.get('AddPostStreet'), 
            self.get('AddPostCity'), 
            self.get('AddPostState'), 
            self.get('AddPostPost'), 
            self.get('AddPostCountry'))
        addr = []
        if home_address is not None:
            addr.append('Physical|' + home_address)
        if post_address is not None:
            addr.append('Postal|' + post_address)
        row['###Address'] = '#ng#'.join(addr)
        return row


def main(agent_xml, out_dir, log_file=None):
    if log_file is not None:
        audit_log = metadata_funcs.audit_log(log_file)
    templates = metadata_funcs.template_handler()
    template_name = 'People-and-Organisations'
    templates.add_template(template_name)
    with TemporaryDirectory(dir=out_dir) as t:
        for r in agent.parse_xml(agent_xml):
            row = r.convert_to_row(out_dir)
            xml_path = Path(t, metadata_funcs.slugify(row['NODE_TITLE']) + '.xml')
            record.serialise_to_xml('eparties', [r], xml_path)
            row['ATTACHMENTS'] = [xml_path]
            if log_file is not None:
                log = audit_log.get_record_log(r['irn'], t)
                if log is not None:
                    row['ATTACHMENTS'].append(log)
            templates.add_row(template_name, row)
        templates.serialise(out_dir, sort_by="NODE_TITLE")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu items into ReCollect data')
    parser.add_argument(
        'input', metavar='i', help='EMu catalogue xml')
    parser.add_argument(
        'output', help='directory for assets and output sheets')
    parser.add_argument(
        '--audit', '-a',
        help='audit log export')


    args = parser.parse_args()
    main(args.input, args.output, log_file=args.audit)