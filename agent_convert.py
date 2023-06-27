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

def convert_agent(record, out_dir):
    row = {}
    row['NODE_TITLE'] = record.get('NamCitedName')
    row['Internal Notes'] = record.get('NotNotes')
    row['EMu IRN'] = record.get('irn')
    row['Legacy Data'] = record.get('AdmOriginalData')
    row['Agent Type'] = record.get('NamPartyType')
    if record.get('AdmPublishWebNoPassword').lower() == 'yes':
        row['Publication Status'] = 'Public'
    else:
        row['Publication Status'] = 'Not for publication'
    dates = metadata_funcs.format_date(format_strdate(record.get('BioBirthDate'), record.get('BioDeathDate')), record.get('BioBirthEarliestDate'), record.get('BioDeathLatestDate'))
    row['###Dates'] = dates
    row['History'] = '\n'.join(filter(None, (record.get('BioCommencementNotes'), record.get('HisBeginDateNotes'), record.get('HisEndDateNotes'))))
    row['Title'] = record.get('NamTitle')
    row['Given Name'] = record.get('NamFirst')
    row['Middle Name'] = record.get('NamMiddle')
    row['Family Name'] = record.get('NamLast')
    row['Suffix'] = record.get('NamSuffix')
    row['Other Names'] = []
    row['Other Names'].extend(metadata_funcs.flatten_table(record, 'NamOtherNames_tab'))
    row['Other Names'].extend(metadata_funcs.flatten_table(record, 'NamOrganisationOtherNames_tab'))
    row['Acronym'] = record.get('NamOrganisationAcronym')
    row['Place of Birth'] = record.get('BioBirthPlace')
    row['Place of Death'] = record.get('BioDeathPlace')
    row['Activities & Occupations'] = metadata_funcs.flatten_table(record, 'NamSpecialities_tab')
    row['Telephone (Business)'] = metadata_funcs.flatten_table(record, 'NamBusiness_tab')
    row['Telephone (Mobile)'] = record.get('AddWeb')
    row['Telephone (Home)'] = record.get('NamHome')
    row['Gender'] = record.get('NamSex')
    row['Website'] = record.get('AddWeb')
    row['Email'] = record.get('AddEmail')
    
    home_address = metadata_funcs.concat_fields(
        record.get('AddPhysStreet'),
        record.get('AddPhysCity'),
        record.get('AddPhysState'),
        record.get('AddPhysPost'),
        record.get('AddPhysCountry'))
    post_address = metadata_funcs.concat_fields(
        record.get('AddPostStreet'), 
        record.get('AddPostCity'), 
        record.get('AddPostState'), 
        record.get('AddPostPost'), 
        record.get('AddPostCountry'))
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
        for r in record.parse_xml(agent_xml):
            row = convert_agent(r, out_dir)
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