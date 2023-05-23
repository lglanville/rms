import argparse
import csv
import re
import shutil
from pathlib import Path
import re
import metadata_funcs
from emu_xml_parser import record
import openpyxl


def format_strdate(birth, death):
    if birth is not None:
        birth += '-'
        if death is not None:
            birth += death
    return birth

def convert_agent(record, out_dir):
    row = {}
    row['NODE_TITLE'] = record.get('NamFullName')
    row['Internal Notes'] = record.get('NotNotes')
    row['EMu IRN'] = record.get('irn')
    row['Legacy Data'] = record.get('AdmOriginalData')
    dates = metadata_funcs.format_date(format_strdate(record.get('BioBirthDate'), record.get('BioDeathDate')), record.get('BioBirthEarliestDate'), record.get('BioDeathLatestDate'))
    if record.get('NamPartyType').lower() == 'person':
        row['Title'] = record.get('NamTitle')
        row['Given Name'] = record.get('NamFirst')
        row['Middle Name'] = record.get('NamMiddle')
        row['Family Name'] = record.get('NamLast')
        row['Suffix'] = record.get('NamSuffix')
        row['Other Names'] = metadata_funcs.flatten_table(record, 'NamOtherNames_tab')
        row['Biography'] = record.get('BioCommencementNotes')
        row['Place of Birth'] = record.get('BioBirthPlace')
        row['Place of Death'] = record.get('BioDeathPlace')
        row['Activities & Occupations'] = metadata_funcs.flatten_table(record, 'NamSpecialities_tab')
        row['Telephone (Business)'] = metadata_funcs.flatten_table(record, 'NamBusiness_tab')
        row['Telephone (Mobile)'] = record.get('AddWeb')
        row['Telephone (Home)'] = record.get('NamHome')
        row['Gender'] = record.get('NamSex')
        row['Dates'] = dates
    elif record.get('NamPartyType').lower() == 'organisation':
        row['Other Names'] = metadata_funcs.flatten_table(record, 'NamOrganisationOtherNames_tab')
        row['Activities'] = metadata_funcs.flatten_table(record, 'NamSpecialities_tab')
        row['History'] = '\n'.join(filter(None, (record.get('HisBeginDateNotes'), record.get('HisEndDateNotes'))))
        row['###Dates'] = dates
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


def main(agent_xml, out_dir, template, log_file=None):
    if log_file is not None:
        audit_log = metadata_funcs.audit_log(log_file)
    templates = metadata_funcs.template_handler()
    for r in record.parse_xml(agent_xml):
        row = convert_item(r, out_dir)
        if log_file is not None:
            row['ATTACHMENTS'] = audit_log.get_record_log(r['irn'])
        template_name = r.get('NamPartyType')
        if templates.get(template_name) is None:
            templates.add_template(template_name, template)
        templates.add_row(template_name, row)
    templates.serialise(out_dir)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu items into ReCollect data')
    parser.add_argument(
        'input', metavar='i', help='EMu catalogue xml')
    parser.add_argument(
        'output', help='directory for assets and output sheets')
    parser.add_argument(
        '--template', '-t',
        help='ReCollect csv template')
    parser.add_argument(
        '--audit', '-a',
        help='audit log export')


    args = parser.parse_args()
    main(args.input, args.output, args.template)