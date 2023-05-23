import argparse
import re
from pathlib import Path
from pprint import pprint
import metadata_funcs
from emu_xml_parser import record
import openpyxl

def extract_redirect(html_file):
    p = re.compile(r'window\.location\.href = "(.*)"')
    with html_file.open() as filter:
        t = filter.read()
        m = p.search(t)
        if m is not None:
            return m[1]


def get_finding_aid(record):
    multi = record.get('MulMultiMediaRef_tab')
    asset_dict = {'ASSETS': [], 'ATTACHMENTS': []}
    if multi is not None:
        for m in multi:
            if m.get('Multimedia') is not None:
                fpath = Path(m.get('Multimedia'))
                if m['AdmPublishWebNoPassword'].lower() == 'no':
                    asset_dict['ATTACHMENTS'].append(fpath)
                elif fpath.suffix == '.html':
                    url = extract_redirect(fpath)
                    asset_dict['Other Finding Aids'] = url
                else:
                    asset_dict['ASSETS'].append(fpath)
    return asset_dict

def previous_ids(record):
    ids = metadata_funcs.flatten_table(record, 'EADPreviousID_tab')
    id_dict = {}
    if ids is not None:
        for id in ids:
            if id.startswith('UM'):
                id_dict['Records Services ID'] = id
            else:
                id_dict['Other IDs'] = id
    return id_dict

def accession_data(record):
    acc_data = {}
    acc_data['Legacy Data'] = record.get('AdmOriginalData')
    acc_lot = record.get('AccAccessionLotRef')
    if record.get('TitOwnersNameRef_tab') is not None:
        acc_data['Ownership'] = [x.get('NamFullName') for x in record.get('TitOwnersNameRef_tab')]
    if acc_lot is not None:
        acc_data['EMu Accession Lot IRN'] = acc_lot.get('irn')
        acc_data['Method of Acquisition'] = acc_lot.get('AcqAcquisitionMethod')
        acc_data['Authorised by'] = acc_lot.get('AcqAuthorisedBy')
        acc_data['Lot Description'] = acc_lot.get('LotDescription')
        acc_data['EMu Accession Lot IRN'] = acc_lot.get('irn')
        acc_data['Acquisition Notes'] = '\n'.join(filter(None, (acc_lot.get('AcqAcquisitionRemarks'), acc_lot.get('NotNotes'))))
        acc_data['Date Received'] = metadata_funcs.format_date(acc_lot.get('AcqDateReceived'), record.get('AcqDateReceivedLower'), record.get('AcqDateReceivedUpper'))
        sources = acc_lot.get('source')
        acc_data['Transferror'] = []
        if sources is not None:
            for source in sources:
                acc_data['Transferror'].append(source.get('NamFullName'))
        agreements = acc_lot.get('MulMultiMediaRef_tab')
        acc_data['Deposit Agreement'] = []
        if agreements is not None:
            for agreement in agreements:
                acc_data['Deposit Agreement'].append(agreement.get('MulTitle'))
    return acc_data


def provenance(record):
    prov = record.get('EADOriginationRef_tab')
    if prov is not None:
        prov = filter(None, [x['NamFullName'] for x in prov])
    else:
        prov = record["AssParentObjectRef"].get("EADOriginationRef_tab")
        if prov is not None:
            prov = filter(None, [x['NamFullName'] for x in prov])
    if prov is not None:
        return "#ng#".join(prov)

def convert_recordset(record):
    row = {}
    title, full_title = metadata_funcs.shorten_title(record.get('EADUnitTitle'))
    row['NODE_TITLE'] = title
    row['Full Title'] = full_title
    if record.get('EADLevelAttribute').lower() == 'series':
        row['Identifier'] = 'UMA-SRE-' + record.get('EADUnitID').replace('.', '')
    row['Scope and Content'] = record.get('EADScopeAndContent')
    row['Access Status'], row['Access Conditions'] = metadata_funcs.get_access(record)
    row['Appraisal'] = record.get('EADAppraisalInformation')
    row['Arrangement'] = record.get('EADArrangement')
    row['Extent'] = metadata_funcs.flatten_table(record, 'EADExtent_tab')
    row['Genre/Form'] = metadata_funcs.flatten_table(record, 'EADGenreForm_tab')
    row['Collection Category'] = record.get('TitObjectCategory')
    row['Accruals'] = record.get('EADAccruals')
    row['Archival History'] = record.get('EADCustodialHistory')
    row['Internal Notes'] = record.get('NotNotes')
    row['Subject'] = metadata_funcs.flatten_table(record, 'EADSubject_tab')
    row['Related Places'] = metadata_funcs.flatten_table(record, 'EADGeographicName_tab')
    row['###Dates'] = metadata_funcs.format_date(record.get('EADUnitDate'), record.get('EADUnitDateEarliest'), record.get('EADUnitDateLatest'))+'|'
    row['EMu Catalogue IRN'] = record.get('irn')
    row['Descriptive Note'] = record.get('EADOtherFindingAid')
    row['Previous System ID'] = record.get('EADUnitID')
    row['###Provenance'] = provenance(record)
    if record['AdmPublishWebNoPassword'].lower() == 'yes':
        row['Publication Status'] = 'Public'
    else:
        row['Publication Status'] = 'Not for publication'
    row['###Condition'] = metadata_funcs.concat_fields(record.get('ConDateChecked'), record.get('ConConditionStatus'), record.get('ConConditionDetails'))
    row['Handling Instructions'] = record.get('ConHandlingInstructions')
    row.update(get_finding_aid(record))
    row.update(previous_ids(record))
    return row


def convert_accession(record, row):
    row['Identifier'] = 'UMA-ACE-' + record.get('EADUnitID').replace('.', '')
    row.update(accession_data(record))
    return row

def create_accession(record):
    acc_lot = record.get('AccAccessionLotRef')
    if acc_lot is not None:
        row = {}
        acc_num = str(record.get('LotLotNumber'))
        row['NODE_TITLE'] = acc_lot.get('SummaryData')
        row['Publication Status'] = 'Not for publication'
        row = accession_data(record)
        if acc_num in record['EADUnitID']:
            row['Identifier'] = 'UMA-ACE-' + record.get('EADUnitID').replace('.')
            row['NODE_TITLE'] = record.get('EADUnitTitle')
        else:
            row['NODE_TITLE'] = acc_lot.get('SummaryData')
        return row

def is_accession(record):
    level = record.get('EADLevelAttribute')
    if level in ('Acquisition', 'Consolidation'):
        return True
    elif level == 'Series':
        acc_lot = record.get('AccAccessionLotRef')
        if acc_lot is not None:
            if record['EADUnitID'] in str(acc_lot['LotLotNumber']):
                return True
            else:
                return False
    else:
        return False


def main(emu_xml, out_dir, accession_template, series_template, log_file=None):
    if log_file is not None:
        audit_log = metadata_funcs.audit_log(log_file)
    templates = metadata_funcs.template_handler()
    templates.add_template('Accession', accession_template)
    templates.add_template('Series', series_template)
    for r in record.parse_xml(emu_xml):
        row = convert_recordset(r)
        if log_file is not None:
            row['ATTACHMENTS'] = audit_log.get_record_log(r['irn'])
        if r['EADLevelAttribute'].lower() in ('acquisition', 'consolidation'):
            row = convert_accession(r, row)
            templates.add_row('Accession', row)
        else:
            templates.add_row('Series', row)
            acc = create_accession(r)
            if acc is not None:
                templates.add_row('Accession', acc)
    templates.serialise(out_dir)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu items into ReCollect data')
    parser.add_argument(
        'input', metavar='i', help='EMu catalogue xml')
    parser.add_argument(
        'output', help='directory for multimedia assets and output sheets')
    parser.add_argument(
        'acc_template',
        help='ReCollect accession csv template')
    parser.add_argument(
        'ser_template',
        help='ReCollect series csv template')
    parser.add_argument(
        '--audit', '-a',
        help='audit log export')


    args = parser.parse_args()
    main(args.input, args.output, args.acc_template, args.ser_template)