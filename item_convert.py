import argparse
import csv
import re
import shutil
from pathlib import Path
import re
from tempfile import TemporaryDirectory
import metadata_funcs
from emu_xml_parser import record
import openpyxl

def guess_copyright(text):
    if text is not None:
        status = None
        statuses = {
            "Public domain": "out of copyright",
            "University copyright": "copyright owned by university of melbourne",
            "In copyright (publication rights granted)": "research purposes only",
            "In copyright": "can only be viewed at",
            "Orphan work": "all reasonable efforts"
            }
        for status, words in statuses.items():
            if words in text.lower():
                return status



def identify_template(row):
    """based on some dicey conditional logic, work out what template each item should use"""
    gf = row.get('Genre/Form')
    template = "item"
    assets = row.get('ASSETS')
    if assets is not None:
        exts = {x.suffix for x in assets}
        if len(exts) > 1:
            raise TypeError('Multiple asset extensions for record', record)
        elif len(exts) == 1:
            if '.jpg' in exts:
                template = 'Image'
            elif '.pdf' in exts:
                template = 'Document'
    if gf is not None:
        gf = '|'.join(gf).lower()
        if gf.startswith('pictures'):
            if 'architectural drawings' in gf:
                template = 'plan'
            elif 'posters' in gf:
                template = 'poster'
            else:
                template = 'image'
        elif gf.startswith('moving image'):
            template = 'Moving Image'
        elif gf.startswith('audio recordings'):
            template = 'sound'
        elif gf.startswith('documents'):
            template = 'document'
    if str(row.get('NODE_TITLE')).startswith('Digital Asset:'):
        template = 'Legacy Asset'
    return template

def get_sets(record):
    record_sets = {}
    for x in t.findall(".//tuple[@name='AssParentObjectRef']"):
        if x is not None:
            level = x.find("atom[@name='EADLevelAttribute']")
            name = x.find("atom[@name='EADUnitTitle']")
            ident = x.find("atom[@name='EADUnitID']")
            n = f"[{ident}] {name}"
            if x.text.lower() == 'series':
                record_sets['Series'] = n
            elif x.text.lower() == 'acquisition':
                record_sets['Accession'] = n
            elif x.text.lower() == 'consolidation':
                if t.find(".//atom[@name='LotLotNumber']") == ident:
                    record_sets['Accession'] = n
                else:
                    record_sets['Accrued to Accession'] = n

    return record_sets

def get_series(record):
    """Identify the item's series (if it has one)"""
    series = ""
    while record.get("AssParentObjectRef") is not None:
        record = record.get("AssParentObjectRef")
        if record.get("EADLevelAttribute") is not None and record['EADLevelAttribute'].lower() == 'series':
            return record['EADUnitTitle']

def get_accession(record):
    """Identify the item's accession (if it has one)"""
    accession_name = ""
    accrued_to = ""
    acc_lot = record.get('AccAccessionLotRef')
    lot_number = acc_lot.get('LotLotNumber')
    lot_irn = acc_lot.get('irn')
    while record.get("AssParentObjectRef") is not None:
        record = record.get("AssParentObjectRef")
        if record.get("EADLevelAttribute") is not None and record['EADLevelAttribute'].lower() == 'acquisition':
            accession_name = f"{record['EADUnitID']} {record['EADUnitTitle']}"
        elif record.get("EADLevelAttribute") is not None and record['EADLevelAttribute'].lower() == 'consolidation':
            if record.get("EADUnitID") in lot_number:
                accession_name = f"{record['EADUnitID']} {record['EADUnitTitle']}"
            else:
                accrued_to = f"{record['EADUnitID']} {record['EADUnitTitle']}"
    if not bool(accession_name):
        if lot_irn is not None:
            accession_name = "Accesion lot " + lot_irn
    return accession_name, accrued_to

def get_item(record):
    """Identify the item a sub-item belongs to"""
    item = ""
    while record.get("AssParentObjectRef") is not None:
        record = record.get("AssParentObjectRef")
        if record.get("EADLevelAttribute") is not None and record['EADLevelAttribute'].lower() == 'item':
            return record['EADUnitTitle']

def previous_ids(record):
    """disambiguate the EADPreviousID_tab field to specific sorts of identifier"""
    id_dict = {}
    format_number = re.compile(r'((BWP|CP|OSB|BWN|CN|GPN|NN|SL|PA)[A-D]?/?\d{1,5})', flags=re.IGNORECASE)
    UMAIC_number = re.compile(r"UMA/I/\d{1,5}", flags=re.IGNORECASE)
    faid_number = re.compile('\d{1,2}(/\d{1,2})*', flags=re.IGNORECASE)
    for id in record.findall('EADPreviousID'):
        if format_number.match(id):
            id_dict['Format Number'] = id
        elif UMAIC_number.match(id):
            id_dict['UMAIC ID'] = id
        elif faid_number.match(id):
            id_dict['Finding Aid Reference'] = id
        elif id.startswith('YFA'):
            id_dict['Classification'] = id
        else:
            id_dict['Other IDs'] = id
    return id_dict

def job_no(record):
    r = re.compile('yfa job no.? \d{1,3}')
    return r.findall(record['EADScopeAndContent'])


def facet(record):
    """disambiguate the EADPhysicalDescription_tab field to specific physical facets where possible"""
    facet_dict = {'Physical Facet' : []}
    facets = record.get('EADPhysicalFacet_tab')
    if facets is not None:
        for facet in facet:
            facet = facet.get('EADPhysicalFacet')
            if facet.lower().startswith('colour:'):
                facet_dict['Colour Depth'] = facet.split(':').strip()
            elif facet.lower() in ('colour', 'black and white', 'hand coloured', 'sepia'):
                facet_dict['Colour Depth'] = facet
            elif facet.lower().startswith('base:'):
                facet_dict['Base Material'] = facet.split(':').strip()
            elif facet.lower().startswith('duration:'):
                facet_dict['Duration'] = facet.split(':').strip()
            else:
                facet_dict['Physical Facet'].append(facet)
    facet_dict['Physical Facet'] = '|'.join(facet_dict['Physical Facet'])
    return facet_dict


def contributors(record):
    """determine whether a contributor is actually a record creator"""
    prov = []
    contrib = []
    contributors = record.get('contributors')
    if contributors is not None:
        print(contributors)
        creator_roles = ["Artist", "Architect", "Creator", "Author", "Director", "Photographer", "Producer"]
        for c in contributors:
            role = c.get('AssRelatedPartiesRelationship')
            if role is None:
                if 'photograph' in ''.join(metadata_funcs.flatten_table(record, 'EADGenreForm_tab')).lower():
                    role = 'Photographer'
                else:
                    role = ''
            if role in creator_roles:
                prov.append(str(c['NamCitedName']) + '|' + role)
            else:
                contrib.append(str(c['NamCitedName']) + '|' + role)

    return {"###Provenance": "#ng#".join(prov), "###Contributor": "#ng#".join(contrib)}

def creation_place(record):
    p = [record.find('CreCreationPlace4'),record.find('CreCreationPlace3'), record.find('CreCreationPlace2'), record.find('CreCreationPlace1')]
    return ', '.join(filter(None, p))

def convert_item(record, out_dir):
    """Convert EMu json for an item into a relatively flat dictionary mapped for ReCollect"""
    row = {}
    title, full_title = metadata_funcs.shorten_title(record.get('EADUnitTitle'))
    row['NODE_TITLE'] = title
    row['Full Title'] = full_title
    row['Series'] = get_series(record)
    row['Accession'], accrued_to = get_accession(record)
    if bool(accrued_to):
        row['Accrued to Accession'] = accrued_to
    row['Part of Item'] = get_item(record)
    row['Scope and Content'] = record.get('EADScopeAndContent')
    row['Dimensions'] = record.get('EADDimensions')
    row['Internal Notes'] = record.get('NotNotes')
    row['Access Status'], row['Access Conditions'] = metadata_funcs.get_access(record)
    row['Copyright Status'] = guess_copyright(record.get('EADUseRestrictions'))
    row['Conditions of Use and Reproduction'] = record.get('EADUseRestrictions')
    gf = metadata_funcs.flatten_table(record, 'EADGenreForm_tab')
    if gf is not None:
        row['Genre/Form'] = [t.split('--')[-1] for t in gf]
    row['Subject'] = metadata_funcs.flatten_table(record, 'EADSubject_tab')
    subjects = []
    subjects.extend(metadata_funcs.flatten_table(record, 'EADPersonalName_tab'))
    subjects.extend(metadata_funcs.flatten_table(record, 'EADCorporateName_tab'))
    row['Subject (Agent)'] = []
    row['Subject (Agent)'].extend(metadata_funcs.flatten_table(record, 'EADPersonalName_tab'))
    row['Subject (Agent)'].extend(metadata_funcs.flatten_table(record, 'EADCorporateName_tab'))
    row['Subject (Place)'] = list(record.findall('EADGeographicName'))
    cp = creation_place(record)
    if cp is not None:
        row['Subject (Place)'].append(cp)
    row['Subject (Work)'] = metadata_funcs.flatten_table(record, 'EADTitle_tab')
    row['###Dates'] = metadata_funcs.format_date(record.get('EADUnitDate'), record.get('EADUnitDateEarliest'), record.get('EADUnitDateLatest'))+'|'
    row['EMu IRN'] = record.get('irn')
    row['Previous System ID'] = record.get('EADUnitID')
    row['Identifier'] = "UMA-ITE-" + record.get('EADUnitID').replace('.', '')
    location = record.findall('LocCurrentLocationRef')
    row['Unit'] = record.find('LocHolderName')
    row['Location if unenclosed'] = record.find('LocLocationCode')
    row.update(previous_ids(record))
    row.update(contributors(record))
    row.update(facet(record))
    if record['AdmPublishWebNoPassword'].lower() == 'yes':
        row['Publication Status'] = 'Public'
    else:
        if row['Access Status'] == 'Closed for public access':
            row['Publication Status'] = 'Not for publication'
        else:
            row['Publication Status'] = 'Review'
    row['###Condition'] = metadata_funcs.concat_fields(record.get('ConDateChecked'), record.get('ConConditionStatus'), record.get('ConConditionDetails'))
    row['Handling Instructions'] = record.get('ConHandlingInstructions')
    row.update(metadata_funcs.get_multimedia(record))
    return row


def main(item_xml, out_dir, log_file=None, batch_id=None):
    if log_file is not None:
        audit_log = metadata_funcs.audit_log(log_file)
    templates = metadata_funcs.template_handler()
    with TemporaryDirectory(dir=out_dir) as t:
        for r in record.parse_xml(item_xml):
            row = convert_item(r, out_dir)
            xml_path = Path(t, metadata_funcs.slugify(row['NODE_TITLE']) + '.xml')
            record.serialise_to_xml('ecatalogue', [r], xml_path)
            row['ATTACHMENTS'] = [xml_path]
            if log_file is not None:
                log = audit_log.get_record_log(r['irn'], t)
                if log is not None:
                    row['ATTACHMENTS'].append(log)
            template_name = identify_template(row).lower()
            if templates.get(template_name) is None:
                templates.add_template(template_name)
            templates.add_row(template_name, row)
        templates.serialise(out_dir, sort_by='Identifier', batch_id=batch_id)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu items into ReCollect data')
    parser.add_argument(
        'input', metavar='i', help='EMu catalogue xml')
    parser.add_argument(
        'output', help='directory for multimedia assets and output sheets')

    parser.add_argument(
        '--audit', '-a',
        help='audit log export')
    parser.add_argument(
        '--batch_id', '-b',
        help='name for batch (if not supplied, will use a uuid)')


    args = parser.parse_args()
    main(args.input, args.output, log_file=args.audit, batch_id=args.batch_id)