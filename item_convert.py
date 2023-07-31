import argparse
import csv
import re
import shutil
from pathlib import Path
import re
from tempfile import TemporaryDirectory
import logging
import metadata_funcs
from emu_xml_parser import record
import openpyxl
from cat_mapper import ReCollect_report

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

class item(record):
    def guess_copyright(self):
        text = self.get('EADUseRestrictions')
        if text is not None:
            status = None
            statuses = {
                "Public domain": "out of copyright",
                "University copyright": "copyright owned by university of melbourne",
                "In copyright - publication rights granted": "research purposes only",
                "In copyright": "can only be viewed at",
                "Orphan work": "all reasonable efforts"
                }
            for status, words in statuses.items():
                if words in text.lower():
                    return status

    def get_unit(self):
        unit_name = self.find('LocHolderName')
        if unit_name is not None:
            repl = lambda x: x.group(1) +" " + "0"*(4-len(x.group(2)))+x.group(2)
            return re.sub(r'(unit|album) (\d+)', repl, unit_name, flags=re.IGNORECASE)

    def get_location(self):
        loc_name = self.find('LocLocationCode')
        if loc_name is not None:
            repl = lambda x: "0"*(2-len(x.group(0)))+x.group(0)
            return re.sub('\d+', repl, loc_name, flags=re.IGNORECASE)

    def get_genreform(self):
        gf = []
        for x in self.findall('EADGenreForm'):
            for term in x.split('-'):
                if term not in gf and bool(term):
                    gf.append(term.lower())
        return gf

    def is_parent(self):
        parent = False
        subitems = self.get('Subitems')
        if subitems is not None:
            parent = True
        return parent

    def identify_genreform(self):
        template = 'item'
        gf = self.get_genreform()
        if 'pictures' in gf:
            if 'design drawings' in gf:
                template = 'plan'
            else:
                template = 'image'
        elif 'moving image' in gf:
             template = 'moving-image'
        elif 'audio recordings' in gf:
            template = 'sound'
        return template

    def get_exts(self):
        exts = set()
        assets = self.findall('Multimedia')
        if assets is not None:
            exts = {Path(x).suffix for x in assets}
        return exts

    def identify_template(self):
        """based on some dicey conditional logic, work out what template each item should use"""
        template = self.identify_genreform()
        exts = self.get_exts()
        if template == 'item':
            if '.jpg' in exts or '.tif' in exts:
                template = 'image'
        if self.get('EADLevelAttribute').lower() == 'multiple items':
            template = 'multiple-items'
        if self.is_parent():
            template = 'item'
        if str(self.get('EADUnitTitle')).startswith('Digital Asset:'):
            template = 'legacy-asset'
        return template


    def get_parent_records(self, acc_report):

        """Identify the item's accession (if it has one)"""
        data = {}
        lot_irn = None
        for x in self.find_in_tuple('AccAccessionLotRef', ['irn']):
            if x['irn'] is not None:
                lot_irn = x['irn']
                break
        if lot_irn is not None:
            rows = acc_report.retrieve_row("EMu Accession Lot IRN", lot_irn)
            data['Accession'] = '|'.join([r['Node Title'] for r in rows])
        for x in self.find_in_tuple("AssParentObjectRef", ['EADUnitID', 'EADUnitTitle', 'EADLevelAttribute']):
            if x['EADLevelAttribute'] is not None:
                level = x['EADLevelAttribute'].lower()
                name = f"[{x['EADUnitID']}] {x['EADUnitTitle']}"
                if level == 'series':
                    data['Series'] = name
                elif level == 'item':
                    title, _ = metadata_funcs.shorten_title(x.get('EADUnitTitle'))
                    data['Part of Item'] = f"[{x['EADUnitID']}] {title}"
                elif level == 'acquisition':
                    data['Accession'] = name
                elif level == 'consolidation':
                    if data.get('Accession') is not None:
                        if data.get('Accession') != name:
                            data['Accrued to Accession'] = name
                    else:
                        data['Accession'] = name
                else:
                    logger.warning("Unrecognised parent: "+ " ".join(x.values()))
        if data.get('Accession') is None:
            if lot_irn is not None:
                data['Accession'] = "Accession lot " + lot_irn
        return data

    def previous_ids(self):
        """disambiguate the EADPreviousID_tab field to specific sorts of identifier"""
        id_dict = {}
        format_number = re.compile(r'((BWP|CP|OSB|BWN|CN|GPN|NN|SL|PA)[A-D]?/?\d{1,5})', flags=re.IGNORECASE)
        UMAIC_number = re.compile(r"UMA/I/\d{1,5}", flags=re.IGNORECASE)
        faid_number = re.compile('\d{1,2}(/\d{1,2})*', flags=re.IGNORECASE)
        for id in self.findall('EADPreviousID'):
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

    def facet(self):
        """disambigu`ate the EADPhysicalDescription_tab field to specific physical facets where possible"""
        facet_dict = {'Physical Facet' : []}
        facets = self.get('EADPhysicalFacet_tab')
        if facets is not None:
            for facet in facet:
                facet = facet.get('EADPhysicalFacet')
                if facet.lower().startswith('colour:'):
                    facet_dict['Colour Depth'] = facet.split(':')[1].strip()
                elif facet.lower() in ('colour', 'black and white', 'hand coloured', 'sepia'):
                    facet_dict['Colour Depth'] = facet
                elif facet.lower().startswith('base:'):
                    facet_dict['Base Material'] = facet.split(':')[1].strip()
                elif facet.lower().startswith('duration:'):
                    facet_dict['Duration'] = facet.split(':')[1].strip()
                else:
                    facet_dict['Physical Facet'].append(facet)
        facet_dict['Physical Facet'] = '|'.join(facet_dict['Physical Facet'])
        return facet_dict


    def contributors(self):
        """determine whether a contributor is actually a record creator"""
        prov = []
        contrib = []
        creator_roles = ["artist", "architect", "creator", "author", "director", "photographer", "producer"]
        
        for x in self.find_in_table('contributors', ['AssRelatedPartiesRelationship', 'NamCitedName']):
            role = x.get('AssRelatedPartiesRelationship')
            if role is None:
                if 'photograph' in '|'.join(self.findall('EADGenreForm')).lower():
                    role = 'Photographer'
                else:
                    role = ''
            if role.lower().strip('s') in creator_roles:
                prov.append((str(x['NamCitedName']), role))
            else:
                contrib.append((str(x['NamCitedName']), role))
        for x in self.find_in_table('EADOriginationRef_tab', ['NamCitedName']):
            
            if x['NamCitedName'] is not None and x['NamCitedName'] not in [i[0] for i in prov]:
                prov.append((x['NamCitedName'], 'Provenance'))
        return {"###Provenance": "#ng#".join(['|'.join(x) for x in prov]), "###Contributor": "#ng#".join(['|'.join(x) for x in contrib])}

    def creation_place(self):
        p = [self.find('CreCreationPlace4'),self.find('CreCreationPlace3'), self.find('CreCreationPlace2'), self.find('CreCreationPlace1')]
        return ', '.join(filter(None, p))
    
    def get_dates(self):
        string_date = self.get('EADUnitDate')
        if string_date is None:
            string_date = 'Undated'
        date = metadata_funcs.format_date(string_date, self.get('EADUnitDateEarliest'), self.get('EADUnitDateLatest'))
        if string_date.lower() == 'undated':
            for x in self.find_in_tuple('AssParentObjectRef', ['EADUnitDate', 'EADUnitDateEarliest', 'EADUnitDateLatest', 'EADLevelAttribute']):
                d = metadata_funcs.format_date(x.get('EADUnitDate'), x.get('EADUnitDateEarliest'), x.get('EADUnitDateLatest'))
                if d is not None:
                    level = x.get('EADLevelAttribute')
                    if level.lower() in ('acquisition', 'consolidation'):
                        level = 'accession'
                    date = "Undated;;#ng#" + d + "|Date of " + level
                    break
        return date

    def get_title(self):
        title, full_title = metadata_funcs.shorten_title(self.get('EADUnitTitle'))
        if self.is_parent():
            title = f"[{self.get('EADUnitID')}] {title}"
        return title, full_title

    def convert_to_row(self, acc_report, out_dir):
        """Convert EMu json for an item into a relatively flat dictionary mapped for ReCollect"""
        row = metadata_funcs.Row()

        row['NODE_TITLE'], row['Full Title'] = self.get_title()
        row.update(self.get_parent_records(acc_report))
        row['Scope and Content'] = self.get('EADScopeAndContent')
        row['Dimensions'] = self.get('EADDimensions')
        row['Internal Notes'] = '\n'.join(filter(None, [self.get('NotNotes') self.get('EADAppraisalInformation')]))
        row['Access Status'], row['Access Conditions'] = metadata_funcs.get_access(self)
        row['Copyright Status'] = self.guess_copyright()
        row['Conditions of Use and Reproduction'] = self.get('EADUseRestrictions')
        row['Genre/Form'] = [t.split('--')[-1] for t in self.findall('EADGenreForm')]
        row['Subject'] = list(self.findall('EADSubject'))
        row['Subject (Agent)'] = []
        row['Subject (Agent)'].extend(self.findall('EADPersonalName'))
        row['Subject (Agent)'].extend(self.findall('EADCorporateName'))
        row['Subject (Place)'] = list(self.findall('EADGeographicName'))
        cp = self.creation_place()
        if cp is not None:
            row['Subject (Place)'].append(cp)
        row['Subject (Work)'] = list(self.findall('EADGeographicName'))
        row['###Dates'] = self.get_dates()
        row['EMu IRN'] = self.get('irn')
        row['Previous System ID'] = self.get('EADUnitID')
        row['Identifier'] = "UMA-ITE-" + self.get('EADUnitID').replace('.', '')
        row['Unit'] = self.get_unit()
        row['Location if unenclosed'] = self.get_location()

        if row['Unit'] is not None:
            row['Request Type'] = 'Request unit'
        else:
            row['Request Type'] = 'Request item'
        if self.get('TitObjectStatus') is not None:
            if self.get('TitObjectStatus').lower() == 'missing':
                row['Request Type'] = 'This item is not available'
                row['Status'] = 'Location unknown'
        row.update(self.previous_ids())
        row.update(self.contributors())
        row.update(self.facet())
        if self['AdmPublishWebNoPassword'].lower() == 'yes':
            row['Publication Status'] = 'Public'
        else:
            if row['Access Status'] == 'Closed for public access':
                row['Publication Status'] = 'Not for publication'
            else:
                row['Publication Status'] = 'Review'
        row['###Condition'] = metadata_funcs.concat_fields(self.get('ConDateChecked'), self.get('ConConditionStatus'), self.get('ConConditionDetails'))
        row['Handling Instructions'] = self.get('ConHandlingInstructions')
        row['ASSETS'] = metadata_funcs.get_multimedia(self)
        if bool(row['ASSETS']):
            if metadata_funcs.is_redacted(self):
                row['#REDACT'] = "Yes"
                if row['Access Status'] != "Closed for public access":
                    row['Menu'] = "Browse digitised items#n#Available in the Reading Room"
            else:
                row['Menu'] = "Browse digitised items#n#Available online"
        return row


def main(item_xml, accession_csv, out_dir, log_file=None, batch_id=None, logging=False):
    if log_file is not None:
        audit_log = metadata_funcs.audit_log(log_file)
    templates = metadata_funcs.template_handler(batch_id=batch_id)
    acc_report = ReCollect_report(accession_csv)
    metadata_funcs.configlogfile(Path(out_dir, templates.batch_id + '.log'), logger)
    with TemporaryDirectory(dir=out_dir) as t:
        for i in item.parse_xml(item_xml):
            row = i.convert_to_row(acc_report, out_dir)
            xml_path = Path(t, metadata_funcs.slugify(row['Identifier']) + '.xml')
            record.serialise_to_xml('ecatalogue', [i], xml_path)
            row['ATTACHMENTS'] = [xml_path]
            if log_file is not None:
                log = audit_log.get_record_log(r['irn'], t)
                if log is not None:
                    row['ATTACHMENTS'].append(log)
            template_name = i.identify_template().lower()
            templates.add_row(template_name, row)
        templates.serialise(out_dir, sort_by='Identifier')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu items into ReCollect data')
    parser.add_argument(
        'input_xml', metavar='i', help='EMu catalogue xml')
    parser.add_argument(
        'accession_csv', metavar='i', help='recollect accession csv for mapping')
    parser.add_argument(
        'output', help='directory for multimedia assets and output sheets')
    parser.add_argument(
        '--audit', '-a',
        help='audit log export')
    parser.add_argument(
        '--batch_id', '-b',
        help='name for batch (if not supplied, will use a uuid)')


    args = parser.parse_args()
    main(args.input_xml, args.accession_csv, args.output, log_file=args.audit, batch_id=args.batch_id, logging=True)