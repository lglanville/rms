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

    def identify_template(self):
        """based on some dicey conditional logic, work out what template each item should use"""
        gf = self.findall('EADGenreForm')
        template = "item"
        assets = self.findall('Multimedia')
        if assets is not None:
            exts = {Path(x).suffix for x in assets}
            if len(exts) > 1:
                raise TypeError('Multiple asset extensions for record', record)
            elif len(exts) == 1:
                if '.jpg' in exts:
                    template = 'Image'
                elif '.pdf' in exts:
                    template = 'Item'
        if gf is not None:
            gf = '|'.join(gf).lower()
            if 'architectural drawings' in gf:
                template = 'plan'
            elif gf.startswith('moving image'):
                template = 'Moving Image'
            elif gf.startswith('audio recordings'):
                template = 'sound'
        if str(self.get('EADUnitTitle')).startswith('Digital Asset:'):
            template = 'Legacy Asset'
        return template

    def get_series(self):
        """Identify the item's series (if it has one)"""
        for x in self.find_in_tuple("AssParentObjectRef", ['EADUnitID', 'EADUnitTitle', 'EADLevelAttribute']):
            if x['EADLevelAttribute'] is not None and x['EADLevelAttribute'].lower() == 'series':
                return f"[{x['EADUnitID']} {x['EADUnitTitle']}]"

    def get_accession(self):
        """Identify the item's accession (if it has one)"""
        accession_name = ""
        accrued_to = ""
        acc_lot = self.get('AccAccessionLotRef')
        lot_number = acc_lot.get('LotLotNumber')
        lot_irn = acc_lot.get('irn')
        for x in self.find_in_tuple("AssParentObjectRef", ['EADUnitID', 'EADUnitTitle', 'EADLevelAttribute']):
            if x.get("EADLevelAttribute") is not None and x['EADLevelAttribute'].lower() == 'acquisition':
                accession_name = f"{x['EADUnitID']} {x['EADUnitTitle']}"
            elif x.get("EADLevelAttribute") is not None and x['EADLevelAttribute'].lower() == 'consolidation':
                if x.get("EADUnitID") == lot_number:
                    accession_name = f"{x['EADUnitID']} {x['EADUnitTitle']}"
                else:
                    accrued_to = f"{x['EADUnitID']} {x['EADUnitTitle']}"
        if not bool(accession_name):
            if lot_irn is not None:
                accession_name = "Accesion lot " + lot_irn
        return accession_name, accrued_to

    def get_item(self):
        """Identify the item a sub-item belongs to"""
        item = ""
        for x in self.find_in_tuple("AssParentObjectRef", ['EADUnitID', 'EADUnitTitle', 'EADLevelAttribute']):
            if x.get("EADLevelAttribute") is not None and x['EADLevelAttribute'].lower() == 'item':
                return x['EADUnitTitle']

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


    def contributors(self):
        """determine whether a contributor is actually a record creator"""
        prov = []
        contrib = []
        contributors = self.get('contributors')
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

    def creation_place(self):
        p = [record.find('CreCreationPlace4'),record.find('CreCreationPlace3'), record.find('CreCreationPlace2'), record.find('CreCreationPlace1')]
        return ', '.join(filter(None, p))

    def convert_to_row(self, out_dir):
        """Convert EMu json for an item into a relatively flat dictionary mapped for ReCollect"""
        row = {}
        title, full_title = metadata_funcs.shorten_title(self.get('EADUnitTitle'))
        row['NODE_TITLE'] = title
        row['Full Title'] = full_title
        row['Series'] = self.get_series()
        row['Accession'], row['Accrued to Accession'] = self.get_accession()
        row['Part of Item'] = self.get_item()
        row['Scope and Content'] = self.get('EADScopeAndContent')
        row['Dimensions'] = self.get('EADDimensions')
        row['Internal Notes'] = self.get('NotNotes')
        row['Access Status'], row['Access Conditions'] = metadata_funcs.get_access(self)
        row['Copyright Status'] = self.guess_copyright()
        row['Conditions of Use and Reproduction'] = self.get('EADUseRestrictions')
        row['Genre/Form'] = [t.split('--')[-1] for t in self.findall('EADGenreForm')]
        row['Subject'] = self.findall('EADSubject')
        row['Subject (Agent)'] = []
        row['Subject (Agent)'].extend(self.findall('EADPersonalName'))
        row['Subject (Agent)'].extend(self.findall('EADCorporateName'))
        row['Subject (Place)'] = list(self.findall('EADGeographicName'))
        cp = self.creation_place()
        if cp is not None:
            row['Subject (Place)'].append(cp)
        row['Subject (Work)'] = list(self.findall('EADGeographicName'))
        row['###Dates'] = metadata_funcs.format_date(self.get('EADUnitDate'), self.get('EADUnitDateEarliest'), self.get('EADUnitDateLatest'))
        row['EMu IRN'] = self.get('irn')
        row['Previous System ID'] = self.get('EADUnitID')
        row['Identifier'] = "UMA-ITE-" + self.get('EADUnitID').replace('.', '')
        location = record.findall('LocCurrentLocationRef')
        row['Unit'] = record.find('LocHolderName')
        row['Location if unenclosed'] = record.find('LocLocationCode')
        row.update(self.previous_ids())
        row.update(self.contributors())
        row.update(self.facet())
        if self['AdmPublishWebNoPassword'].lower() == 'yes':
            row['Publication Status'] = 'Public'
        else:
            if self['Access Status'] == 'Closed for public access':
                row['Publication Status'] = 'Not for publication'
            else:
                row['Publication Status'] = 'Review'
        row['###Condition'] = metadata_funcs.concat_fields(self.get('ConDateChecked'), self.get('ConConditionStatus'), self.get('ConConditionDetails'))
        row['Handling Instructions'] = self.get('ConHandlingInstructions')
        row.update(metadata_funcs.get_multimedia(self))
        return row


def main(item_xml, out_dir, log_file=None, batch_id=None):
    if log_file is not None:
        audit_log = metadata_funcs.audit_log(log_file)
    templates = metadata_funcs.template_handler()
    with TemporaryDirectory(dir=out_dir) as t:
        for i in item.parse_xml(item_xml):
            row = i.convert_to_row(out_dir)
            xml_path = Path(t, metadata_funcs.slugify(row['NODE_TITLE']) + '.xml')
            record.serialise_to_xml('ecatalogue', [i], xml_path)
            row['ATTACHMENTS'] = [xml_path]
            if log_file is not None:
                log = audit_log.get_record_log(r['irn'], t)
                if log is not None:
                    row['ATTACHMENTS'].append(log)
            template_name = i.identify_template().lower()
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