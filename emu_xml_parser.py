from lxml import etree
import sys
from pprint import pprint
from collections.abc import MutableMapping
import csv

class record(dict):
    def __repr__(self):
        return super(record, self).__repr__()

    @classmethod
    def parse_tuple(cls, tuple_elem):
        "Parse out a record within an EMu xml report to JSON"
        new_rec = cls()
        for elem in tuple_elem:
            if elem.tag == 'atom':
                new_rec[elem.attrib['name']] = elem.text
            elif elem.tag == 'tuple':
                new_rec[elem.attrib['name']] = cls.parse_tuple(elem)
            elif elem.tag == 'table':
                new_rec[elem.attrib['name']] = []
                for record in elem.iterfind('tuple'):
                    new_rec[elem.attrib['name']].append(cls.parse_tuple(record))
        return new_rec

    @classmethod
    def parse_xml(cls, xml_doc):
        "iterator to generate JSON EMu records from xml"
        with open(xml_doc, encoding='utf-8') as f:
            text = f.read()
            for x in range(8):
                text = text.replace(chr(x), ' ')
            text = text.replace(chr(19), '')
            text = text.replace(chr(24), '')
            text = text.replace(chr(25), '')
            text = text.replace('\u2014', '-')
            text = text.replace('\u2013', '-')
            text = text.replace('encoding="UTF-8"', '')
        tree = etree.fromstring(text)
        for record in tree.iterfind('tuple'):
            yield cls.parse_tuple(record)

    def merge_field(self, **kwargs):
        """merge a given field based on kwargs - table fields are appended as
        new levels, text fields are added as new lines"""
        for key, val in kwargs.items():
            if key.endswith('_tab'):
                if self.get(key) is not None:
                    self[key].append(val)
                else:
                    self[key] = [val]
            else:
                self[key] += '\n' + val

    def migrate_field(self, source, dest):
        "Migrate a field to another field, blanking the source field"
        val = self[source]
        if dest.endswith('_tab'):
            self[dest] = [val]
        else:
            self[dest] = val
        self[source] = None

    def to_xml(self):
        "Export a record as EMu xml"
        root_element = etree.Element('tuple')
        for key, value in self.items():
            if type(value) == str:
                etree.SubElement(root_element, 'atom', name=key).text = value
            if type(value) == record:
                tup_element = value.to_xml()
                tup_element.attrib['name'] = key
                root_element.append(tup_element)
            elif type(value) == list:
                tab_element = etree.SubElement(root_element, 'table', name=key)
                for val in value:
                    if type(val) != record:
                        tup_element = etree.SubElement(tab_element, 'tuple')
                        field, _ = key.split('_')
                        etree.SubElement(tup_element, 'atom', name=field).text = val
                    else:
                        tab_element.append(val.to_xml())
        return root_element

    def flatten(self, parent_key='', sep='.'):
        "Flatten the record JSON structure to a format suitable for CSV output"
        items = []
        for k, v in self.items():
            new_key = parent_key + sep + k if parent_key else k
            if isinstance(v, MutableMapping):
                items.extend(v.flatten(new_key, sep=sep).items())
            elif isinstance(v, list):
                for x in v:
                    items.extend(x.flatten(new_key, sep=sep).items())
            else:
                items.append((new_key, v))
        row = {}
        for field, value in items:
            if value is None:
                value = ''
            if not row.get(field):
                row[field] = value
            else:
                row[field] += '|' + value
        return row

    def print_xml(self):
        return etree.tostring(self.to_xml(), pretty_print=True).decode()

    @staticmethod
    def serialise_to_xml(table, records, out_file):
        table_elem = etree.Element('table', name=table)
        for rec in records:
            table_elem.append(rec.to_xml())
        tree = etree.ElementTree(table_elem)
        tree.write(
            out_file, pretty_print=True, standalone=True, xml_declaration=True,
            encoding='UTF-8')

    @classmethod
    def serialise_to_csv(cls, in_file, out_file):
        fieldnames = set()
        rows = []
        for r in cls.parse_xml(in_file):
            row = r.flatten()
            fieldnames.update(row.keys())
            rows.append(row)
        fieldnames = list(fieldnames)
        fieldnames.sort()
        with open(out_file, 'w', encoding='utf-8-sig', newline='') as f:
            writer= csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)

if __name__ == '__main__':
    record.serialise_to_csv(sys.argv[1], sys.argv[2])
