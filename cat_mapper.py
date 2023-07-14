import csv
import argparse
from emu_xml_parser import record

class ReCollect_report():
    def __init__(self, csv_file):
        self.rows = []
        with open(csv_file, encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if 'Item ID' in row.keys():
                    row['Node ID'] = row.pop('Item ID')
                if 'Item Title' in row.keys():
                    row['Node Title'] = row.pop('Item Title')
                assert 'Node ID' in row.keys()
                self.rows.append(row)

    def retrieve_row(self, fieldname, value):
        return [row for row in self.rows if row.get(fieldname) == value]

def main(emu_xml, recollect_csv, output_csv, map_on_field, map_to_field, fields_to_map):
    node_ids = ReCollect_report(recollect_csv)
    fieldnames = ['Node ID', 'Node Type', 'Node Title', '#REDACT']
    fieldnames.extend(fields_to_map)
    with open(output_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in record.parse_xml(emu_xml):
            rows = node_ids.retrieve_row(map_on_field, r.get(map_to_field))
            for row in rows:
                row = {'Node ID': row['Node ID'], 'Node Title': row['Node Title']}
                for x in fields_to_map:
                    row[x] = '|'.join(r.findall(x))
                writer.writerow(row)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert EMu items into ReCollect data')
    parser.add_argument(
        'emu_xml', metavar='i', help='EMu catalogue xml')
    parser.add_argument(
        'recollect_csv', metavar='i', help='ReCollect report or metadata download')
    parser.add_argument(
        'output_csv', metavar='o', help='EMu accession lot xml')
    parser.add_argument(
        '--map_fields', '-m', nargs=2, default=['Previous System ID', 'EADUnitID'], help='Equivalent ReCollect and Emu fields to map data')
    parser.add_argument(
        '--fields', '-f', nargs='+',
        help='fields to map from EMu')
    args = parser.parse_args()
    main(args.emu_xml, args.recollect_csv, args.output_csv, *args.map_fields, args.fields)
