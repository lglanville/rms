import sys
import csv

def get_fieldnames(template):
    with open(template, encoding='utf-8') as f:
        reader = csv.DictReader(f)
        return reader.fieldnames

def check_fieldnames(base_fieldnames, additional_fieldnames):
    for x, f in enumerate(additional_fieldnames):
        try:
            pos = base_fieldnames.index(f)
            if x != pos:
                print("Field", f, "is at column", x, "but should be at column", pos, "(should be", base_fieldnames[x], ")")
        except ValueError:
            print("Field", f, "is not in base template")
    if len(base_fieldnames) > len(additional_fieldnames):
        for f in base_fieldnames[len(additional_fieldnames):]:
            print("Field", f, "is missing from template")

def main(base_template, additional_templates):
    fieldnames = get_fieldnames(base_template)
    for t in additional_templates:
        print("Checking template", t, "against", base_template)
        comp_fieldnames = get_fieldnames(t)
        check_fieldnames(fieldnames, comp_fieldnames)

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2:])