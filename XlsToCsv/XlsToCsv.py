# -*- coding: utf-8 -*-

import io, sys, getopt
import xlrd

def parse_cell(s, i):
    if not isinstance(s, float):
        return str(s).replace(';', ',')
    precision = 0 if i == 0 else 2
    return ("{0:."+str(precision)+"f}").format(s).replace('.', ',')

def create_csv(in_filename, split_sheets):
    rb = xlrd.open_workbook(in_filename)
    csv = []
    for i in range(rb.nsheets):
        sheet = rb.sheet_by_index(i)
        sheet_csv = []
        for j in range(sheet.nrows):
            row = sheet.row_values(j)
            row_csv = [parse_cell(row[k], k) for k in range(len(row))]   
            sheet_csv.append(';'.join(row_csv) + '\n')
        if split_sheets:
            csv.append(sheet_csv)
        else:
            csv.extend(sheet_csv)
    return csv

def parse_cl_args(argv):
    in_filename = None
    out_filename = 'out.csv'
    split_sheets = False
    help_msg = 'Usage: XlsToCsv.py -i <input filename> [-o <output filename>] [--split]'
    try:
        opts, args = getopt.getopt(argv, "hi:o:", ["split"])
    except getopt.GetoptError:
        print(help_msg)
        sys.exit()
    for opt, arg in opts:
        if opt == '-h':
            print(help_msg)
            sys.exit()
        elif opt == '-i':
            in_filename = arg
        elif opt == '-o':
            out_filename = arg if str.endswith(arg, '.csv') else arg + '.csv'
        elif opt == '--split':
            split_sheets = True
    if in_filename == None:
        print('Input filename not specified, aborting.')
        print(help_msg)
        sys.exit()
    return in_filename, out_filename, split_sheets

def write_csv(csv, out_filename, split_sheets):
    if not split_sheets:
        with io.open(out_filename, 'w', encoding='utf8') as f:
            f.writelines(csv)
            return
    for i in range(len(csv)):
        with io.open(out_filename[:-4] + str(i) + '.csv', 'w', encoding='utf8') as f:
            f.writelines(csv[i])

def main(argv):
    in_filename, out_filename, split_sheets = parse_cl_args(argv)
    csv = create_csv(in_filename, split_sheets)
    write_csv(csv, out_filename, split_sheets)
    
if __name__ == "__main__":
    main(sys.argv[1:])