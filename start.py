import argparse
import os
from fillnprint import FillNPrint

parser = argparse.ArgumentParser(prog="fillnprint")
parser.add_argument("excel", help="excel input file")
parser.add_argument("yaml", help="config file")
parser.add_argument("output", nargs='?', default="ouput.pdf", help="path and name for output file")
parser.add_argument("-s","--sheet", help="specify sheet name or number")
parser.add_argument("-c","--cell", help="specify starting cell")
parser.add_argument("-l","--limit", help="limit the ammount of rows to read")
args = parser.parse_args()

fnp_inst = FillNPrint(args.yaml, args.excel)
sheets = fnp_inst.get_sheets()

#exception handling
if not args.excel.endswith('.xlsx') or not os.path.exists(args.excel):
    print("Invalid excel file")
    exit()
if fnp_inst.cfg == "error: invalid yaml file" or not os.path.exists(args.yaml):
    print("Invalid yaml file")
    exit()
if "config error:" in fnp_inst.cfg:
    error = fnp_inst.cfg.split("\n")
    print(error[0])
    exit()
if not args.output.endswith('.pdf'):
    print("Output file must be a pdf file")
    exit()
if not args.sheet is None and (not args.sheet in sheets and str(args.sheet).upper().isupper()):
    print("Selected sheet is not a valid sheet")
    exit()
if not args.limit is None and str(args.limit).upper().isupper():
        print("'Limit' setting must be an integer or left empty")
        exit()

com = "fnp_inst.generate('{}'".format(str(args.output))
if not args.sheet is None:
    if args.sheet.upper().isupper():
        com = com + ", sheet='{}'".format(str(args.sheet))
    else:
        com = com + ", sheet=" + str(args.sheet)
if args.cell != None:
    com = com + ", cell='{}'".format(str(args.cell))
if args.limit != None:
    com = com + ", limit=" + str(args.limit)

exec(com+')')