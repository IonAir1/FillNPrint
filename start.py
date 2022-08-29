from fillnprint import FillNPrint
import argparse

parser = argparse.ArgumentParser(prog="fillnprint")
parser.add_argument("excel", help="excel input file")
parser.add_argument("yaml", help="config file")
parser.add_argument("output", nargs='?', default="ouput.pdf", help="path and name for output file")
parser.add_argument("-s","--sheet", help="specify sheet name or number")
parser.add_argument("-c","--cell", help="specify starting cell")
parser.add_argument("-l","--limit", help="limit the ammount of rows to read")
args = parser.parse_args()

com = "FillNPrint('{}', '{}').generate('{}'".format(str(args.yaml), str(args.excel), str(args.output))
if args.sheet != None:
    com = com + ", sheet='{}'".format(str(args.sheet))
if args.cell != None:
    com = com + ", cell='{}'".format(str(args.cell))
if args.limit != None:
    com = com + ", cell={}".format(str(args.limit))

exec(com+')')