import yaml
import pandas as pd
import re

class FillNPrint:

    def parse_yaml(self, file): #parse yaml files
        with open(file, 'r') as stream:
            try:
                return yaml.safe_load(stream)
            except yaml.YAMLError as exc:
                return "error: invalid yaml file"
    

    def col2num(self, col): #convert excel column letter to integer
        c = 0
        for b in range(len(col)):
            c *= 26
            c += ord(col[b].upper()) - ord('A') + 1
        c -= 1
        return c


    def read_excel(self, file, **kwargs): #read excel files as data frame
        #kwargs
        start = kwargs.get('start', 'A1')
        limit = kwargs.get('limit', None)
        columns = kwargs.get('columns', None)
        sheet = kwargs.get('sheet', None)

        skip_row = int(re.findall(r'\d+', start)[0]) - 1 #skip row to starting_cell

        #use sheet as sheet_name if specified
        if sheet:
            df = pd.read_excel(file, sheet_name=sheet, skiprows=skip_row, header=None)
        else:
            df = pd.read_excel(file, skiprows=skip_row, header=None)

        #skip column to starting_cell
        df = df.iloc[: , self.col2num(re.sub(r"\d+", "", start)):]
        df.columns = pd.RangeIndex(df.columns.size)

        #apply limit
        if limit:
            df = df.head(limit)

        #only read up to nessecary columns
        if columns:
            last_column = 0
            for col in columns:
                colnum = self.col2num(col)
                if colnum > last_column:
                    last_column = colnum
            df = df.iloc[:, : last_column+1]

        return df

