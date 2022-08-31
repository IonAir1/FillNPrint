import ast
import pandas as pd
import os
import re
import yaml
from PIL import Image, ImageFont, ImageDraw 
import textwraps


class FillNPrint:

    def __init__(self, yaml, excel):
        self.cfg = self.parse_yaml(yaml)
        self.excel = excel
        self.progress_bar = None
        self.progress_text = None


    def parse_yaml(self, file): #parse yaml files
        if file is None:
            return
        if os.path.exists(file):
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


    #return list of sheet names
    def get_sheets(self):
        try:
            return pd.ExcelFile(self.excel).sheet_names
        except:
            return ['']


    #stamp text to image
    def stamp(self, img, text, pos, dpi, font, size, color, max_width, line_height, max_lines):
        draw = ImageDraw.Draw(img)
        position = tuple(i * dpi for i in pos)
        font_final = ImageFont.truetype(font, size)

        #wrap text
        lines = textwrap.wrap(text, width=max_width)
        lines = lines[:max_lines]
        y_text = position[1]

        #anchor to last line if height is less than 0
        if line_height < 0:
            for line in lines:
                width, height = font_final.getbbox(text)
                offset = line_height * -1
                y_text -= height * offset
        else:
            offset = line_height

        #draw text line by line
        for line in lines:
            bbox = font_final.getbbox(text)
            width = bbox[2] - bbox[0]
            height = bbox[3] - bbox[1]
            draw.text((position[0], y_text), line, color, font=font_final)
            y_text += height * offset


    #assign progress bar and text to variable
    def assign_progress(self, bar, text):
        self.progress_bar = bar
        self.progress_text = text


    #generator routine
    def generate(self, path, **kwargs):
        #kwargs
        sheet = kwargs.get('sheet', None)
        start = kwargs.get('start', 'A1')
        limit = kwargs.get('limit', None)
        print_text = kwargs.get('print', True)

        #return progress if enabled
        def progress(progress, text):
            if print_text:
                print(text)
            if not self.progress_text is None:
                self.progress_text.config(text=text)
            if not self.progress_bar is None:
                self.progress_bar['value'] = progress

        progress(0, 'Starting...')

        default_values = {
                'size': 12,
                'color': (0,0,0),
                'line-height': 1,
                'max-width': 50,
                'max-line': 1
            }
        #get list of columns listed in yaml file
        columns = []
        for item in self.cfg['text']:
            columns.append(self.cfg['text'][item]['column'].upper())
            
            #fill in missing values with default value
            for val in default_values:
                if not val in self.cfg['text'][item]:
                    self.cfg['text'][item][val] = default_values[val]
                
            #convert string tuples to numbers for yaml support
            for value in ('position', 'color'):
                if type(self.cfg['text'][item][value][0]) == str:
                    self.cfg['text'][item][value] = ast.literal_eval(self.cfg['text'][item][value])

        #fill in missing values in ['document'] with default value
        default_values = {'background': (255, 255, 255, 255), 'rotate': 0}
        for val in default_values:
            if not val in self.cfg['document']:
                self.cfg['document'][val] = default_values[val]

        #convert string tuples to numbers for yaml support
        for value in ('size', 'background'):
            if type(self.cfg['document'][value][0]) == str:
                self.cfg['document'][value] = ast.literal_eval(self.cfg['document'][value])

        #read excel file as data frame
        df = self.read_excel(self.excel, sheet=sheet, start=start, limit=limit, columns=columns)
        images = []
        document = self.cfg['document']

        #find last row
        length = len(df.index)
        for r in range(len(df.index) - 1):
            empty = True
            for item in df.iloc[r+1]:
                if not pd.isnull(item):
                    empty = False
            if empty:
                length = r+1
                break

        #generate for each row in data frame
        for r in range(length):
            progress((r+1)/length*100, "Processing ("+str(r+1)+"/"+str(length)+")")
            img = Image.new('RGB', tuple(i * document['dpi'] for i in document['size']), color=document['background'])

            #stamp for each item in text in yaml file
            for item in self.cfg['text']:
                curr = self.cfg['text'][item]

                #catch if column is non existent in data frame
                try:
                    text = df.iloc[r][self.col2num(curr['column'])]
                except:
                    text = ''

                #dont stamp if value is NaN
                if pd.isnull(text):
                    text = ''

                self.stamp(img, str(text), curr['position'], document['dpi'], curr['font'], curr['size'], curr['color'], curr['max-width'], curr['line-height'], curr['max-line'])
            images.append(img.rotate(document['rotate']*-1, expand=1))

        if not os.path.isdir(os.path.dirname(path)):
            os.makedirs(os.path.dirname(path))
        images[0].save(path, save_all=True, append_images=images[1:], resolution=document['dpi']) #save
        progress(100, 'Done!')
