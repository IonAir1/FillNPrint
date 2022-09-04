import ast
import jsonschema
import pandas as pd
import os
import re
import textwrap
import yaml
from openpyxl import load_workbook
from PIL import Image, ImageFont, ImageDraw 


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
                    cfg = yaml.safe_load(stream)
                except yaml.YAMLError as exc:
                    return "error: invalid yaml file"

            schema_file = {
                "type": "object",
                "properties": {
                    "document": {
                        "type": "object",
                        "required": ["size", "dpi"],
                        "properties": {
                            "size": {"type": "string", "pattern": "^\d+(?:\.\d+)?[a-zA-Z]+\s*x\s*\d+(?:\.\d+)?[a-zA-Z]+$"},
                            "dpi": {"type": "number"},
                            "rotate": {"type": "number"},
                            "background": {"type": "string", "pattern": "^\((\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})\)$"},
                            "reference": {"type": "string"},
                        }
                    },
                    "text": {"type": "object"}
                }
            }

            schema_text = {
                "type": "object",
                "required": ["column", "position", "font"],
                "properties": {
                    "column": {"type": "string", "pattern": "^[a-zA-Z]$"},
                    "position": {"type": "string", "pattern": "^\d+(?:\.\d+)?[a-zA-Z]+\s*,\s*\d+(?:\.\d+)?[a-zA-Z]+$"},
                    "font": {"type": "string"},
                    "size": {"type": "number"},
                    "color": {"type": "string", "pattern": "^\((\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})\)$"},
                    "line-height": {"type": "number"},
                    "max-width": {"type": "number"},
                    "max-line": {"type": "number"}
                }
            }

            try:
                jsonschema.validate(cfg, schema=schema_file)
            except jsonschema.exceptions.ValidationError as err:
                return "config error: document: "+str(err)

            for text in cfg['text']:
                try:
                    jsonschema.validate(cfg['text'][text], schema=schema_text)
                except jsonschema.exceptions.ValidationError as err:
                    return "config error: "+str(text)+": "+str(err)

            return cfg


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


        workbook = load_workbook(file, data_only=True)

        #load excel file
        if sheet:
            worksheet = workbook[sheet]
        elif sheet in ['', None]:
            worksheet = workbook.worksheets[0]
        else:
            worksheet = workbook[sheet]
        df = pd.DataFrame(worksheet.values)

        #skip to starting_cell
        df = df.iloc[(int(re.findall(r'\d+', start)[0]) - 1): , self.col2num(re.sub(r'\d+', '', start)):]
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
        except Exception:
            return ['']


    #stamp text to image
    def stamp(self, img, text, pos, dpi, font, **kwargs):
        #kwargs
        size = kwargs.get('size', 12)
        color = kwargs.get('color', (0,0,0))
        max_width = kwargs.get('max_width', 50)
        line_height = kwargs.get('line_height', 1)
        max_lines = kwargs.get('max_lines', 1)
        error = kwargs.get('error', '')

        draw = ImageDraw.Draw(img)
        position = (int(self.to_inch(pos.replace(' ','').split(',')[0], error=error) * dpi + 0.5), int(self.to_inch(pos.replace(' ','').split(',')[1], error=error) * dpi + 0.5))

        #verify that the font is existent and is a font file
        if not os.path.isfile(font) or (not font.endswith(".ttf") and not font.endswith(".otf")):
            self.progress(0, "config error: " + error + ": nonexistent file or unsupported font \"" + font + "\"")
            exit()

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


    #convert value to inch
    def to_inch(self, input, **kwargs):
        error = kwargs.get('error', '')
        try:
            unit = re.sub(r'\d+', '', input).lower().replace('.', '').replace(',', '')
            val = float(input.lower().replace(unit, '').replace(' ', ''))
        except Exception:
            unit = ''
            val = 0

        if unit == 'cm':
            return val/2.54
        elif unit == 'mm':
            return val/25.4
        elif unit == 'm':
            return val*39.3701
        elif unit == 'ft':
            return val*12
        elif unit == 'yd':
            return val *36
        elif unit == 'in':
            return val
        #return error if unit is unknown
        else:
            self.progress(0, "config error: " + error + ": unknown or unsupported unit \"" + input + "\"")
            exit()


    #assign progress bar and text to variable
    def assign_progress(self, bar, text):
        self.progress_bar = bar
        self.progress_text = text


    #return progress if enabled
    def progress(self, progress, text):
        if self.print_text:
            print(text)
        if not self.progress_text is None:
            self.progress_text.config(text=text)
        if not self.progress_bar is None:
            self.progress_bar['value'] = progress


    #generator routine
    def generate(self, path, **kwargs):
        #kwargs
        sheet = kwargs.get('sheet', None)
        start = kwargs.get('start', 'A1')
        limit = kwargs.get('limit', None)
        self.print_text = kwargs.get('print', True)

        self.progress(0, 'Starting...')

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
            if type(self.cfg['text'][item]['color'][0]) == str:
                self.cfg['text'][item]['color'] = ast.literal_eval(self.cfg['text'][item]['color'])

        #fill in missing values in ['document'] with default value
        default_values = {'background': (255, 255, 255, 255), 'rotate': 0}
        for val in default_values:
            if not val in self.cfg['document']:
                self.cfg['document'][val] = default_values[val]

        #convert string tuples to numbers for yaml support
        if type(self.cfg['document']['background'][0]) == str:
            self.cfg['document']['background'] = ast.literal_eval(self.cfg['document']['background'])

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

        #set template
        size = tuple(document['size'].replace(' ','').split('x'))
        template = Image.new('RGB', (int(self.to_inch(size[0]) * document['dpi'] + 0.5), int(self.to_inch(size[1]) * document['dpi'] + 0.5)), color=document['background'])
        #paste reference image if required
        if 'reference' in document and os.path.isfile(document['reference']):
            reference = Image.open(document['reference'], 'r')

            #resizs reference image to max size if it exceeds max size
            if reference.size[0] > template.size[0]:
                reference = reference.resize((int(template.size[0] + 0.5), int(template.size[0] * reference.size[1] / reference.size[0] + 0.5)), Image.ANTIALIAS)
            if reference.size[1] > template.size[1]:
                reference = reference.resize((int(template.size[1] * reference.size[0] / reference.size[1] + 0.5), int(template.size[1] + 0.5)), Image.ANTIALIAS)

            template.paste(reference, (0, 0), reference.convert('RGBA'))

        #generate for each row in data frame
        for r in range(length):
            self.progress((r+1)/length*100, "Processing ("+str(r+1)+"/"+str(length)+")")
            img = template.copy()

            #stamp for each item in text in yaml file
            for item in self.cfg['text']:
                curr = self.cfg['text'][item]

                #catch if column is non existent in data frame
                try:
                    text = df.iloc[r][self.col2num(curr['column'])]
                except Exception:
                    text = ''

                #dont stamp if value is NaN
                if pd.isnull(text):
                    text = ''

                self.stamp(img, str(text), curr['position'], document['dpi'], curr['font'], size=curr['size'], color=curr['color'], max_width=curr['max-width'], line_height=curr['line-height'], max_lines=curr['max-line'], error=item)
            images.append(img.rotate(document['rotate']*-1, expand=1))

        if not os.path.isdir(os.path.dirname(path)):
            try:
                os.makedirs(os.path.dirname(path))
            except Exception:
                pass
        images[0].save(path, save_all=True, append_images=images[1:], resolution=document['dpi']) #save
        self.progress(100, 'Done!')
