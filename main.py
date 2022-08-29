import yaml
import pandas as pd
import re
from PIL import Image, ImageFont, ImageDraw 
import ast
import textwrap


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
    

    #create new blank image
    def new_image(self, size, dpi):
        img = Image.new('RGB', tuple(i * dpi for i in ast.literal_eval(size)), color=(255, 255, 255))
        return img
    
    #stamp text to image
    def stamp(self, img, text, pos, dpi, font, size, color, max_width, line_height, max_lines):
        draw = ImageDraw.Draw(img)
        position = tuple(i * dpi for i in ast.literal_eval(pos))
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
            print(font_final.getbbox(line), line)
            bbox = font_final.getbbox(text)
            width = bbox[2] - bbox[0]
            height = bbox[3] - bbox[1]
            draw.text((position[0], y_text), line, ast.literal_eval(color), font=font_final)
            y_text += height * offset
        