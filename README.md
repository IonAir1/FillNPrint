# FillNPrint
## Print excel data directly to paper forms!    
### Table of Contents:  
[How to download](#how-to-download)  
[How to use the gui](#how-to-use-the-gui)  
[Using the source code](#using-the-source-code)  
[Importing FIllNPrint to other projects](#importing-fillnprint-to-other-projects)  
[Creating a config file](#creating-a-config-file)    


### How to download  
[Click here](https://github.com/IonAir1/fillnprint/releases/tag/v1.2.0) to go to downloads page and select the appropriate version.
For Windows and Linux, executables are available but for MacOS, you will have to use the source code.
**Note that windows defender may mark the executable as suspicious due to pyinstaller using self extracting exe files. you can either make a folder with an exclusion then directly extract the file there or you may choose to download the source code instead**    

### How to use the gui  
1. **Excel File** - select the the excel file to be used  
2. **Sheet** - select the specific sheet. if left empty, it will default to the first sheet  
3. **Starting Cell** - start reading from a specific cell. i.e if B2 is the input, only the data from row 2 will be read and the new column A will be column B, column B will be column C, etc. if left empty, it will read from A1  
4. **Limit** - will only read the first n rows of data from the starting cell with n as the limit. if left empty, it will read until the last column or until it reaches an empty row.  
5. **Configuration File** - select the configuration file to be used. if you haven't already, go to [Creating a config file](#creating-a-config-file) to create one.  
6. **Output File** - select the path and filename fot the pdf output file
7. **Generate** - click the generate button an wait for it to finish

### Using the source code
After you have dwnloaded the code, installed python and ran the setup.py, You can run the gui.py for the gui and start.py for the cli. You may also use and import fillnprint.py to your other python projects and call the functions directly. **Note that FillNPrint was tested and built on python 3.9.2**  

### Importing FIllNPrint to other projects  

1. `FillNPrint(excel file, yaml file)` with excel file containing the data to be printed ans yaml file containing the config file written in yaml.    

2. `.parse_yaml(file)` - checks and validates the yaml file while returning either an error or the data from the file as a dictionary.    

3. `.col2num(col)` - accepts a column letter from excel and returns its numeric equivalent i.e A->0, B->1, E->4.    

4. `read_excel(file, **kwargs)` - reads and returns the excel file as a data frame.
   - `start` - read from a specific cell. i.e if B2 is the input, only the data from row 2 will be read and the new column a will be column B, column b will be column c, etc.  
   - `limit` - will only read the first n rows of data from the starting cell with n as the limit. leaving it as None will let it read until the last column or until it reaches an empty row.  
   - `columns` - accepts a list of columns, it will only read up to the alphabetically last column in the list.  
   - `sheet` - will read the specific sheet. accepts sheet name or number. defaults to the first sheet.    
 
5. `.get_sheets()` - returns a list of sheet names from the excel file    

6. `.stamp(img, text, pos, dpi, font, **kwargs)` - stamps text to an image with img as a PIL image, text as the text to be stamped, pos as the real world placement of the stamp, dpi as the dpi of the image and font as the font to be used.  
   - `size` - font size, default is 12.  
   - `color` - font color, default is rgb(0,0,0).  
   - `max_width` - maximum characters per line. default is 50.  
   - `line_height` - line spacing. by using negative numbers, the text will align from the bottom instead of from the top.  
   - `max_lines` - maximum lines allowed. default is 1.  
   - `error` - this is the code name (usually from the config file) of the current data/column being stamped. this is used for showing more specific errors.    

7. `.to_inch(input, **kwargs)` - converts number+unit string input to inches. accepts the following units: mm,cm,m,in,ft,yd. i.e 5cm -> 1.9685.  
   - `error` - this is the code name (usually from the config file) of the current data/column being stamped. this is used for showing more specific errors.   

8. `.assign_progress(bar, text)` - allow you to assign tkinter progress bar and tkinter label for output.  

9. `.progress(progress, text)` - prints text and sets the assigned tkinter label to text. also sets the assigned tkinter progress bar to to progress.  

10. `.generate(path, **kwargs)` - procedural function to generate the pdfs with path as the save file.  
    - `sheet` - will read the specific sheet. accepts sheet name or number. defaults to the first sheet.    
    - `start` - read from a specific cell. i.e if B2 is the input, only the data from row 2 will be read and the new column a will be column B, column b will be column c, etc.  
    - `limit` - will only read the first n rows of data from the starting cell with n as the limit. leaving it as None will let it read until the last column or until it reaches an empty row.  
    - `print_text` - if True, will output progress text to the commandline or terminal.    
 
 ### Creating a config file  
The config file will be used to determine which data is to be placed where. The config file shall use **YAML** syntax. (every yaml file start with three dahses `---` . they are also dependent on the indentation so be careful with that. make sure all indentations made are **two spaces**. yaml files also follow the format `<key>: <value>`)  

example:  
```
---
document:
  size: 5in x 7in
  dpi: 96
text:
  name:
    column: A
    position: 5cm,10cm
    font: Roboto-Regular.ttf
    size: 12
    line-height: -2
    max-width: 10
    max-line: 3
  number:
    column: B
    position: 2cm, 2cm
    font: Roboto-Regular.ttf
    size: 10
```     

1. open notepad or any of your preffered text editor.    
2. the first line shall always be three dashes `---`. This will tell the program that it is a yaml file.    
3. the first section is the `document:` section. This will contain information abot the form itself.    
   - Required fields:  
     - `size: <width><unit> x <height><unit>` - this will determine the real life size of the form.  
     - `dpi: <number>` - enter the resolution/quality of the print. common dpi settings are 96, 300, 600, 900, etc.  
   - Optional Fields:  
     - `background: (<r>,<g>,<b>,<a>)` - this will set the image background to the set rgba color. default is (255,255,255,255).  
     - `rotate: <number>` - rotates the image to the set number. value is in degress but only enter the value itself without the degree sign. default is 0  
     - `reference: <path to image>` - will set background to image. (meant to be used by scanning the form itself and placing it here as a reference as to where the text fields should be placed). if image is bigger than the page, it will automatically resize to fit, else it will be in actual size.
     - `print-size: <width><unit> x <height><unit> , <x-offset><unit> , <y-offset><unit>` - will paste the forms on an image of a different size. This can be used for printing purposes wherein the page size cannot be configured manually.  
4. the second sectio is the `text: ` section. Tths will conatain information about what data to place and where to place it. you can put as much "code" as you want with the format `<code name>: `. with \<code name> being any name but prefferably something that will describe the data it uses. i.e  

```
text:
  code1:
    #<this is where the text fields will go>
  code2:
    #<this is where the text fields will go>
  code3:
    #<this is where the text fields will go>
```    

5. each "code" shall contain the following fields for more info about the data.
   - Required Fields:
     - `column: <letter>` - enter the column letter of the text data from the excel file. the starting cell will be column A.
     - `position: <x_position><unit>, <y_position><unit>` - enter the actual position of where you would like to place the text. anchor is the top left.
     - `font: <path to .ttf/.otf file>` - enter the font file you would like to use.
   - Optional Fields:
     - `size: <number>` - this is the font size. default is 12.
     - `color: (<r>,<g>,<b>)` - this is the text color. default is (0,0,0).
     - `line-height: <number>` - this is theline spacing. using negative numbers will align text to bottom while using positive numbers will align from top (when number is positive, do not include a "+" sign). default is 1.
     - `max-width: <number>` - max number of characters until it goes to the next line. default is 50.
     - `max-line: <number>` - max nmber of lines until it cuts off the text. default is 1.
6. after you are finished, save the file as anything you want. file extension also does not matter.



