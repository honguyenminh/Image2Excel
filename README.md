# Image2Excel
An image file to excel .xlsx file converter.  
Yes, every cell is a pixel in your image. Mind-blowing, right?

## Limitations
- **Your source image MUST have at most 256 different colors.**  
This is due to Excel's limitation of maximum 256 different fill styles.
Blame Excel I guess, can't do anything to bypass that.  
Until someone write a custom Excel file parser that does not have this
fill style limit, *this repo will not allow images with more than 256 colors.*
- **Your source image must have a resolution under 1,048,576 x 16,384**  
But in practice this rarely ever become a problem.
- **Can only write to .xlsx OpenOfficeXML format** ~~for now~~
because writing to binary .xls is hard, and libraries don't play well
with using Excel as an image viewer. Who would've thought?

## Usage
This is a console app (for now) so you have to run it from a command prompt
(or shell if you're a *NIX user)

Syntax: `Image2Excel <image-path> <output-path>(optional)`
- `Image2Excel` is the path to the executable
- `<image-path>` is the path to source image
- `<output-path>` (optional) is the path to output file.  
*If left out, this will be the image path with `.xlsx` appended*

This help can also be viewed with `Image2Excel --help` or `Image2Excel -h`

THIS APP HAS SUPER OWO POWER