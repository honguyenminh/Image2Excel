# Image2Excel
An image file to excel .xlsx file converter.  
Yes, every cell is a pixel in your image. Mind-blowing, right?

## Limitations
- **Your source image should have at most 256 different colors to keep the same color**  
This is due to Excel's limitation of maximum 256 different fill styles.  
**Update: Now the app can use normal image too, but the image will be quantized (indexed color) to 256 color before converting**
- **Your source image must have a resolution under 1,048,576 x 16,384**  
But in practice this rarely ever becomes a problem.
- **Can only write to .xlsx OpenOfficeXML format**
because writing to binary .xls is hard, and libraries don't play well
with using Excel as an image viewer. Who would've thought? 🤷

## Usage
This is a console app (for now) so you have to run it from a command prompt
(or terminal if you're a \*NIX user)

Syntax: `Image2Excel <image-path> <output-path>(optional)`
- `Image2Excel` is the path to the executable
- `<image-path>` is the path to source image
- `<output-path>` (optional) is the path to output file.  
*If left out, this will be the image path with `.xlsx` appended*

This help can also be viewed with `Image2Excel --help` or `Image2Excel -h`

THIS APP HAS SUPER OWO POWER
