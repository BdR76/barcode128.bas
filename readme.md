barcode128.bas
==============

Barcode128 Encoded is a macro function for LibreOffice Calc or Excel to
encode texts for use with a barcode 128 font. To create a barcode that can be
read by a handscanner, it is not enough to just change the font a barcode font.

![LibreOffice barcode 128 encoded texts](/preview.png?raw=true "preview")

The text or product code must first be encoded, the actually barcode consists of:
* A start symbol
* One or more encoded data characters
* A checksum symbol
* A stop symbol

This results in a compact barcode, using as little space as possible,
especially when the text contains only digits.
For more information see [Wikipedia](https://en.wikipedia.org/wiki/Code_128).

How to use
----------

1) Open LibreOffice Calc
2) Go to Tools > Macros > Edit macros
3) Go to File > Import Basic and import the .bas file
4) Go to sheet cell A1, paste a product code
5) In cell B1 type the function `=BARCODE128_ENCODED(A1)`
6) Set barcode of cell B1 to Barcode 128 font

Barcode font
------------
Use the [Libre Barcode 128 font](https://fonts.google.com/specimen/Libre+Barcode+128+Text) 
on the encoded texts to print scannable barcodes.

History
-------
2020-mar-30 upload to github

Questions, comments -- Bas de Reuver (bdr1976@gmail.com)
