# IRTlibTextImageGenerator

[![Build IRTlibTextImageGenerator Tool](https://github.com/kroehne/IRTlibTextImageGenerator/actions/workflows/dotnet.yml/badge.svg)](https://github.com/kroehne/IRTlibTextImageGenerator/actions/workflows/dotnet.yml)

Please provide as an argument the path and file name of an Excel file with texts to render into image(s). Expected column names are: 

 - text: Text to Render
 - font: Font name (Font must be installed on the system)
 - style: Font style (regular, bold, italic, or bolditalic)
 - fontsize: Font Size (e.g., 24)
 - linespacing: Line Space (relative to font size, e.g., 1.0)
 - textalignment: (Horizontal) Text Alignment (center, left, right)
 - file: Filename for the generated image (extension is used to detect image type, e.g., test.png for png; png, jpg, bmp, and gif supported)
 - width: Image Width (Pixel)
 - height: Image Height (Pixel)
 - left: Left Margin (Pixel)
 - right: Right Margin (Pixel)
 - top: Top Margin (Pixel)
 - background: R;G;B (3 numbers between 0 and 255, separated by ;
 - foreground: R;G;B (3 numbers between 0 and 255, separated by ;
 - bordercolor: R;G;B (3 numbers between 0 and 255, separated by ;
 - borderwidth: Border Width (Pixel)
