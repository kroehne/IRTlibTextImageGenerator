using System;
using System.IO;
using System.Collections.Generic;
using ExcelDataReader;
using SixLabors.Fonts;
using SixLabors.ImageSharp; 
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.ImageSharp.PixelFormats; 
using SixLabors.ImageSharp.Processing;
using System.Globalization;

namespace TextImageGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("IRTlib: TextImageGenerator ({0})\n", typeof(Program).Assembly.GetName().Version.ToString());
            Console.ResetColor();
              
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            bool useBlindText = false;
            string blindTextTemplate = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, arcu. In enim justo, rhoncus ut, imperdiet a, venenatis vitae, justo. Nullam dictum felis eu pede mollis pretium. Integer tincidunt. Cras dapibus. Vivamus elementum semper nisi. Aenean vulputate eleifend tellus. Aenean leo ligula, porttitor eu, consequat vitae, eleifend ac, enim. Aliquam lorem ante, dapibus in, viverra quis, feugiat a, tellus. Phasellus viverra nulla ut metus varius laoreet. Quisque rutrum. Aenean imperdiet. Etiam ultricies nisi vel augue. Curabitur ullamcorper ultricies nisi. Nam eget dui. Etiam rhoncus. Maecenas tempus, tellus eget condimentum rhoncus, sem quam semper libero, sit amet adipiscing sem neque sed ipsum. Nam quam nunc, blandit vel, luctus pulvinar, hendrerit id, lorem. Maecenas nec odio et ante tincidunt tempus. Donec vitae sapien ut libero venenatis faucibus. Nullam quis ante. Etiam sit amet orci eget eros faucibus tincidunt. Duis leo. Sed fringilla mauris sit amet nibh. Donec sodales sagittis magna. Sed consequat, leo eget bibendum sodales, augue velit cursus nunc,";

            if (args.Length < 1)
            {
                Console.WriteLine("Please provide excel-file with texts to render into image(s). Expected column names are: ");
                Console.WriteLine(" - text: Text to Render");
                Console.WriteLine(" - font: Font name (Font must be installed on the system)");
                Console.WriteLine(" - style: Font style (regular, bold, italic, or bolditalic)");
                Console.WriteLine(" - fontsize: Font Size (e.g., 24)");
                Console.WriteLine(" - linespacing: Line Space (relative to font size, e.g., 1.0)");
                Console.WriteLine(" - textalignment: (Horizontal) Text Alignment (center, left, right)");
                Console.WriteLine(" - file: File name for the generated image (extension is used to detect image type, e.g., test.png for png; png, jpg, bmp and gif supported)");
                Console.WriteLine(" - width: Image Width (Pixel)");
                Console.WriteLine(" - height: Image Height (Pixel)");
                Console.WriteLine(" - left: Left Margin (Pixel)");
                Console.WriteLine(" - right: Right Margin (Pixel)");
                Console.WriteLine(" - top: Top Margin (Pixel)");
                Console.WriteLine(" - background: R;G;B (3 numbers between 0 and 255, separated by ;");
                Console.WriteLine(" - foreground: R;G;B (3 numbers between 0 and 255, separated by ;");
                Console.WriteLine(" - bordercolor: R;G;B (3 numbers between 0 and 255, separated by ;");
                Console.WriteLine(" - borderwidth: Border Width (Pixel)");
            }
            else
            {
                try
                {
                    if (args.Length > 1)
                    {
                        bool.TryParse(args[1], out useBlindText);
                    }
                    if (File.Exists(args[0]))
                    {
                        Dictionary<string, int> _columnOrder = new Dictionary<string, int>();
                        using (var stream = File.Open(args[0], FileMode.Open, FileAccess.Read))
                        {
                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                bool isHeader = true;
                                do
                                {
                                    while (reader.Read())
                                    {
                                     
                                        if (isHeader && reader.GetString(0).ToLower().Trim() != "text")
                                        {
                                            isHeader = false;
                                        }

                                        if (!isHeader)
                                        {
                                            try
                                            {
                                                int _width = 778;
                                                if (_columnOrder.ContainsKey("width"))
                                                    int.TryParse(reader.GetValue(_columnOrder["width"]).ToString(), out _width);

                                                int _height = 60;
                                                if (_columnOrder.ContainsKey("height"))
                                                    int.TryParse(reader.GetValue(_columnOrder["height"]).ToString(), out _height);

                                                int _left = 10;
                                                if (_columnOrder.ContainsKey("left"))
                                                    int.TryParse(reader.GetValue(_columnOrder["left"]).ToString(), out _left);

                                                int _right = 10;
                                                if (_columnOrder.ContainsKey("right"))
                                                    int.TryParse(reader.GetValue(_columnOrder["right"]).ToString(), out _right);

                                                int _top = 5;
                                                if (_columnOrder.ContainsKey("top"))
                                                    int.TryParse(reader.GetValue(_columnOrder["top"]).ToString(), out _top);

                                                float _fontsize = 24;
                                                if (_columnOrder.ContainsKey("fontsize")) 
                                                  float.TryParse(reader.GetValue(_columnOrder["fontsize"]).ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out _fontsize);

                                                float _lineSpacing = (float)0.95;
                                                if (_columnOrder.ContainsKey("linespacing"))
                                                    float.TryParse(reader.GetValue(_columnOrder["linespacing"]).ToString(), NumberStyles.Any, CultureInfo.InvariantCulture,  out _lineSpacing);

                                                string _text = "Example Text.";
                                                if (_columnOrder.ContainsKey("text"))
                                                    _text = reader.GetValue(_columnOrder["text"]).ToString();

                                                if (_text.Contains("…"))
                                                    _text = _text.Replace("…", "...");

                                                if (_text.Contains(@"\n"))
                                                {
                                                    _text = _text.Replace(@"\n", "\r\n");
                                                }

                                                string _font = "TW Cen MT";
                                                if (_columnOrder.ContainsKey("font"))
                                                    _font = reader.GetValue(_columnOrder["font"]).ToString();

                                                string _style = "regular";
                                                if (_columnOrder.ContainsKey("style"))
                                                    _style = reader.GetValue(_columnOrder["style"]).ToString();

                                                string _file = "test.jpg";
                                                if (_columnOrder.ContainsKey("file"))
                                                    _file = reader.GetString(_columnOrder["file"]).ToString();

                                                string _foreground = "0;0;0";
                                                if (_columnOrder.ContainsKey("foreground"))
                                                    _foreground = reader.GetValue(_columnOrder["foreground"]).ToString();

                                                string _bordercolor = "0;0;0";
                                                if (_columnOrder.ContainsKey("bordercolor"))
                                                    _bordercolor = reader.GetValue(_columnOrder["bordercolor"]).ToString();

                                                string _background = "255;255;255";
                                                if (_columnOrder.ContainsKey("background"))
                                                    _background = reader.GetValue(_columnOrder["background"]).ToString();

                                                float _borderwidth = (float)1;
                                                if (_columnOrder.ContainsKey("borderwidth"))
                                                    float.TryParse(reader.GetValue(_columnOrder["borderwidth"]).ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out _borderwidth);

                                                string _textalignment = "left";
                                                if (_columnOrder.ContainsKey("textalignment"))
                                                    _textalignment = reader.GetValue(_columnOrder["textalignment"]).ToString();

                                                if (useBlindText)
                                                    _text = blindTextTemplate.Substring(0, _text.Length);

                                                GenerateImage(_width, _height, _left, _right, _top, _fontsize, _lineSpacing, _text, _font, _style, _file, _foreground, _bordercolor, _background, _borderwidth, _textalignment);

                                            }
                                            catch (Exception _ex)
                                            {
                                                Console.WriteLine(_ex);
                                            }
                                        } 
                                        else
                                        {
                                            for (int i = 0; i < reader.FieldCount; i++)
                                            {
                                                string _key = reader.GetString(i).ToLower();
                                                if (_columnOrder.ContainsKey(_key))
                                                    _columnOrder[_key] = i;
                                                else
                                                {
                                                    _columnOrder.Add(_key, i);
                                                }
                                            }
                                            isHeader = false;
                                        }                                        
                                    }
                                } while (reader.NextResult());
                            }
                        } 
                    }
                    else
                    {
                        Console.WriteLine(args[0] + " not found");
                    }
                }
                catch (Exception _ex)
                {
                    Console.WriteLine(_ex);
                }
            }           
        }

        private static void GenerateImage(int _width, int _height, int _left, int _right, int _top, float _fontsize, float _lineSpacing, string _text, string _font, string _style, string _file, string _foreground, string _bordercolor, string _background, float _borderwidth, string _textalignment)
        {
            using (var image = new Image<Rgba32>(_width, _height))
            {

                var _createdFont = SystemFonts.CreateFont(_font, _fontsize, FontStyle.Regular);
                if (_style.ToLower().Trim() == "bold")
                    _createdFont = SystemFonts.CreateFont(_font, _fontsize, FontStyle.Bold);
                else if (_style.ToLower().Trim() == "italic")
                    _createdFont = SystemFonts.CreateFont(_font, _fontsize, FontStyle.Italic);
                else if (_style.ToLower().Trim() == "regular")
                    _createdFont = SystemFonts.CreateFont(_font, _fontsize, FontStyle.Regular);
                else if (_style.ToLower().Trim() == "bolditalic")
                    _createdFont = SystemFonts.CreateFont(_font, _fontsize, FontStyle.BoldItalic);

                var _textOptions = new TextOptions
                {
                    ApplyKerning = true,
                    VerticalAlignment = VerticalAlignment.Top,
                    HorizontalAlignment = SixLabors.Fonts.HorizontalAlignment.Left,
                    WrapTextWidth = image.Width - _left - _right,
                    LineSpacing = _lineSpacing
                };

                if (_textalignment.ToLower().Trim() == "right")
                    _textOptions.HorizontalAlignment = SixLabors.Fonts.HorizontalAlignment.Right;
                else if (_textalignment.ToLower().Trim() == "center")
                    _textOptions.HorizontalAlignment = SixLabors.Fonts.HorizontalAlignment.Center;

                if (_borderwidth > 0)
                {
                    var _points = new PointF[5];
                    _points[0] = new PointF { X = _borderwidth / 2, Y = _borderwidth / 2 };
                    _points[1] = new PointF { X = (float)image.Width - _borderwidth, Y = _borderwidth / 2 };
                    _points[2] = new PointF { X = (float)image.Width - _borderwidth, Y = (float)image.Height - _borderwidth };
                    _points[3] = new PointF { X = _borderwidth / 2, Y = (float)image.Height - _borderwidth };
                    _points[4] = new PointF { X = _borderwidth / 2, Y = _borderwidth / 2 };

                    var _linePen = new Pen(GetColor(_bordercolor), _borderwidth);
                    image.Mutate(ctx => ctx 
                         .Fill(GetBrush(_background, _height, _width))
                        .DrawLines(_linePen, _points)
                        .DrawText(new DrawingOptions() { TextOptions = _textOptions }, _text, _createdFont, GetColor(_foreground), new PointF(_left, _top))
                       );
                }
                else
                { 
                    image.Mutate(ctx => ctx 
                       .Fill(GetBrush(_background, _height, _width))
                       .DrawText(new DrawingOptions() { TextOptions = _textOptions }, _text, _createdFont, GetColor(_foreground), new PointF(_left, _top)));
                }

                bool _success = false;
                if (_file.ToLower().EndsWith(".png"))
                {
                    image.SaveAsPng(_file);
                    _success = true;
                }
                else if (_file.ToLower().EndsWith(".jpg") || _file.ToLower().EndsWith(".jepg"))
                {
                    image.SaveAsJpeg(_file);
                    _success = true;
                }
                else if (_file.ToLower().EndsWith(".gif"))
                {
                    image.SaveAsGif(_file);
                    _success = true;
                }
                else if (_file.ToLower().EndsWith(".bmp"))
                {
                    image.SaveAsBmp(_file);
                    _success = true;
                }

                if (_success)
                {
                    Console.WriteLine(" - " + _file + " generated (" + _text + ")");
                }
            }  
        }

        private static Rgba32 GetColor(string colorRGB)
        {
            int _R = 0;
            int _G = 0;
            int _B = 0;
            int _A = 0;
            string[] _colorComp = colorRGB.Split(";");
            if (_colorComp.Length != 3 && _colorComp.Length != 4)
            {
                return new Rgba32(_R, _G, _B);
            } 
            else
            {
                int.TryParse(_colorComp[0], out _R);
                int.TryParse(_colorComp[1], out _G);
                int.TryParse(_colorComp[2], out _B);

                if (_colorComp.Length == 3)
                { 
                    return new Rgba32((float)_R/255, (float) _G / 255, (float)_B / 255);
                } 
                else
                {
                    int.TryParse(_colorComp[3], out _A);
                    return new Rgba32((float)_R / 255, (float)_G / 255, (float)_B / 255, (float)_A / 255);
                }
            }
        }

        private static LinearGradientBrush GetBrush(string colorRGB, int height, int width)
        {
            
            var linearGradientBrush = new LinearGradientBrush(new Point(0, 0), new Point(0, height), GradientRepetitionMode.Repeat,
                                            new ColorStop(0, Color.White), new ColorStop(1, Color.White));

            string[] _multipleColors = colorRGB.Split("|");
             
            string[] _colorComp1 = _multipleColors[0].Split(";");
            if (_colorComp1.Length != 3 && _colorComp1.Length != 4)
            {
                return linearGradientBrush;
            }
            else
            {
                ColorStop c1 = new ColorStop();
                ColorStop c2 = new ColorStop();

                byte _R1 = 0;
                byte _G1 = 0;
                byte _B1 = 0;
                byte _A1 = 0;

                byte.TryParse(_colorComp1[0], out _R1);
                byte.TryParse(_colorComp1[1], out _G1);
                byte.TryParse(_colorComp1[2], out _B1);

                if (_colorComp1.Length == 3)
                {
                    c1 = new ColorStop(0, Color.FromRgb(_R1, _G1, _B1));
                    c2 = new ColorStop(1, Color.FromRgb(_R1, _G1, _B1));
                }
                else
                {
                    byte.TryParse(_colorComp1[3], out _A1);
                    c1 = new ColorStop(0, Color.FromRgba(_R1, _G1, _B1, _A1));
                    c2 = new ColorStop(1, Color.FromRgba(_R1, _G1, _B1, _A1));
                }

                if (_multipleColors.Length == 2)
                {
                    string[] _colorComp2 = _multipleColors[1].Split(";");

                    byte _R2 = 0;
                    byte _G2 = 0;
                    byte _B2 = 0;
                    byte _A2 = 0;

                    byte.TryParse(_colorComp2[0], out _R2);
                    byte.TryParse(_colorComp2[1], out _G2);
                    byte.TryParse(_colorComp2[2], out _B2);

                    if (_colorComp2.Length == 3)
                    {
                        c2 = new ColorStop(1, Color.FromRgb(_R2, _G2, _B2));
                    }
                    else
                    {
                        byte.TryParse(_colorComp2[3], out _A2);
                        c2 = new ColorStop(1, Color.FromRgba(_R2, _G2, _B2, _A2));
                    }

                }
                  
                return new LinearGradientBrush(new Point(0, 0), new Point(0, height), GradientRepetitionMode.Repeat, c1, c2);

            }
        }

    } 
}
