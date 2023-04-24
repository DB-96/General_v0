using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

public static class Program {
  static void Main() {
    using (StreamReader reader = new StreamReader(Console.OpenStandardInput()))
    while (!reader.EndOfStream) {
      string line = reader.ReadLine();
      Console.WriteLine(HSVToRGB(line));
    }
  }
////////////////////////////////////////////////////////////////////////////////////
// Legacy // Alternate timeline // Not used
  static string[] NumberExtraction(string input){
      string pattern = @"\d{3}";
      MatchCollection matches = Regex.Matches(input, pattern);
      string[] numbers = new string[matches.Count];
      for (int i = 0; i < matches.Count; i++) {
            numbers[i] = matches[i].Value;
        }
        return numbers;
    }
//////////////////////////////////////////////////////////////////////////////////////
   static string HSVToRGB(string input)
   /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   //  Input ::          input {string} - 3xY String Array (comma delimited HSV triplet)
   //  Fn Description :: Converts HSV triplet to RGB triplet using provided algorithm
   //  File ::
   //  Date ::
   /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
{
    string[] columns = input.Split(',');  //Split into h, s, v columns based on comma delimiter
    int h;              // variable declaration
    double s, v, c, x, m, red, green, blue;       

    int.TryParse(columns[0], out h);      // Convert 1st column {H} -> Int
    double.TryParse(columns[1], out s);   // Convert 2nd column {S} -> double
    double.TryParse(columns[2], out v);   // Convert 3rd column {V} -> double
  
    s /= 100.0; //input provided as % - rectify
    v /= 100.0; //input provided as % - rectify 

    c = v * s;
    x = c * (1 - Math.Abs((h / 60.0) % 2 - 1));
    m = v - c;

    switch ((int)h / 60)
    {
        case 0: // 0 <= H < 60
            red = c;
            green = x;
            blue = 0 ;
            break;
        case 1: // 60 <= H < 120
            red = x ;
            green = c;
            blue = 0;
            break;
        case 2: // 120 <= H < 180
            red = 0;
            green = c;
            blue = x;
            break;
        case 3: // 180 <= H < 240
            red =  0;
            green = x;
            blue = c;
            break;
        case 4: // 240 <= H < 300
            red = x;
            green = 0;
            blue = c;
            break;
        default: // 300 <= H < 360
            red = c;
            green = 0;
            blue = x;
            break;
    }
            green = (green + m) * 255.0;
            blue = (blue + m) * 255.0;
            red = (red + m) * 255.0;

            // Round the RGB values to the nearest integer
            int r = (int)Math.Round(red);
            int g = (int)Math.Round(green);
            int b = (int)Math.Round(blue);

            // Format the output as a comma-delimited string
    return $"{r:D3},{g:D3},{b:D3}";// return as 3 decimal digits (including padding zeros if applicable)
}
}
