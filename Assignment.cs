// using System;
// using System.IO;
// using System.Collections.Generic;
// using System.Text;
// using System.Text.RegularExpressions;

// public static class Program {
//   static void Main() {
//     using (StreamReader reader = new StreamReader(Console.OpenStandardInput()))
//     while (!reader.EndOfStream) {
//       string line = reader.ReadLine();
//       Console.WriteLine(HSVToRGB(line));
//     }
    
//     static string[] NumberExtraction(string input){
//           // input is provided as comma delimited HSV triplets
//           // Output(1)= Hue (H) : : Degrees
//           // Output(2)= Saturation : : %
//           // Output(3)= Value : : %
//           string pattern = @"\d{3}"; //looking for 3 consecutive digits
//           MatchCollection matches = Regex.Matches(input, pattern);
//           string[] numbers = new string[matches.Count]; //match string array size with number of regexp pattern matches
//           for (int i = 0; i < matches.Count; i++) {
//                 numbers[i] = matches[i].Value;
//             }
//             //Console.WriteLine(numbers)
//             return numbers; //return an array of type string - ideally should by the same input without columns
//         }

//        static string HSVToRGB(string input){
//           //////////////////////////////////////////////////////////////
//           //
//           //
//           //
//           //////////////////////////////////////////////////////////////
//           // C = (S/100)*(V/100)
//           // X = C*(1- |(H/60)mod2-1|)
//           // m = (V/100) - C
//           //  string[] Extract = NumberExtraction(input);
 
//             // Split the row into an array of string values
//             string[] Extract = input.Split(',');
//             int h = Convert.ToInt32(Extract[0]); // column 1
//             int s = Convert.ToInt32(Extract[1]); // column 2
//             int v = Convert.ToInt32(Extract[2]); // column 3
//           // Convert the HSV values to RGB values
//           int r, g, b;
//           if (s == 0) { // Boundary case aka achromatic
//               r = g = b = v;
//           } else {
//               double h /= 60.0;
//               int i = (int)Math.Floor(h); // round to nearest integer
//               double f = h - i;
//               int p = v * (1 - s);
//               int q = v * (1 - (s * f));
//               int t = v * (1 - (s * (1 - f)));
//               switch (i) {
//                   case 0: r = v; g = t; b = p; break;
//                   case 1: r = q; g = v; b = p; break;
//                   case 2: r = p; g = v; b = t; break;
//                   case 3: r = p; g = q; b = v; break;
//                   case 4: r = t; g = p; b = v; break;
//                   default: r = v; g = p; b = q; break;
//               }
//           }
//             return $"{r},{g},{b}";
//         }
//   }

//   public static void HSVToRGB(double h, double s, double v, out int r, out int g, out int b) {
//     if (s == 0) {
//         // achromatic (gray)
//         r = g = b = (int)(v * 255);
//         return;
//     }

//     h /= 60;            // sector 0 to 5
//     int i = (int)Math.Floor(h); // round to *down* nearest integer
//     double f = h - i;          // factorial part of h
//     double p = v * (1 - s);
//     double q = v * (1 - s * f);
//     double t = v * (1 - s * (1 - f));

//     switch (i) {
//         case 0:
//             r = (int)(v * 255);
//             g = (int)(t * 255);
//             b = (int)(p * 255);
//             break;
//         case 1:
//             r = (int)(q * 255);
//             g = (int)(v * 255);
//             b = (int)(p * 255);
//             break;
//         case 2:
//             r = (int)(p * 255);
//             g = (int)(v * 255);
//             b = (int)(t * 255);
//             break;
//         case 3:
//             r = (int)(p * 255);
//             g = (int)(q * 255);
//             b = (int)(v * 255);
//             break;
//         case 4:
//             r = (int)(t * 255);
//             g = (int)(p * 255);
//             b = (int)(v * 255);
//             break;
//         default:        // case 5:
//             r = (int)(v * 255);
//             g = (int)(p * 255);
//             b = (int)(q * 255);
//             break;
//     }
// }

// }


static string HSVToRGB(string input)
{
    // Extract the hue, saturation, and value from the input
    string[] parts = input.Split(',');
    int h = int.Parse(parts[0]);                // Convert 1st column {H} -> Int
    double s = double.Parse(parts[1]) / 100.0;  // Convert 2nd column {S} -> double
    double v = double.Parse(parts[2]) / 100.0;  // Convert 3rd column {V} -> double

    // Convert HSV to RGB
    double c = v * s;
    double x = c * (1 - Math.Abs((h / 60.0) % 2 - 1));
    double m = v - c;

    double r, g, b;
    if (h < 60)
    {
        r = c;
        g = x;
        b = 0;
    }
    else if (h < 120)
    {
        r = x;
        g = c;
        b = 0;
    }
    else if (h < 180)
    {
        r = 0;
        g = c;
        b = x;
    }
    else if (h < 240)
    {
        r = 0;
        g = x;
        b = c;
    }
    else if (h < 300)
    {
        r = x;
        g = 0;
        b = c;
    }
    else
    {
        r = c;
        g = 0;
        b = x;
    }

    r = (r + m) * 255.0 + 0.5;
    g = (g + m) * 255.0 + 0.5;
    b = (b + m) * 255.0 + 0.5;

    // Round the RGB values to the nearest integer
    int red = (int)Math.Round(r);
    int green = (int)Math.Round(g);
    int blue = (int)Math.Round(b);

    // Format the output as a comma-delimited string
    return $"{red:D3},{green:D3},{blue:D3}";
}