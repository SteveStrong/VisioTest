using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace VisioShapeExtractor
{
    /// <summary>
    /// Parses Visio formulas like "Width*0.5" into numerical values
    /// </summary>
    public static class FormulaParser
    {
        private static readonly Regex WidthMultiplier = new Regex(@"Width\s*\*\s*(\d+(\.\d+)?)", RegexOptions.IgnoreCase);
        private static readonly Regex HeightMultiplier = new Regex(@"Height\s*\*\s*(\d+(\.\d+)?)", RegexOptions.IgnoreCase);
        
        /// <summary>
        /// Parse a formula that may contain references to Width or Height
        /// </summary>
        /// <param name="formula">The formula string (e.g. "Width*0.5")</param>
        /// <param name="width">Current width value if needed for calculation</param>
        /// <param name="height">Current height value if needed for calculation</param>
        /// <returns>Calculated value or 0 if parsing fails</returns>
        public static double ParseFormula(string formula, double width = 1.0, double height = 1.0)
        {
            if (string.IsNullOrEmpty(formula))
                return 0;
                
            // First check for direct numeric value
            if (double.TryParse(formula, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
                return value;
                
            // Check for Width*multiplier pattern
            var widthMatch = WidthMultiplier.Match(formula);
            if (widthMatch.Success && widthMatch.Groups.Count >= 2)
            {
                if (double.TryParse(widthMatch.Groups[1].Value, 
                    NumberStyles.Any, CultureInfo.InvariantCulture, out double multiplier))
                {
                    return width * multiplier;
                }
            }
            
            // Check for Height*multiplier pattern
            var heightMatch = HeightMultiplier.Match(formula);
            if (heightMatch.Success && heightMatch.Groups.Count >= 2)
            {
                if (double.TryParse(heightMatch.Groups[1].Value, 
                    NumberStyles.Any, CultureInfo.InvariantCulture, out double multiplier))
                {
                    return height * multiplier;
                }
            }
            
            // If we couldn't parse the formula, return 0
            return 0;
        }
    }
}
