using System.Text.RegularExpressions;

namespace ChatFormatterPro
{
    public static class MathLatexHelper
    {
        // Converts simple exponent text into LaTeX-style math
        // Example: (3^4)^2  ->  (3^{4})^{2}
        public static string ConvertPowersToLatex(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return input;

            // a^b  →  a^{b}
            input = Regex.Replace(
                input,
                @"(\w+)\^(\w+)",
                "$1^{$2}"
            );

            return input;
        }
    }
}
