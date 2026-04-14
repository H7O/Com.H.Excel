using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;

namespace Com.H.Excel
{
    internal static class TextExtensions
    {
        public static bool EqualsIgnoreCase(
            this string originalString,
            string stringToCompare)
            =>
            originalString?.IsNullEqual(stringToCompare) ??
                // IsNullEquals ensures no scenario would result in originalString and/or stringToCompare are null
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                originalString
                .ToUpper(CultureInfo.InvariantCulture)
                .Equals(stringToCompare.ToUpper(CultureInfo.InvariantCulture));

        private static bool? IsNullEqual(
            this string originalString,
            string stringToCompare)
        {
            if (originalString == null && stringToCompare == null) return true;
            if ((originalString != null && stringToCompare == null)
                ||
                (originalString == null && stringToCompare != null)
                ) return false;
            return null;
        }

        public static string ExtractAlphabet(this string text)
            => Regex.Matches(text, "[a-z]+", RegexOptions.IgnoreCase)
            .Cast<Match>()
                .Aggregate("", (i, n) => i + n.Value);



    }
}
