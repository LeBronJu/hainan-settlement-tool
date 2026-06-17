using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace HainanSettlementTool.Core.Services
{
    public static class TextUtil
    {
        public static string S(object value)
        {
            return value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture).Trim();
        }

        public static double N(object value)
        {
            if (value == null)
            {
                return 0d;
            }

            if (value is double)
            {
                return (double)value;
            }

            if (value is float)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }

            if (value is decimal)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }

            if (value is int || value is long || value is short)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }

            var text = S(value).Replace(",", string.Empty);
            if (text.Length == 0 || text.StartsWith("=", StringComparison.Ordinal))
            {
                return 0d;
            }

            double parsed;
            return double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed) ? parsed : 0d;
        }

        public static string CustomerKey(object value)
        {
            return Regex.Replace(S(value), "\\s+", string.Empty);
        }

        public static string SafeFileName(string value)
        {
            var invalid = new string(System.IO.Path.GetInvalidFileNameChars());
            return Regex.Replace(S(value), "[" + Regex.Escape(invalid) + "]", "_");
        }
    }
}
