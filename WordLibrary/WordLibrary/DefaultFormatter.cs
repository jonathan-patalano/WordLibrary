using System;
using System.Globalization;

namespace WordLibrary
{
    class DefaultFormatter : ICustomFormatter, IFormatProvider
    {

        // IFormatProvider.GetFormat implementation. 
        public object GetFormat(Type formatType)
        {
            // Determine whether custom formatting object is requested. 
            if (formatType == typeof(ICustomFormatter))
                return this;
            else
                return null;
        }

        public string Format(string format, object arg, IFormatProvider formatProvider)
        {
            if (arg != null)
            {
                if (arg is DateTime)
                {
                    return ((DateTime)arg).ToString("d");
                }
                else if (arg is bool)
                {
                    bool value = (bool)arg;
                    if (value)
                        return @"☒ ";
                    else
                        return @"☐ ";

                }
            }

            return HandleOtherFormats(format, arg);
        }


        private string HandleOtherFormats(string format, object arg)
        {
            if (arg is IFormattable)
                return ((IFormattable)arg).ToString(format, CultureInfo.GetCultureInfo("fr-FR"));
            else if (arg != null)
                return arg.ToString();
            else
                return String.Empty;
        }


    }
}
