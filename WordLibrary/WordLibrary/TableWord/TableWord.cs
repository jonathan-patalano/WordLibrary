using CG13.Infrastructure.CrossCutting.Template;
using System;
using System.Collections.Generic;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WordLibrary
{
    public static class TableWord
    {
        #region Private Members

        private static Row Header;
        private static List<Row> ContentCollection;

        #endregion

        #region Constructors



        #endregion

        #region Méthodes publiques

        public static void ReplaceText(ITemplate template, DocX document, IDictionary<string, string> values)
        {
            List<string> fields = document.FindUniqueByPattern(@"\#{var\.([^}]+)}", System.Text.RegularExpressions.RegexOptions.Singleline);
            if (fields != null && fields.Any())
            {
                foreach (string field in fields)
                {
                    string key = field.Substring(6, field.Length - 7);
                    if (values.ContainsKey(key))
                    {
                        document.ReplaceText(field, getStringValue(template, key, values[key]));
                    }
                    else
                    {
                        document.ReplaceText(field, getStringValue(template, key, ""));
                    }
                }
            }
        }

        public static string getStringValue(ITemplate template, string key, object value)
        {
            IFormatProvider formatter = template.getFormatter(key);
            return String.Format(formatter, "{0}", value);

        }

        public static void AddTableHeader(DocX document)
        {
            Table table = document.AddTable(0, 0);
            document.InsertTable(table);
        }

        public static void CopyHeaderTable(DocX document)
        {
            Header = document.Tables[0].Rows[4];
            document.Tables[0].Rows[4].Remove();
        }

        public static void PasteHeaderTable(DocX document)
        {
            if (null != Header)
            {
                document.Tables[0].InsertRow(Header, true);
            }
        }

        public static void CopycontentTable(DocX document)
        {
            ContentCollection = new List<Row>();
            for (int i = 4; i < 9; i++)
            {
                ContentCollection.Add(document.Tables[0].Rows[i]);
            }
            var max = document.Tables[0].Rows.Count;
            for (int i = max; i >= 5; i--)
            {
                document.Tables[0].Rows[i - 1].Remove();
            }

        }

        public static void PastecontentTable(DocX document)
        {
            foreach (Row content in ContentCollection)
            {
                document.Tables[0].InsertRow(content, true);
            }
        }


        public static void RemoveTemplateDuplicated(DocX document)
        {
            for (int i = 1; i <= 6; i++)
            {
                document.Tables[0].Rows[4].Remove();
            }
        }

        #endregion

        #region Méthodes privées


        #endregion
    }
}
