using System.Collections.Generic;
using System.IO;
using System.Linq;
using WordLibrary.Abstract;
using WordLibrary.Interface;
using WordLibrary.Table;
using Xceed.Words.NET;
using XTable = Xceed.Words.NET.Table;
namespace WordLibrary
{
    public class SimpleWordTemplateEngine : AbstractTemplateEngine, ITemplateBuilder
    {
        public SimpleWordTemplateEngine()
            : base("DOCX")
        {
        }

        protected override bool checkOutputFormat(TemplateServiceOutputFormat output)
        {
            return TemplateServiceOutputFormat.Docx.Equals(output);
        }

        public override void renderDocument(ITemplate template, IDictionary<string, object> values, ref Stream outputStream)
        {
            using (Stream st = template.getContentAsStream())
            {
                using (DocX templateDoc = DocX.Load(st))
                {
                    TraitementTableau(templateDoc, values, template);
                    TraitementTexte(templateDoc, values, template);
                    TraitementImage(templateDoc, values);
                    templateDoc.SaveAs(outputStream);
                }
            }

        }

        public bool canBuild(string type)
        {
            return (type != null && sType.Equals(type));
        }

        public ITemplate createTemplate(string name, string templateContent)
        {
            return new DefaultTemplate(name, templateContent, sType);
        }

        public ITemplate createTemplateFromFile(string name, string path)
        {
            if (File.Exists(path))
            {
                return new FileTemplate(name, path, sType);
            }

            return null;
        }

        /// <summary>
        /// Traitement des tableaux
        /// </summary>
        /// <param name="templateDoc"></param>
        /// <param name="values"></param>
        /// <param name="template"></param>
        private void TraitementTableau(DocX templateDoc, IDictionary<string, object> values, ITemplate template)
        {
            List<XTable> listeTab = templateDoc.Tables;
            foreach (XTable tab in listeTab)
            {
                if (tab.Rows[0].Cells[0].Paragraphs[0].Text.StartsWith("#{tab.")
                    && values.ContainsKey(tab.Rows[0].Cells[0].Paragraphs[0].Text.Substring(6, tab.Rows[0].Cells[0].Paragraphs[0].Text.Length - 7)))
                {
                    TableDocX tableAInserer = (TableDocX)values[tab.Rows[0].Cells[0].Paragraphs[0].Text.Substring(6, tab.Rows[0].Cells[0].Paragraphs[0].Text.Length - 7)];
                    int nbCol = tableAInserer.ListeEntete.Count;
                    int nbLig = tableAInserer.ListeCellule.Max(x => x.NumeroLigne) + 1;
                    XTable table = templateDoc.AddTable(nbLig, nbCol);
                    table.Alignment = Alignment.center;
                    table.Design = TableDesign.Custom;
                    table.Rows[0].TableHeader = true;
                    foreach (TableDocXEntete entete in tableAInserer.ListeEntete)
                    {
                        table.Rows[0].Cells[entete.NumeroColonne].Paragraphs[0].InsertText(this.getStringValue(template, string.Empty, entete.ValeurEntete));
                        table.Rows[0].Cells[entete.NumeroColonne].Paragraphs[0].Color(System.Drawing.Color.White).Font("Arial").FontSize(12).Alignment = Alignment.center;
                        table.Rows[0].Cells[entete.NumeroColonne].FillColor = System.Drawing.Color.RoyalBlue;
                    }
                    foreach (TableDocXCellule cell in tableAInserer.ListeCellule)
                    {
                        table.Rows[cell.NumeroLigne].Cells[cell.NumeroColonne].Paragraphs[0].InsertText(this.getStringValue(template, string.Empty, cell.ValeurCellule));
                        table.Rows[cell.NumeroLigne].Cells[cell.NumeroColonne].Paragraphs[0].Font("Arial").FontSize(12);
                    }
                    tab.InsertTableAfterSelf(table);
                    tab.Remove();
                }
            }
        }

        /// <summary>
        /// Traitement des différents champs
        /// </summary>
        /// <param name="templateDoc"></param>
        /// <param name="values"></param>
        /// <param name="template"></param>
        private void TraitementTexte(DocX templateDoc, IDictionary<string, object> values, ITemplate template)
        {
            List<string> fields = templateDoc.FindUniqueByPattern(@"\#{var\.([^}]+)}", System.Text.RegularExpressions.RegexOptions.Singleline);
            if (fields != null && fields.Any())
            {
                foreach (string field in fields)
                {
                    string key = field.Substring(6, field.Length - 7);
                    if (values.ContainsKey(key))
                    {
                        templateDoc.ReplaceText(field, getStringValue(template, key, values[key]));

                    }
                    else
                    {
                        templateDoc.ReplaceText(field, getStringValue(template, key, ""));

                    }
                }

            }
        }

        /// <summary>
        /// Traitement des images
        /// </summary>
        /// <param name="templateDoc"></param>
        /// <param name="values"></param>
        private void TraitementImage(DocX templateDoc, IDictionary<string, object> values)
        {
            foreach (var paragraph in templateDoc.Paragraphs)
            {
                if (paragraph.Text.StartsWith("#{img.") && paragraph.Text.EndsWith("}"))
                {
                    string key = paragraph.Text.Substring(6, paragraph.Text.Length - 7);
                    if (values.ContainsKey(key))
                    {
                        paragraph.RemoveText(0, paragraph.Text.Count());
                        Picture picture = GetPicture(templateDoc, values[key] as byte[]);
                        paragraph.AppendPicture(picture);
                        paragraph.Alignment = Alignment.center;
                    }
                }
            }
        }
    }
}
