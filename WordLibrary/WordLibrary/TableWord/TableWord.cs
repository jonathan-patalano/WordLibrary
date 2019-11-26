using System;
using System.Linq;
using System.Reflection;
using WordLibrary.Bookmark;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WordLibrary.TableWord
{
    public static class TableWord
    {
        #region Private Members


        #endregion

        #region Constructors



        #endregion

        #region Méthodes publiques

        /// <summary>
        /// Charge un document, récupère un tableau et remplace la ligne par défaut par des copies mises à jour.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="legende"></param>
        public static void CreateRowsFromTemplate(DocX document, string legende)
        {
            // obtenir le tableau avec la légende du document.
            var ListTable = document.Tables.FirstOrDefault(t => t.TableCaption == legende);
            if (ListTable == null)
            {
                throw new Exception("Tableau est null");
            }
            else
            {
                for (int i = 1; ListTable.RowCount > i; i++)
                {
                    // Obtenir le motif de ligne de la deuxième rangée.
                    var rowPattern = ListTable.Rows[i];

                    // Ajouter des données
                    string aze = "";
                    FieldInfo[] fields = typeof(Constantes).GetFields();
                    foreach (FieldInfo field in fields)
                    {
                        AddItemToTable(ListTable, rowPattern, field.GetValue(field).ToString(), aze);
                    }


                    // Supprimer la ligne de motif.
                    rowPattern.Remove();
                }
            }
        }

        /// <summary>
        /// Créer un tableau, insérer des lignes, une image et remplacer du texte.
        /// </summary>
        //public static void InsertRowAndImageTable(DocX document)
        //{
        //    // Add a Table into the document and sets its values.
        //    var t = document.AddTable(5, 2);
        //    t.Design = TableDesign.ColorfulListAccent1;
        //    t.Alignment = Alignment.center;
        //    t.Rows[0].Cells[0].Paragraphs[0].Append("Mike");
        //    t.Rows[0].Cells[1].Paragraphs[0].Append("65");
        //    t.Rows[1].Cells[0].Paragraphs[0].Append("Kevin");
        //    t.Rows[1].Cells[1].Paragraphs[0].Append("62");
        //    t.Rows[2].Cells[0].Paragraphs[0].Append("Carl");
        //    t.Rows[2].Cells[1].Paragraphs[0].Append("60");
        //    t.Rows[3].Cells[0].Paragraphs[0].Append("Michael");
        //    t.Rows[3].Cells[1].Paragraphs[0].Append("59");
        //    t.Rows[4].Cells[0].Paragraphs[0].Append("Shawn");
        //    t.Rows[4].Cells[1].Paragraphs[0].Append("57");

        //    // Add a row at the end of the table and sets its values.
        //    var r = t.InsertRow();
        //    r.Cells[0].Paragraphs[0].Append("Mario");
        //    r.Cells[1].Paragraphs[0].Append("54");

        //    // Add a row at the end of the table which is a copy of another row, and sets its values.
        //    var newPlayer = t.InsertRow(t.Rows[2]);
        //    newPlayer.ReplaceText("Carl", "Max");
        //    newPlayer.ReplaceText("60", "50");

        //    // Add an image into the document.    
        //    var image = document.AddImage(Table.TableSampleResourcesDirectory + @"logo_xceed.png");
        //    // Create a picture from image.
        //    var picture = image.CreatePicture(25, 100);

        //    // Calculate totals points from second column in table.
        //    var totalPts = 0;
        //    foreach (var row in t.Rows)
        //    {
        //        totalPts += int.Parse(row.Cells[1].Paragraphs[0].Text);
        //    }

        //    // Add a row at the end of the table and sets its values.
        //    var totalRow = t.InsertRow();
        //    totalRow.Cells[0].Paragraphs[0].Append("Total for ").AppendPicture(picture);
        //    totalRow.Cells[1].Paragraphs[0].Append(totalPts.ToString());
        //    totalRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;

        //    // Insert a new Paragraph into the document.
        //    var p = document.InsertParagraph("Xceed Top Players Points:");
        //    p.SpacingAfter(40d);

        //    // Insert the Table after the Paragraph.
        //    p.InsertTableAfterSelf(t);

        //    document.Save();
        //}

        #endregion

        #region Méthodes privées

        /// <summary>
        /// 
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rowPattern"></param>
        /// <param name="textecible"></param>
        /// <param name="texteremplacement"></param>
        private static void AddItemToTable(Table table, Row rowPattern, string textecible, string texteremplacement)
        {
            // Insert a copy of the rowPattern at the last index in the table.
            var newItem = table.InsertRow(rowPattern, table.RowCount - 1);

            // Replace the default values of the newly inserted row.
            newItem.ReplaceText(textecible, texteremplacement);
        }

        #endregion
    }
}
