using System;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WordLibrary.HyperlinkWord
{
    public class HyperlinkWord
    {
        #region Public Methods
        /// <summary>
        /// Ajouter un lien Hypertexte dans un document.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="word"></param>
        /// <param name="linkword"></param>
        /// <param name="sentence"></param>
        public static void Hyperlinks(DocX document, string word, string linkword, string sentence)
        {
            // Ajout d'un lien dans le document.
            Hyperlink hyperlink = document.AddHyperlink(word, new Uri(linkword));
            // Ajout d'un paragraphe
            Paragraph paragraph = document.InsertParagraph();
            paragraph.Append(sentence);
            // Insérer un hyperlien à un index spécifique dans le présent paragraphe.
            paragraph.AppendHyperlink(hyperlink);
        }

        #endregion
    }
}
