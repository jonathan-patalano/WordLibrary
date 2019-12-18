using Xceed.Words.NET;

namespace WordLibrary.MarginsWord
{
    public class MarginsWord
    {
        /// <summary>
        /// Ajout de marges sur un document
        /// </summary>
        /// <param name="document"></param>
        public static void Margins(DocX document)
        {
            // Définisr la largeur de la page pour qu'elle soit plus petite.
            document.PageWidth = 350f;

            // Définissez les marges du document.
            document.MarginLeft = 85f;
            document.MarginRight = 85f;
            document.MarginTop = 0f;
            document.MarginBottom = 50f;
        }
    }
}
