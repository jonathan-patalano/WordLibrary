using System.Collections.Generic;
using System.Text.RegularExpressions;
using Xceed.Words.NET;

namespace WordLibrary
{
    public static class DocumentWord
    {

        #region Constructors

        #endregion

        private static Dictionary<string, string> _replacePatterns;

        /// <summary>
        /// Remplacer le texte qui se trouve entre <> dans un document.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="valeur"></param>
        public static void ReplaceText(DocX document, IDictionary<string, string> valeur)
        {
            // Vérifiez si certains des modèles de remplacement sont utilisés dans le document chargé.
            if (document.FindUniqueByPattern(@"<[\w \=]{4,}>", RegexOptions.IgnoreCase).Count > 0)
            {
                _replacePatterns = (Dictionary<string, string>)valeur;
                // Effectue le remplacement de tous les tag trouvé.
                document.ReplaceText("<(.*?)>", DocumentWord.ReplaceFunc, false, RegexOptions.IgnoreCase);
            }
        }

        #region Private Methods

        /// <summary>
        /// Fonction de remplacement de string.
        /// </summary>
        /// <param name="findStr"></param>
        /// <returns></returns>
        private static string ReplaceFunc(string findStr)
        {
            if (null != _replacePatterns && _replacePatterns.ContainsKey(findStr))
            {
                return _replacePatterns[findStr];
            }
            return findStr;
        }

        #endregion
    }
}
