using System.Collections.Generic;
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WordLibrary.DocumentWord
{
    public static class DocumentWord
    {
        /// <summary>
        /// Remplacer le texte qui se trouve entre <> dans un document.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="valeur"></param>
        public static void ReplaceText(DocX document, Dictionary<string, string> valeur)
        {
            // Vérifiez si tous les modèles de remplacement sont utilisés dans le document chargé.
            if (document.FindUniqueByPattern(@"<[\w \=]{4,}>", RegexOptions.IgnoreCase).Count == valeur.Count)
            {
                // Effectuer le remplacement.
                for (int i = 0; i < valeur.Count; ++i)
                {
                    document.ReplaceText("<(.*?)>", ReplaceFunc(null, valeur), false, RegexOptions.IgnoreCase, null, new Formatting());
                }
            }

        }

        #region Private Methods
        /// <summary>
        /// 
        /// </summary>
        /// <param name="findStr"></param>
        /// <param name="valeur"></param>
        /// <returns></returns>
        private static string ReplaceFunc(string findStr, Dictionary<string, string> valeur)
        {
            if (valeur.ContainsKey(findStr))
            {
                return valeur[findStr];
            }
            return findStr;
        }

        #endregion
    }
}
