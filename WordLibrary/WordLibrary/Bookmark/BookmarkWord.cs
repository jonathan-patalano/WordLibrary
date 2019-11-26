using Xceed.Words.NET;

namespace WordLibrary.Bookmark
{
    public static class BookmarkWord
    {
        #region Public Methods
        /// <summary>
        /// On récupére le document cible, ainsi que le texte qui doit être remplacer, puis on remplace ce texte par le texte souhaitée 
        /// </summary>
        /// <param name="document"></param>
        /// <param name="textecible"></param>
        /// <param name="texteremplacement"></param>
        public static void ReplaceText(DocX document, string textecible, string texteremplacement)
        {
            // On recupere notre textecible puis on le remplace par texteremplacement.
            var Bookmark = document.Bookmarks[textecible];
            if (Bookmark != null)
            {
                Bookmark.SetText(texteremplacement);
            }
        }
        #endregion
    }
}
