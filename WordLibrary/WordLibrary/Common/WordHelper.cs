using System.Collections.Generic;
using WordLibrary.Abstract;
using WordLibrary.Table;

namespace WordLibrary
{
    public class WordHelper
    {
        public Dictionary<string, object> EcrireImage(ref Dictionary<string, object> Dictionary, string imagePath)
        {
            System.Drawing.Image image = System.Drawing.Image.FromFile(imagePath);
            byte[] byteArray = AbstractTemplateEngine.imageToByteArray(image);
            Dictionary.Add(imagePath, byteArray);
            return Dictionary;
        }

        public Dictionary<string, object> EcrireTexte(ref Dictionary<string, object> Dictionary, string texte)
        {
            Dictionary.Add("Test", texte);
            return Dictionary;
        }

        public Dictionary<string, object> EcrireTableau(ref Dictionary<string, object> Dictionary, string entete, string valeur)
        {
            List<TableDocXEntete> listeEntete = new List<TableDocXEntete>();
            listeEntete.Add(new TableDocXEntete { NumeroColonne = 0, ValeurEntete = entete });
            List<TableDocXCellule> listeCellule = new List<TableDocXCellule>();
            listeCellule.Add(new TableDocXCellule { NumeroColonne = 0, NumeroLigne = 1, ValeurCellule = valeur });
            Dictionary.Add("Test", new TableDocX { ListeEntete = listeEntete, ListeCellule = listeCellule });
            return Dictionary;
        }

    }
}
