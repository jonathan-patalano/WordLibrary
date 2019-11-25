using NUnit.Framework;
using System.Collections.Generic;

namespace WordLibrary.TU
{
    public class WordTest
    {

        public static Dictionary<string, object> dictionnaire = new Dictionary<string, object>();


        // Test d'écriture d'une image. Possibilité de verification visuelle de l'écriture dans le fichier: 
        // imageRempli.docx (dossier UnitTest)
        [Test]
        public void TestImage()
        {
            string template = new FileTemplate("a", @"D:/CD13/Publipostage/Test/UnitTest/image.docx", "s").ToString();
            WordHelper.EcrireImage(ref dictionnaire, template);
        }
    }
}
