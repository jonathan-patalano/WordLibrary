using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WordLibrary
{
    public static class ImageWord
    {

        public static Picture AddPicture(DocX document, string image)
        {
            // Ajout d'une image par un flux
            var streamImage = document.AddImage(new FileStream(image, FileMode.Open, FileAccess.Read));
            var pictureStream = streamImage.CreatePicture(85, 85);
            return pictureStream;
        }

        public static void InsertImage(DocX document, Picture p)
        {
            document.Tables[0].Rows[0].Cells[2].Paragraphs[0].InsertPicture(p);
            document.Tables[0].Rows[0].Cells[2].Paragraphs[0].RemoveText(1, 15);
        }


        /// <summary>
        /// Obtenir une picture à partir d'un tableau de byte, dont la taille est calculée pour qu'elle s'affiche correctement dans le document. 
        /// </summary>
        /// <param name="templateDoc"></param>
        /// <param name="byteArrayIn"></param>
        /// <returns></returns>
        public static Picture GetPicture(DocX document, byte[] byteArrayIn)
        {
            using (var ms = new MemoryStream(byteArrayIn))
            {
                Image image = document.AddImage(ms);
                Picture picture = image.CreatePicture();
                if ((picture.Width > document.PageWidth || picture.Height > document.PageHeight) && picture.Width >= picture.Height)
                {
                    picture.Height = (int)(document.PageWidth * ((float)picture.Height / picture.Width));
                    picture.Width = (int)document.PageWidth;
                }
                else if ((picture.Width > document.PageWidth || picture.Height > document.PageHeight) && picture.Width < picture.Height)
                {
                    picture.Width = (int)(595 * ((float)picture.Height / picture.Width));
                    picture.Height = 595;
                }

                return picture;
            }
        }
    }
}
