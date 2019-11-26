﻿using System.IO;
using Xceed.Words.NET;

namespace WordLibrary.ImageWord
{
    public static class ImageWord
    {

        public static void AddPicture(DocX document, string image)
        {
            // Ajout d'une image par un flux
            var streamImage = document.AddImage(new FileStream(image, FileMode.Open, FileAccess.Read));
            var pictureStream = streamImage.CreatePicture(150, 150);
            // TODO : Enlever le texte, pour ajout seul d'une image
            var p3 = document.InsertParagraph("Here is the same picture added from a stream:");
            p3.AppendPicture(pictureStream);
        }
    }
}