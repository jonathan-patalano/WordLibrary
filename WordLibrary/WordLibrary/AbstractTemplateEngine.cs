using CG13.Infrastructure.CrossCutting.Template;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WordLibrary
{
    public abstract class AbstractTemplateEngine : ITemplateEngine
    {
        protected string sType;


        public AbstractTemplateEngine(string sType)
        {
            this.sType = sType;
        }


        protected abstract bool checkOutputFormat(TemplateServiceOutputFormat output);

        public bool isValidFor(ITemplate template, TemplateServiceOutputFormat output)
        {
            return (sType.Equals(template.getType()) && checkOutputFormat(output));
        }


        public abstract void renderDocument(ITemplate template, IDictionary<string, object> values, ref System.IO.Stream outputStream);

        public string renderDocument(ITemplate template, IDictionary<string, object> values)
        {
            StringBuilder sBuilder = new StringBuilder();

            Stream s = new MemoryStream();
            renderDocument(template, values, ref s);
            StreamReader reader = new StreamReader(s, Encoding.UTF8);
            return reader.ReadToEnd();
        }

        public string getStringValue(ITemplate template, string key, object value)
        {
            IFormatProvider formatter = template.getFormatter(key);


            return String.Format(formatter, "{0}", value);

        }

        /// <summary>
        /// Obtenir une picture à partir d'un tableau de byte, dont la taille est calculée pour qu'elle s'affiche correctement dans le document. 
        /// </summary>
        /// <param name="templateDoc"></param>
        /// <param name="byteArrayIn"></param>
        /// <returns></returns>
        public Picture GetPicture(DocX templateDoc, byte[] byteArrayIn)
        {
            using (var ms = new MemoryStream(byteArrayIn))
            {
                Xceed.Document.NET.Image image = templateDoc.AddImage(ms);
                Picture picture = image.CreatePicture();
                if ((picture.Width > templateDoc.PageWidth || picture.Height > templateDoc.PageHeight) && picture.Width >= picture.Height)
                {
                    picture.Height = (int)(templateDoc.PageWidth * ((float)picture.Height / picture.Width));
                    picture.Width = (int)templateDoc.PageWidth;
                }
                else if ((picture.Width > templateDoc.PageWidth || picture.Height > templateDoc.PageHeight) && picture.Width < picture.Height)
                {
                    picture.Width = (int)(595 * ((float)picture.Height / picture.Width));
                    picture.Height = 595;
                }

                return picture;
            }
        }

        /// <summary>
        /// Convertir une image en tableau de bytes.
        /// </summary>
        /// <param name="imageIn"></param>
        /// <returns></returns>
        public static byte[] imageToByteArray(System.Drawing.Image image)
        {
            ImageConverter imageConverter = new ImageConverter();
            byte[] byteArray = (byte[])imageConverter.ConvertTo(image, typeof(byte[]));
            return byteArray;
        }
    }
}
