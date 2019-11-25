using System;
using System.IO;

namespace WordLibrary.Interface
{
    public interface ITemplate
    {
        string getName();

        /// <summary>
        /// Retourne le type du modèle, ce qui permet de sélectionner le moteur de génération acceptant ce type
        /// </summary>
        string getType();

        /// <summary>
        /// Retourne le contenu du modèle
        /// </summary>
        string getContent();

        Stream getContentAsStream();

        IFormatProvider getFormatter(string key);

        void registerFormatter(IFormatProvider formatter, String key);
    }
}
