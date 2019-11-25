using System;
using System.Collections.Generic;
using System.IO;

namespace WordLibrary.Interface
{
    public interface ITemplateEngine
    {
        /// <summary>
        /// Permet de valider que ce moteur est valide pour ce type de template
        /// </summary>
        /// <param name="template">Template à valider</param>
        bool isValidFor(ITemplate template, TemplateServiceOutputFormat output);
        void renderDocument(ITemplate template, IDictionary<String, Object> values, ref Stream outputStream);
        string renderDocument(ITemplate template, IDictionary<string, object> values);
    }
}
