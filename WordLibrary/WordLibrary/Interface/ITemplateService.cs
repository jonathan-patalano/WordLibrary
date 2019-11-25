using System.IO;

namespace WordLibrary.Interface
{
    public enum TemplateServiceOutputFormat { Pdf, Docx, Html };

    public interface ITemplateService
    {
        /// <summary>
        /// Génère un document en sortie à partir d'un modèle de document
        /// </summary>
        /// <param name="template">Modèle instancié pour servir de base aux données fournies par le modèle de données</param>
        /// <param name="outputFormat">Format de sortie</param>
        /// <param name="model">Modèle de données à fournir au modèle de document</param>
        /// <returns></returns>
        void renderDocument(ITemplate template, TemplateServiceOutputFormat outputFormat, ITemplateModel model, out string outDocument);

        void renderDocument(ITemplate template, TemplateServiceOutputFormat outputFormat, ITemplateModel model, ref Stream outStream);


        ITemplateBuilder getTemplateBuilder(string type);
        //TODO faire plutot un DefaultTemplateBuilder
        ITemplate createTemplateFromString(string name, string templateContent, string type);
    }
}
