namespace WordLibrary.Interface
{
    public interface ITemplateBuilder
    {
        ITemplate createTemplate(string name, string templateContent);
        ITemplate createTemplateFromFile(string name, string path);
        bool canBuild(string type);
    }
}
