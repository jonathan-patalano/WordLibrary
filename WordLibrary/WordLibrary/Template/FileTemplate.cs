using System;
using System.IO;
using WordLibrary.Abstract;
using WordLibrary.Formatter;

namespace WordLibrary
{
    public class FileTemplate : AbstractTemplate
    {
        string _path;

        public FileTemplate(string name, string path, string type) : base(name, type)
        {
            this._path = path;
            registerFormatter(new DefaultFormatter());

        }

        public override string getContent()
        {
            throw new NotImplementedException();
        }

        public override Stream getContentAsStream()
        {

            FileStream fileStream = new FileStream(_path, FileMode.Open, FileAccess.Read);
            BufferedStream result = new BufferedStream(fileStream);

            return result;


        }
    }
}
