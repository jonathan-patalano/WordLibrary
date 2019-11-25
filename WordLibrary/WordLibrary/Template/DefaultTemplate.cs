using System;
using WordLibrary.Abstract;
using WordLibrary.Formatter;

namespace WordLibrary
{
    class DefaultTemplate : AbstractTemplate
    {
        private string _content;

        public DefaultTemplate(string name, string content, string type) : base(name, type)
        {
            this._content = content;
            registerFormatter(new DefaultFormatter());
        }

        public override string getContent()
        {
            return _content;
        }

        public override System.IO.Stream getContentAsStream()
        {
            throw new NotImplementedException();
        }
    }
}
