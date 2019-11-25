using System;
using System.Collections.Generic;
using WordLibrary.Interface;

namespace WordLibrary.Abstract
{
    public abstract class AbstractTemplate : ITemplate
    {
        protected string _name;
        protected string _type;
        private static string DEFAULT = @"{DEFAULT}";
        protected IDictionary<string, IFormatProvider> formatters
        {
            get;
            set;
        }

        public AbstractTemplate(string name, string type)
        {
            this._type = type;
            this._name = name;
            this.formatters = new Dictionary<string, IFormatProvider>();
        }

        public string getType()
        {
            return _type;
        }

        public IFormatProvider getFormatter(string key)
        {
            if (formatters.ContainsKey(key))
                return formatters[key];
            else
                return formatters[DEFAULT];


        }

        public void registerFormatter(IFormatProvider formatter, string type = null)
        {
            if (type == null)
            {
                formatters.Add(DEFAULT, formatter);
            }
            else
            {
                formatters.Add(type, formatter);
            }

        }

        public string getName()
        {
            return _name;
        }



        public abstract string getContent();

        public abstract System.IO.Stream getContentAsStream();
    }
}
