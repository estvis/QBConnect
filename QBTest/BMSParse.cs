using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.XPath;

namespace QBTest
{
    abstract class BMSParse
    {
        string file;

        protected BMSParse(string file)
        {
            this.file = file;
        }

        public void ReadFile()
        {
            if (!File.Exists(file)) { return; }

            string data = File.ReadAllText(file);
            StringReader rdr = new StringReader(data);
            XPathDocument doc = new XPathDocument(rdr);
            XPathNavigator nav = doc.CreateNavigator();


        }
    }
}
