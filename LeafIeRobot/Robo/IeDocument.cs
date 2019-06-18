using System.Collections.Generic;
using mshtml;

namespace LeafIeRobot.Robo
{
    public class IeDocument
    {
        public HTMLDocument Document;
        public IHTMLWindow2 Window => Document.parentWindow;
        public IHTMLElement Body => Document.body;

        public IeDocument(HTMLDocument document)
        {
            this.Document = document;
        }

        public List<IeElement> GetElementsByTagName(string value)
        {
            var result = new List<IeElement>();

            var elements = Document.getElementsByTagName(value);

            foreach (IHTMLElement element in elements)
                result.Add(new IeElement(element));

            return result;
        }

        public List<IeElement> GetElementsByName(string value)
        {
            var result = new List<IeElement>();

            var elements = Document.getElementsByName(value);

            foreach (IHTMLElement element in elements)
                result.Add(new IeElement(element));

            return result;
        }

        public IeElement GetElementById(string value)
        {
            return new IeElement(Document.getElementById(value));
        }

    }
}
