using mshtml;

namespace LeafIeRobot.Robo
{
    public class IeElement
    {
        public IHTMLElement Element;

        public IeDocument Document => new IeDocument((HTMLDocument) Element.document);
        public bool Disabled => ((IHTMLElement3)Element).isDisabled;
        public bool Editable => ((IHTMLElement3)Element).isContentEditable;
        public string Id => GetAttribute("id");
        public string InnerHtml => Element.innerHTML;
        public string InnerText => Element.innerText;
        public string Name => GetAttribute("name");
        public string OuterHtml => Element.outerHTML;
        public string OuterText => Element.outerText;
        public string Class => Element.className ?? "";
        public string TagName => Element.tagName.ToLower();

        public IeElement(IHTMLElement element)
        {
            Element = element;
        }

        public void SetAttribute(string value, string field = "value")
        {
            Element.setAttribute(field, value);
        }

        public string GetAttribute(string field)
        {
            string result = "";
            try
            {
                result = Element.getAttribute(field).ToString();
            }
            catch
            {                
            }

            return result;
        }

        public void Click()
        {
            Element.click();
        }

        public void DoubleClick()
        {
            Element.click();
            Element.click();
        }

        public void Clear()
        {
            SetAttribute("");
        }

    }
        
    
}
