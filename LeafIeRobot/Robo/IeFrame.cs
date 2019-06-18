using mshtml;

namespace LeafIeRobot.Robo
{
    public class IeFrame
    {
        public IHTMLWindow2 Window;

        public IeDocument Document => new IeDocument((HTMLDocument) Window.document);

        public IeFrame(IHTMLWindow2 window)
        {
            Window = window;
        }



    }
}
