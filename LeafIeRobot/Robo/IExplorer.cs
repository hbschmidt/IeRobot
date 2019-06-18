using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using InternetExplorerAutomation.Classes;
using mshtml;
using Microsoft.Win32;
using SHDocVw;

namespace LeafIeRobot.Robo
{
    public class IExplorer : Imports
    {
        private string _userAgent = "";
        public enum VersaoIe
        { V7, V8, V9, V10, V11 };

        protected InternetExplorer Navegador;
        public string Url => Navegador.LocationURL.ToString();

        public IeDocument Document => new IeDocument((HTMLDocument)Navegador.Document);

        public IExplorer(bool mostrarNavegador = true, bool mostrarBarraDeGuias = false, bool mostrarBarraDeMenus = false, bool modoTeatro = false)
        {
            Navegador = new InternetExplorer
            {
                Visible = mostrarNavegador,
                AddressBar = mostrarBarraDeGuias,
                MenuBar = mostrarBarraDeMenus,
                ToolBar = mostrarBarraDeMenus ? 1 : 0,
            };

            if (!modoTeatro) return;

            Navegador.TheaterMode = true;
        }

        public void Wait(int seconds = -1)
        {
            while (seconds > 0)
            {
                seconds = seconds - 1;
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
            }

            while (Navegador.Busy || Navegador.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE)
                System.Threading.Thread.Sleep(100);
        }

        public void WaitText(string texto, bool mostrarTela)
        {
            if (mostrarTela) Navegador.Visible = true;
            while (Navegador.Busy || Document == null || !Document.Body.innerHTML.Contains(texto))
            {
                System.Threading.Thread.Sleep(200);
            }
            if (mostrarTela) Navegador.Visible = false;
        }

        public List<IeFrame> GetFrames()
        {
            var frames = new List<IeFrame>();

            for (var i = 0; i < Document.Window.frames.length; i++)
            {
                object refIdx = i;
                frames.Add(new IeFrame((IHTMLWindow2)Document.Window.frames.item(ref refIdx)));
            }

            return frames;
        }

        public object InvokeScript(string script)
        {
            return Document.Window.execScript(script, "javascript");
        }

        public void Show(double seconds)
        {
            var fim = DateTime.Now.AddSeconds(seconds);
            while (DateTime.Now < fim)
            {
                Navegador.Visible = true;
                System.Threading.Thread.Sleep(100);
            }

            Navegador.Visible = false;
        }

        public void Navigate(string link)
        {            
            Navegador.Navigate(link, null, null, null, _userAgent != "" ? _userAgent : null );
            Wait();
        }

        public void Post(string link, string parametros, string headers = "")
        {
            if (headers != "") headers = "Content-Type: application/x-www-form-urlencoded\n\r";

            if (_userAgent != "")
            {
                headers += _userAgent;
            }

            object flags = null;
            object frame = null;
            object bytes = (object)Encoding.ASCII.GetBytes(parametros);

            Navegador.Navigate(link, ref flags, ref frame, ref bytes, headers);

            Wait();
        }

        public void VersaoInternetExplorer(VersaoIe versao, string nomeProcesso)
        {
            switch (versao)
            {
                case VersaoIe.V7:
                    _userAgent = "Mozilla/5.0 (Windows; U; MSIE 7.0; Windows NT 6.0; en-US)";
                    EnsureBrowserEmulationEnabled(nomeProcesso, false, 7000);
                    break;
                case VersaoIe.V8:
                    break;
                case VersaoIe.V9:
                    break;
                case VersaoIe.V10:
                    break;
                case VersaoIe.V11:
                    _userAgent = "";
                    break;
                default:
                    break;
            }

            Navegador.BeforeNavigate2 += Navegador_BeforeNavigate2;

            //((SHDocVw.IWebBrowserApp)Navegador.Application).
        }

        private void Navegador_BeforeNavigate2(object pDisp, ref object url, ref object flags, ref object targetFrameName, ref object postData, ref object headers, ref bool cancel)
        {
            if (!string.IsNullOrEmpty(_userAgent))
            {
                //    if (!_renavigating)
                //    {
                headers += $"User-Agent: {_userAgent}\r\n";
                //        _renavigating = true;
                //        cancel = true;
                //        Navegador.Navigate((string)url, (string)targetFrameName, (byte[])postData, (string)headers);
                //    }
                //    else
                //    {
                //        _renavigating = false;
                //    }
            }
        }

        public void ChangeSize(int altura, int largura)
        {
            Navegador.Height = altura;
            Navegador.Width = largura;
        }

        public bool SwitchToWindow(string url)
        {
            var shellWindows = new ShellWindows();

            foreach (SHDocVw.WebBrowser webBrowser in shellWindows)
            {
                var nameWindow = webBrowser?.Name;
                if (nameWindow == null)
                    continue;

                if (nameWindow.Contains("Internet"))
                {
                    var urlWindow = webBrowser?.LocationURL;
                    if (urlWindow == null)
                        continue;

                    if (urlWindow.Equals(url))
                    {
                        Navegador = (InternetExplorer)webBrowser;
                        return true;
                    }
                }

            }

            return false;
        }

        public void CloseAlert(double seconds)
        {
            Task.Run(() => {
                System.Threading.Thread.Sleep(Convert.ToInt32(seconds * 1000));
                IntPtr hwnd = FindWindow("#32770", null);
                if (hwnd.ToString().Equals("0"))
                    return;

                int count = GetWindowTextLength(hwnd);

                StringBuilder txto = new StringBuilder(count);
                GetWindowText(hwnd, txto, txto.Capacity);

                hwnd = FindWindowEx(hwnd, IntPtr.Zero, "Button", "OK");
                SendMessage(hwnd, 0xf5, IntPtr.Zero, IntPtr.Zero); //focus
                SendMessage(hwnd, 0x0D, IntPtr.Zero, IntPtr.Zero); //click

            });
        }

        public List<IeElement> GetElementsByTagName(string value)
        {
            return Document.GetElementsByTagName(value);
        }
        public List<IeElement> GetElementByName(string value)
        {
            return Document.GetElementsByName(value);
        }
        public IeElement GetElementById(string value)
        {
            return Document.GetElementById(value);
        }

        public string GetValue(string name)
        {
            var element = Document.GetElementById(name);
            if (element != null)
                return element.GetAttribute("value");

            var elements = Document.GetElementsByName(name);
            if (elements.Count > 0)
                return elements[0].GetAttribute("value");

            return "";
        }

        public string GetText(string name)
        {
            var element = Document.GetElementById(name);
            if (element != null)
                return element.GetAttribute("text");

            var elements = Document.GetElementsByName(name);
            if (elements.Count > 0)
                return elements[0].GetAttribute("text");

            return "";
        }

        public void SetValue(string name, string value)
        {
            var element = Document.GetElementById(name);
            if (element != null)
            {
                element.SetAttribute(value);
                return;
            }

            var elements = Document.GetElementsByName(name);
            if (elements.Count > 0)
            {
                elements[0].SetAttribute(value);
                return;
            }
        }

        public void Clear(string name)
        {
            var element = Document.GetElementById(name);
            if (element != null)
            {
                element.SetAttribute("");
                return;
            }

            var elements = Document.GetElementsByName(name);
            if (elements.Count > 0)
            {
                elements[0].SetAttribute("");
                return;
            }
        }

        public static void EnsureBrowserEmulationEnabled(string exename, bool uninstall, uint chaveIe)
        {

            try
            {
                using (
                    var rk = Registry.CurrentUser.OpenSubKey(
                        @"SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", true)
                )
                {
                    if (!uninstall)
                    {
                        dynamic value = rk.GetValue(exename);
                        if (value == null)
                            rk.SetValue(exename, chaveIe, RegistryValueKind.DWord);
                    }
                    else
                        rk.DeleteValue(exename);
                }
            }
            catch
            {
            }
        }
    }
}
