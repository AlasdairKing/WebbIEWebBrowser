using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SHDocVw;

// My attempt to change user string for the browser so it works better in BBC TV. 
// Doesn't work, don't know why. 

namespace UserAgentTest
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            //TODO First, this doesn't work: you have to put a breakpoint on the registry key setting (?)
            ePubReader.WebBrowserUtility.SetWebBrowserEmulation(ePubReader.WebBrowserUtility.InternetExplorerEmulationMode.IE10NoMatterWhat, false);
            //webBrowser1.UserAgent = "Mozilla/5.0 (iPad; CPU OS 7_0 like Mac OS X) AppleWebKit/537.51.1 (KHTML, like Gecko) CriOS/30.0.1599.12 Mobile/11A465 Safari/8536.25 (3B92C18B-D9DE-4CB7-A02A-22FD2AF17C8F)";
            //webBrowser1.UserAgent = "Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)";
            // Second, this doesn't work. 
            webBrowser1.UserAgent = "Mozilla/5.0 (Linux; Android 4.4.4; en-us; Nexus 5 Build/JOP40D) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2307.2 Mobile Safari/537.36";
            //webBrowser1.Navigate("https://www.bbc.co.uk/iplayer/");
            webBrowser1.Navigate("http://www.whatsmyua.com/");
        }

        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            if (!e.Url.ToString().Contains("bbc.co.uk"))
            {
                System.Diagnostics.Debug.Print("URL: " + e.Url);
                //e.Cancel = true;
            }
        }
    }

    public class ExtendedWebBrowser : System.Windows.Forms.WebBrowser
    {
        bool renavigating = false;

        public string UserAgent { get; set; }

        public ExtendedWebBrowser()
        {
            DocumentCompleted += SetupBrowser;

            //this will cause SetupBrowser to run (we need a document object)
            Navigate("about:blank");
        }

        void SetupBrowser(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            DocumentCompleted -= SetupBrowser;
            SHDocVw.WebBrowser xBrowser = (SHDocVw.WebBrowser)ActiveXInstance;
            xBrowser.BeforeNavigate2 += BeforeNavigate;
            DocumentCompleted += PageLoaded;
        }

        void PageLoaded(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        void BeforeNavigate(object pDisp, ref object url, ref object flags, ref object targetFrameName,
            ref object postData, ref object headers, ref bool cancel)
        {
            // This alone is sufficient, because headers is a "Ref" parameters, and the browser seems to pick this back up.
            headers += string.Format("User-Agent: {0}\r\n", UserAgent);
        }
    }
}
