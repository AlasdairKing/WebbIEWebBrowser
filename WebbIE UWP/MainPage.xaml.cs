using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace WebbIE_10
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            webMain.Navigate(new Uri("http://www.aljo.org.uk"));
        }

        private void txtAddress_KeyDown(object sender, KeyRoutedEventArgs e)
        {
            if (e.Key == Windows.System.VirtualKey.Enter)
            {
                Uri address = null;
                try
                {
                    address = new Uri(txtAddress.Text);
                }
                catch
                {
                    address = new Uri("https://www.google.com?search=" + txtAddress.Text.Replace(" ", "%20"));
                }
                webMain.Navigate(address);
            }
        }

        private void SetWindowTitle(string text)
        {
            Windows.UI.ViewManagement.ApplicationView appView = Windows.UI.ViewManagement.ApplicationView.GetForCurrentView();
            appView.Title = text;

        }

        private void webMain_NavigationCompleted(WebView sender, WebViewNavigationCompletedEventArgs args)
        {
            txtAddress.Text = args.Uri.ToString();
            if (args.IsSuccess)
            {
                SetWindowTitle(webMain.DocumentTitle);
                UpdateContent();
            }
            else
            {
                SetWindowTitle(args.WebErrorStatus.ToString());
                txtMain.Text = args.WebErrorStatus.ToString();                
            }
            txtMain.Focus(FocusState.Programmatic);
        }

        private async void UpdateContent()
        {
            //txtMain.Text = await webMain.InvokeScriptAsync("eval", new string[] { "document.body.textContent;" });

            var file = await Windows.Storage.ApplicationData.Current.LocalFolder.GetFileAsync("edgie.js");
            var lines = await Windows.Storage.FileIO.ReadLinesAsync(file);


            lines[0] = "window.GetContent = " + lines[0];
            await webMain.InvokeScriptAsync("eval", lines);
            txtMain.Text = await webMain.InvokeScriptAsync("GetContent()", new string[] { "" });
            //await webMain.InvokeScriptAsync("window.Test()", null);


        }

        private void txtAddress_GotFocus(object sender, RoutedEventArgs e)
        {
            txtAddress.SelectAll();
        }
    }
}
