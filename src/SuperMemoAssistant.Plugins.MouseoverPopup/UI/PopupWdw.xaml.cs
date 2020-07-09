using MouseoverPopup.Interop;
using mshtml;
using SuperMemoAssistant.Sys.Remoting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SuperMemoAssistant.Plugins.MouseoverPopup.UI
{
  /// <summary>
  /// Interaction logic for PopupWdw.xaml
  /// </summary>
  public partial class PopupWdw : Window
  {
    public PopupWdw(string url, IContentProvider provider, RemoteCancellationToken ct)
    {
      InitializeComponent();
      ct.Register(new ActionProxy(Cancelled));
      wb1.DocumentCompleted += Wb1_DocumentCompleted;
      FetchHtml(url, provider, ct);
    }

    private void Wb1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
    {
      foreach (HtmlElement link in wb1.Document.Links)
      {
        link.Click += Link_Click;
      }
    }

    private void Link_Click(object sender, HtmlElementEventArgs e)
    {
      var element = ((HtmlElement)sender);
      string href = element.GetAttribute("href");
    }

    private async Task FetchHtml(string url, IContentProvider provider, RemoteCancellationToken ct)
    {
      string html = await provider.FetchHtml(ct, url);
      if (string.IsNullOrEmpty(html))
        return;
      wb1.Navigate(html);
    }

    private void Cancelled() 
    {
      Close();
    }
  }
}
