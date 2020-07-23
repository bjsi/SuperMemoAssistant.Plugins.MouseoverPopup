using MouseoverPopup.Interop;
using PluginManager.Interop.Sys;
using SuperMemoAssistant.Interop.Plugins;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public class MouseoverService : PerpetualMarshalByRefObject, IMouseoverSvc
  {

    public bool RegisterProvider(string name, string[] urlRegexes, IMouseoverContentProvider provider)
    {

      return Svc<MouseoverPopupPlugin>.Plugin.RegisterProvider(name, urlRegexes, provider);

    }

    public bool RegisterProvider(string name, string[] urlRegexes, KeywordScanningOptions keywordScanningOptions, IMouseoverContentProvider provider)
    {

      return Svc<MouseoverPopupPlugin>.Plugin.RegisterProvider(name, urlRegexes, keywordScanningOptions, provider);

    }

    public event EventHandler<HtmlPopupEventArgs> OnShow;

    public void InvokeOnShow(object sender, HtmlPopupEventArgs e)
    {

      OnShow?.Invoke(sender, e);

    }

  }
}
