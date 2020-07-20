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

    public bool RegisterProvider(string name, List<string> urlRegexes, IMouseoverContentProvider provider)
    {

      if (name.IsNullOrEmpty())
        return false;

      if (urlRegexes.IsNull() || urlRegexes.Count == 0)
        return false;

      if (provider.IsNull())
        return false;

      return Svc<MouseoverPopupPlugin>.Plugin.RegisterProvider(name, urlRegexes, provider);

    }

    public event EventHandler<HtmlPopupEventArgs> OnShow;

    public void InvokeOnShow(object sender, HtmlPopupEventArgs e)
    {
      OnShow?.Invoke(sender, e);
    }

  }
}
