using MouseoverPopupInterfaces;
using PluginManager.Interop.Sys;
using SuperMemoAssistant.Services;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public class MouseoverService : PerpetualMarshalByRefObject, IMouseoverSvc
  {
    public bool RegisterProvider(string name, string[] urlRegexes, IMouseoverContentProvider provider)
    {
      return Svc<MouseoverPopupPlugin>.Plugin.RegisterProvider(name, urlRegexes, provider);
    }
  }
}
