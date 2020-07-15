using mshtml;
using SuperMemoAssistant.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{

  public interface IControlHtmlEvent
  {
    event EventHandler<IControlHtmlEventArgs> OnEvent;
    void handler(IHTMLEventObj e);
  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  public class HtmlEvent : IControlHtmlEvent
  {
    public event EventHandler<IControlHtmlEventArgs> OnEvent;

    [DispId(0)]
    public void handler(IHTMLEventObj e)
    {
      if (!OnEvent.IsNull())
        OnEvent(this, new IControlHtmlEventArgs(e));
    }
  }

  public class IControlHtmlEventArgs
  {
    public IHTMLEventObj EventObj { get; set; }
    public IControlHtmlEventArgs(IHTMLEventObj EventObj)
    {
      this.EventObj = EventObj;
    }
  }

  public static class HtmlEventsEx
  {
    public static bool SubscribeTo(this IHTMLElement2 element, EventType eventType, IControlHtmlEvent handlerObj)
    {
      try
      {

        return element.IsNull()
          ? false
          : element.attachEvent(eventType.Name(), handlerObj);

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return false;
    }
  }
}
