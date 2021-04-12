using mshtml;
using SuperMemoAssistant.Extensions;
using System;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public enum EventType
  {

    onkeydown,
    onclick,
    ondblclick,
    onkeypress,
    onkeyup,
    onmousedown,
    onmousemove,
    onmouseout,
    onmouseover,
    onmouseup,
    onselectstart,
    onbeforecopy,
    onbeforecut,
    onbeforepaste,
    oncontextmenu,
    oncopy,
    oncut,
    ondrag,
    ondragstart,
    ondragend,
    ondragenter,
    ondragleave,
    ondragover,
    ondrop,
    onfocus,
    onlosecapture,
    onpaste,
    onpropertychange,
    onreadystatechange,
    onresize,
    onactivate,
    onbeforedeactivate,
    oncontrolselect,
    ondeactivate,
    onmouseenter,
    onmouseleave,
    onmove,
    onmoveend,
    onmovestart,
    onpage,
    onresizeend,
    onresizestart,
    onfocusin,
    onfocusout,
    onmousewheel,
    onbeforeeditfocus,
    onafterupdate,
    onbeforeupdate,
    ondataavailable,
    ondatasetchanged,
    ondatasetcomplete,
    onerrorupdate,
    onfilterchange,
    onhelp,
    onrowenter,
    onrowexit,
    onlayoutcomplete,
    onblur,
    onrowsdelete,
    onrowsinserted,

  }

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

  [Serializable]
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
