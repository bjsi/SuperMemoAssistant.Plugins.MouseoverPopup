using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Interop.SuperMemo.Core;
using SuperMemoAssistant.Services;
using SuperMemoAssistant.Sys.Remoting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

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
    onbeforeeditfocus

    // { "onafterupdate" }
    // { "onbeforeupdate" }
    // { "ondataavailable" }
    // { "ondatasetchanged" }
    // { "ondatasetcomplete" }
    // { "onerrorupdate" }
    // { "onfilterchange" }
    // { "onhelp" }
    // { "onrowenter" }
    // { "onrowexit" }
    // { "onlayoutcomplete" }
    // { "onblur" () => BlurHandler() },
    // { "onrowsdelete" }
    // { "onrowsinserted" }
    // { "onlayoutcomplete" }

  }

  public class EventInitOptions
  {
    /// <summary>
    /// Type of event
    /// </summary>
    public EventType Type { get; set; }

    /// <summary>
    /// Filter the event to certain controls
    /// </summary>
    public Func<IControl, bool> ControlSelector { get; set; }

    /// <summary>
    /// Filter the event to certain HTMLElements
    /// TODO: Had to change to object because "cannot be used across assembly boundaries because it has a generic type argument that is an embedded interop type."
    /// </summary>
    public Func<object, bool> IHTMLElementSelector { get; set; }

    public EventInitOptions(EventType @event, Func<IControl, bool> controlSelector = null, Func<object, bool> elementSelector = null)
    {
      this.Type = @event;
      this.ControlSelector = controlSelector;
      this.IHTMLElementSelector = elementSelector;
    }
  }

  public partial class HTMLControlEvents
  {
    public event EventHandler<IHTMLControlEventArgs> OnKeyDownEvent;
    public event EventHandler<IHTMLControlEventArgs> OnClickEvent;
    public event EventHandler<IHTMLControlEventArgs> OnFocusedEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMouseOverEvent;
    public event EventHandler<IHTMLControlEventArgs> OnKeyPressEvent;
    public event EventHandler<IHTMLControlEventArgs> OnKeyUpEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMouseDownEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMouseMoveEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMouseOutEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMouseUpEvent;
    public event EventHandler<IHTMLControlEventArgs> OnSelectStartEvent;
    public event EventHandler<IHTMLControlEventArgs> OnBeforeCopyEvent;
    public event EventHandler<IHTMLControlEventArgs> OnBeforeCutEvent;
    public event EventHandler<IHTMLControlEventArgs> OnBeforePasteEvent;
    public event EventHandler<IHTMLControlEventArgs> OnContextMenuEvent;
    public event EventHandler<IHTMLControlEventArgs> OnCopyEvent;
    public event EventHandler<IHTMLControlEventArgs> OnCutEvent;
    public event EventHandler<IHTMLControlEventArgs> OnDragEvent;
    public event EventHandler<IHTMLControlEventArgs> OnDragStart;
    public event EventHandler<IHTMLControlEventArgs> OnDragEnterEvent;
    public event EventHandler<IHTMLControlEventArgs> OnDragEndEvent;
    public event EventHandler<IHTMLControlEventArgs> OnDragLeaveEvent;
    public event EventHandler<IHTMLControlEventArgs> OnDragOverEvent;
    public event EventHandler<IHTMLControlEventArgs> OnDropEvent;
    public event EventHandler<IHTMLControlEventArgs> OnFocusEvent;
    public event EventHandler<IHTMLControlEventArgs> OnLoseCaptureEvent;
    public event EventHandler<IHTMLControlEventArgs> OnPasteEvent;
    public event EventHandler<IHTMLControlEventArgs> OnPropertyStateChangeEvent;
    public event EventHandler<IHTMLControlEventArgs> OnReadyStateChangeEvent;
    public event EventHandler<IHTMLControlEventArgs> OnResizeEvent;
    public event EventHandler<IHTMLControlEventArgs> OnActivateEvent;
    public event EventHandler<IHTMLControlEventArgs> OnBeforeDeactivateEvent;
    public event EventHandler<IHTMLControlEventArgs> OnControlSelectEvent;
    public event EventHandler<IHTMLControlEventArgs> OnDeactivateEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMouseEnterEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMouseLeaveEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMoveEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMoveEndEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMoveStartEvent;
    public event EventHandler<IHTMLControlEventArgs> OnPageEvent;
    public event EventHandler<IHTMLControlEventArgs> OnResizeEndEvent;
    public event EventHandler<IHTMLControlEventArgs> OnResizeStartEvent;
    public event EventHandler<IHTMLControlEventArgs> OnFocusInEvent;
    public event EventHandler<IHTMLControlEventArgs> OnFocusOutEvent;
    public event EventHandler<IHTMLControlEventArgs> OnMouseWheelEvent;
    public event EventHandler<IHTMLControlEventArgs> OnBeforeEditFocus;
  }

  public partial class HTMLControlEvents
  {
    private Dictionary<EventType, EventHandler<IHTMLControlEventArgs>> EventMap => new Dictionary<EventType, EventHandler<IHTMLControlEventArgs>>
        {
            { EventType.onkeydown,              OnKeyDownEvent },
            { EventType.onclick,                OnClickEvent },
            { EventType.onmouseover,            OnMouseOverEvent },
            { EventType.onkeypress,             OnKeyPressEvent },
            { EventType.onkeyup,                OnKeyUpEvent },
            { EventType.onmousedown,            OnMouseDownEvent },
            { EventType.onmousemove,            OnMouseMoveEvent },
            { EventType.onmouseout,             OnMouseOutEvent },
            { EventType.onmouseup,              OnMouseUpEvent },
            { EventType.onselectstart,          OnSelectStartEvent },
            { EventType.onbeforecopy,           OnBeforeCopyEvent },
            { EventType.onbeforecut,            OnBeforeCutEvent },
            { EventType.onbeforepaste,          OnBeforePasteEvent },
            { EventType.oncontextmenu,          OnContextMenuEvent },
            { EventType.oncopy,                 OnCopyEvent },
            { EventType.oncut,                  OnCutEvent },
            { EventType.ondrag,                 OnDragEvent },
            { EventType.ondragend,              OnDragEndEvent },
            { EventType.ondragenter,            OnDragEnterEvent },
            { EventType.ondragleave,            OnDragLeaveEvent },
            { EventType.ondragover,             OnDragOverEvent },
            { EventType.ondrop,                 OnDropEvent },
            { EventType.onfocus,                OnFocusEvent },
            { EventType.onlosecapture,          OnLoseCaptureEvent },
            { EventType.onpaste,                OnPasteEvent },
            { EventType.onpropertychange,       OnPropertyStateChangeEvent },
            { EventType.onreadystatechange,     OnReadyStateChangeEvent },
            { EventType.onresize,               OnResizeEvent },
            { EventType.onactivate,             OnActivateEvent },
            { EventType.onbeforedeactivate,     OnBeforeDeactivateEvent },
            { EventType.oncontrolselect,        OnControlSelectEvent },
            { EventType.ondeactivate,           OnDeactivateEvent },
            { EventType.onmouseenter,           OnMouseEnterEvent },
            { EventType.onmouseleave,           OnMouseLeaveEvent },
            { EventType.onmove,                 OnMoveEvent },
            { EventType.onmoveend,              OnMoveEndEvent },
            { EventType.onmovestart,            OnMoveStartEvent },
            { EventType.onpage,                 OnPageEvent },
            { EventType.onresizeend,            OnResizeEndEvent },
            { EventType.onresizestart,          OnResizeStartEvent },
            { EventType.onfocusin,              OnFocusInEvent },
            { EventType.onfocusout,             OnFocusOutEvent },
            { EventType.onmousewheel,           OnMouseWheelEvent },
            { EventType.ondragstart,            OnDragStart},
            { EventType.onbeforeeditfocus,      OnBeforeEditFocus},
        };
  }

  public partial class HTMLControlEvents
  {
    public List<EventInitOptions> ToSubscribeTo { get; set; }

    public HTMLControlEvents(List<EventInitOptions> events)
    {
      this.ToSubscribeTo = events;
      Svc.SM.UI.ElementWdw.OnElementChanged += new ActionProxy<SMDisplayedElementChangedEventArgs>(OnElementChanged);
    }

    private void OnElementChanged(SMDisplayedElementChangedEventArgs obj)
    {
      SubscribeToHtmlDocEvents();
    }

    private void SubscribeToHtmlDocEvents()
    {
      if (ToSubscribeTo == null || ToSubscribeTo.Count == 0)
        return;

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      if (ctrlGroup == null)
        return;

      for (int i = 0; i < ctrlGroup.Count; i++)
      {
        var htmlCtrl = ctrlGroup[i]?.AsHtml();
        var htmlDoc = htmlCtrl?.GetDocument();
        var all = htmlDoc?.all;
        if (all == null || all.length == 0)
          continue;

        foreach (IHTMLElement2 node in all)
        {
          foreach (var htmlDocEvent in ToSubscribeTo)
          {

            var elementSelector = htmlDocEvent.IHTMLElementSelector;
            var controlSelector = htmlDocEvent.ControlSelector;

            if (elementSelector != null && !elementSelector(node))
              continue;

            if (controlSelector != null && !controlSelector(htmlCtrl))
              continue;

            EventType type = htmlDocEvent.Type;
            if (!EventMap.TryGetValue(type, out var csharpEvent))
            {
              // LogTo.Error
              continue;
            }

            string eventName = htmlDocEvent.Type.Name();
            var comEventObj = new GenericHtmlDocEventHandler(csharpEvent, i);

            try
            {
              node.attachEvent(eventName, comEventObj);
            }
            catch (Exception)
            {
              // Log
              // Occasionally there is an access denied error.
            }
          }
        }
      }
    }
  }


  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  public class GenericHtmlDocEventHandler : IHTMLControlEvent
  {
    public int ControlIdx { get; }
    public EventHandler<IHTMLControlEventArgs> EventHandler { get; }

    public GenericHtmlDocEventHandler(EventHandler<IHTMLControlEventArgs> eventHandler, int ControlIdx)
    {
      this.EventHandler = eventHandler;
      this.ControlIdx = ControlIdx;
    }

    [DispId(0)]
    public void handler(IHTMLEventObj e)
    {
      if (EventHandler != null)
      {
        var args = new IHTMLControlEventArgs(e, ControlIdx);
        EventHandler(this, args);
      }
    }
  }

  public interface IHTMLControlEvent
  {
    int ControlIdx { get; }
    void handler(IHTMLEventObj e);
    EventHandler<IHTMLControlEventArgs> EventHandler { get; }
  }

  public class IHTMLControlEventArgs : RoutedEventArgs
  {
    public IHTMLEventObj EventObj { get; set; }
    public int ControlIdx { get; set; }
    public IHTMLControlEventArgs(IHTMLEventObj EventObj, int ControlIdx)
    {
      this.EventObj = EventObj;
      this.ControlIdx = ControlIdx;
    }
  }
}
