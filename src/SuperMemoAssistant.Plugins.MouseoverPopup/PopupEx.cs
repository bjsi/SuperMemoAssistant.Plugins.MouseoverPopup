using MouseoverPopup.Interop;
using mshtml;
using SuperMemoAssistant.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{

  public class HtmlPopupOptions
  {

    public int x;
    public int y;
    public int width;
    public int height;

    public HtmlPopupOptions(int x, int y, int width, int height)
    {

      this.x = x;
      this.y = y;
      this.width = width;
      this.height = height;

    }
  }

  public class HtmlPopup
  {

    private IHTMLPopup _popup { get; set; }

    public event EventHandler<HtmlPopupEventArgs> OnShow;

    // Link button click event
    public event EventHandler<IControlHtmlEventArgs> OnLinkClick;
    private HtmlEvent _linkClickEvent { get; set; }

    // Extract button click event
    public event EventHandler<IControlHtmlEventArgs> OnExtractButtonClick;
    private HtmlEvent _extractButtonClickEvent { get; set; }

    // Browser button click event
    public event EventHandler<IControlHtmlEventArgs> OnBrowserButtonClick;
    private HtmlEvent _browserButtonClickEvent { get; set; }

    // Goto button click event
    public event EventHandler<IControlHtmlEventArgs> OnGotoButtonClick;
    private HtmlEvent _gotoButtonClickEvent { get; set; }

    // Edit button click event
    public event EventHandler<IControlHtmlEventArgs> OnEditButtonClick;
    private HtmlEvent _editButtonClickEvent { get; set; }


    public HtmlPopupOptions Options { get; set; }

    // For icons I created an Images folder in the Plugins\Development\PluginName folder
    private string outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);


    public HtmlPopup(IHTMLWindow4 wdw)
    {
      _popup = wdw?.createPopup() as IHTMLPopup;
      StylePopup();
    }

    private void StylePopup()
    {

      var popupDoc = GetDocument();
      if (popupDoc.IsNull())
        return;

      // Popup Styling
      popupDoc.body.style.border = "solid black 1px";
      popupDoc.body.style.overflow = "scroll";
      popupDoc.body.style.margin = "7px";

    }

    public void AddContent(string content)
    {

      var popupDoc = GetDocument();
      if (popupDoc.IsNull())
        return;

      popupDoc.body.innerHTML = content;
    }

    public void Show(HtmlPopupOptions opts)
    {

      if (_popup.IsNull())
        return;

      this.Options = opts;

      SubscribeToLinkClickEvents();
      _popup.Show(opts.x, opts.y, opts.width, opts.height, null);
      OnShow?.Invoke(this, new HtmlPopupEventArgs(opts.x, opts.y, opts.width, opts.height, _popup));

    }

    public void AddBrowserButton()
    {

      var htmlDoc = GetDocument();
      if (htmlDoc.IsNull())
        return;

      var iconPath = Path.Combine(outPutDirectory, "Images\\web.png");
      string icon_path = new Uri(iconPath).LocalPath;

      // Create browser button
      var browserBtn = htmlDoc.createElement("<button>");
      browserBtn.id = "browser-btn";

      browserBtn.innerHTML = $"<span><img src='{icon_path}' width='16px' height='16px' margin='5px'></span>";
      browserBtn.style.margin = "10px";

      // Create open button
      ((IHTMLDOMNode)htmlDoc.body).appendChild((IHTMLDOMNode)browserBtn);

      SubscribeToBrowserButtonClickEvents(((IHTMLElement2)browserBtn));
    }

    public void AddEditButton()
    {

      var htmlDoc = GetDocument();
      if (htmlDoc.IsNull())
        return;

      var iconPath = Path.Combine(outPutDirectory, "Images\\Editor.png");
      string icon_path = new Uri(iconPath).LocalPath;

      var editBtn = htmlDoc.createElement("<button>");
      editBtn.id = "edit-btn";

      editBtn.innerHTML = $"<span><img src='{icon_path}' width='16px' height='16px' margin='5px'></span>";
      editBtn.style.margin = "10px";

      // Create open button
      ((IHTMLDOMNode)htmlDoc.body).appendChild((IHTMLDOMNode)editBtn);

      SubscribeToEditButtonClickEvent(((IHTMLElement2)editBtn));

    }

    private void SubscribeToEditButtonClickEvent(IHTMLElement2 btn)
    {

      if (btn.IsNull())
        return;

      _editButtonClickEvent = new HtmlEvent();
      btn.SubscribeTo(EventType.onclick, _editButtonClickEvent);
      _editButtonClickEvent.OnEvent += (sender, e) => OnBrowserButtonClick?.Invoke(sender, e);

    }

    public void AddGotoButton()
    {

      var iconPath = Path.Combine(outPutDirectory, "Images\\GotoElement.jpg");
      string icon_path = new Uri(iconPath).LocalPath;

      var htmlDoc = GetDocument();
      if (htmlDoc.IsNull())
        return;

      // Create goto element button
      var gotobutton = htmlDoc.createElement("<button>");
      gotobutton.id = "goto-btn";
      gotobutton.innerHTML = $"<span><img src='{icon_path}' width='16px' height='16px' margin='5px'></span>";
      gotobutton.style.margin = "10px";

      // Add Goto button
      ((IHTMLDOMNode)htmlDoc.body).appendChild((IHTMLDOMNode)gotobutton);

      SubscribeToGotoButtonClickEvents(((IHTMLElement2)gotobutton));

    }

    public void AddExtractButton()
    {

      var htmlDoc = GetDocument();
      if (htmlDoc.IsNull())
        return;

      // Create Extract button
      var extractBtn = htmlDoc.createElement("<button>");
      extractBtn.id = "extract-btn";

      var iconPath = Path.Combine(outPutDirectory, "Images\\SMExtract.png");
      string icon_path = new Uri(iconPath).LocalPath;

      extractBtn.innerHTML = $"<span><img src='{icon_path}' width='16px' height='16px' margin='5px'></span>";
      //extractBtn.innerText = "Extract";
      extractBtn.style.margin = "10px";

      // Add extract button
      ((IHTMLDOMNode)htmlDoc.body).appendChild((IHTMLDOMNode)extractBtn);

      SubscribeToExtractButtonClickEvents(((IHTMLElement2)extractBtn));

    }

    private void SubscribeToGotoButtonClickEvents(IHTMLElement2 btn)
    {

      if (btn.IsNull())
        return;
      
      _gotoButtonClickEvent = new HtmlEvent();
      btn.SubscribeTo(EventType.onclick, _gotoButtonClickEvent  );
      _gotoButtonClickEvent.OnEvent += (sender, e) => OnGotoButtonClick?.Invoke(sender, e);

    }

    private void SubscribeToLinkClickEvents() 
    {

      var htmlDoc = GetDocument();
      var body = htmlDoc?.body;
      if (body.IsNull())
        return;

      // subscribing to link click events on a elements was unreliable
      _linkClickEvent = new HtmlEvent();
      ((IHTMLElement2)body).SubscribeTo(EventType.onclick, _linkClickEvent);
      _linkClickEvent.OnEvent += (sender, e) => OnLinkClick?.Invoke(sender, e);
    
    }

    public void SubscribeToExtractButtonClickEvents(IHTMLElement2 btn)
    {

      if (btn.IsNull())
        return;

      _extractButtonClickEvent = new HtmlEvent();
      btn.SubscribeTo(EventType.onclick, _extractButtonClickEvent);
      _extractButtonClickEvent.OnEvent += (sender, e) => OnExtractButtonClick?.Invoke(sender, e);

    }

    public void SubscribeToBrowserButtonClickEvents(IHTMLElement2 btn)
    {

      if (btn.IsNull())
        return;

      _browserButtonClickEvent = new HtmlEvent();
      btn.SubscribeTo(EventType.onclick, _browserButtonClickEvent);
      _browserButtonClickEvent.OnEvent += (sender, e) => OnBrowserButtonClick?.Invoke(sender, e);

    }

    public int GetOffsetHeight(int width)
    {

      if (_popup.IsNull())
        return -1;

      var doc = _popup.document as IHTMLDocument2;
      var body = doc?.body;
      if (body.IsNull())
        return -1;

      _popup.Show(0, 0, width, 1, null);
      int height = ((IHTMLElement2)body).scrollHeight;
      _popup.Hide();

      return height;

    }

    public bool IsOpen()
    {
      return _popup.IsNull()
        ? false
        : _popup.isOpen;
    }

    public void Hide()
    {
      _popup.Hide();
    }

    public IHTMLDocument2 GetDocument()
    {
      return _popup?.document as IHTMLDocument2;
    }

  }

  public static class PopupEx
  {
    public static HtmlPopup CreatePopup(this IHTMLWindow4 wdw)
    {
      return new HtmlPopup(wdw);
    }
  }
}
