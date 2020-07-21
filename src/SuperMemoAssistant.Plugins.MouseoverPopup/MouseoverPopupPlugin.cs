#region License & Metadata

// The MIT License (MIT)
// 
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the 
// Software is furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.
// 
// 
// Created On:   7/5/2020 10:47:16 PM
// Modified By:  james

#endregion




namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  using System;
  using System.Collections.Generic;
  using System.Diagnostics;
  using System.Diagnostics.CodeAnalysis;
  using System.IO;
  using System.Linq;
  using System.Reflection;
  using System.Runtime.Remoting;
  using System.Text.RegularExpressions;
  using System.Threading;
  using System.Threading.Tasks;
  using System.Windows;
  using Anotar.Serilog;
  using global::MouseoverPopup.Interop;
  using mshtml;
  using SuperMemoAssistant.Extensions;
  using SuperMemoAssistant.Interop.Plugins;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Contents;
  using SuperMemoAssistant.Interop.SuperMemo.Core;
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Builders;
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Models;
  using SuperMemoAssistant.Plugins.MouseoverPopup.Models;
  using SuperMemoAssistant.Plugins.PopupWindow.Interop;
  using SuperMemoAssistant.Services;
  using SuperMemoAssistant.Services.Sentry;
  using SuperMemoAssistant.Sys.Remoting;

  // ReSharper disable once UnusedMember.Global
  // ReSharper disable once ClassNeverInstantiated.Global
  [SuppressMessage("Microsoft.Naming", "CA1724:TypeNamesShouldNotMatchNamespaces")]
  public class MouseoverPopupPlugin : SentrySMAPluginBase<MouseoverPopupPlugin>
  {
    #region Constructors

    /// <inheritdoc />
    public MouseoverPopupPlugin() : base("Enter your Sentry.io api key (strongly recommended)") { }

    #endregion

    #region Properties Impl - Public

    /// <inheritdoc />
    public override string Name => "MouseoverPopup";

    /// <inheritdoc />
    public override bool HasSettings => true;

    /// <summary>
    /// Stores the content providers whose services are requesed on mouseover events.
    /// </summary>
    private Dictionary<string, ContentProviderInfo> providers { get; set; } = new Dictionary<string, ContentProviderInfo>();

    /// <summary>
    /// Service that providers can call to register themselves.
    /// </summary>
    private MouseoverService _mouseoverSvc = new MouseoverService();

    /// <summary>
    /// Service for communication between mouseover popup and popup window.
    /// </summary>
    private IPopupWindowSvc _popupWindowSvc { get; set; }

    private HtmlEvent AnchorElementMouseLeave { get; set; }
    private HtmlEvent MouseEnterEvent { get; set; }

    // Popup window events
    private HtmlEvent PopupWindowLinkClick { get; set; }
    private HtmlEvent ExtractButtonClick { get; set; }
    private HtmlEvent PopupBrowserButtonClick { get; set; }
    private HtmlEvent GotoElementButtonClick { get; set; }
    private HtmlEvent EditButtonClick { get; set; }

    // Used to run certain actions off of the UI thread
    // eg. item creation - otherwise will result in deadlock
    private EventfulConcurrentQueue<Action> EventQueue = new EventfulConcurrentQueue<Action>();

    public MouseoverPopupCfg Config;
    #endregion

    #region Methods Impl

    private void LoadConfig()
    {
      Config = Svc.Configuration.Load<MouseoverPopupCfg>() ?? new MouseoverPopupCfg();
    }

    /// <inheritdoc />
    protected override void PluginInit()
    {

      LoadConfig();

      PublishService<IMouseoverSvc, MouseoverService>(_mouseoverSvc);

      _popupWindowSvc = GetService<IPopupWindowSvc>();

      Svc.SM.UI.ElementWdw.OnElementChanged += new ActionProxy<SMDisplayedElementChangedEventArgs>(ElementWdw_OnElementChanged);

      // Runs the actions added to the event queue
      _ = Task.Factory.StartNew(DispatchEvents, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);

    }

    private void ElementWdw_OnElementChanged(SMDisplayedElementChangedEventArgs obj)
    {

      var htmlDoc = ContentUtils.GetFocusedHtmlDocument();
      if (htmlDoc.IsNull())
        return;

      MouseEnterEvent = new HtmlEvent();

      var all = htmlDoc.body?.all as IHTMLElementCollection;
      all?.Cast<IHTMLElement>()
         ?.Where(x => x.tagName.ToLowerInvariant() == "a")
         ?.ForEach(x => ((IHTMLElement2)x).SubscribeTo(EventType.onmouseenter, MouseEnterEvent));

      MouseEnterEvent.OnEvent += MouseEnterEvent_OnEvent;

    }

    private async void MouseEnterEvent_OnEvent(object sender, IControlHtmlEventArgs obj)
    {

      var ev = obj.EventObj;

      // Coordinates
      var x = ev.screenX;
      var y = ev.screenY;

      var anchor = obj.EventObj.srcElement as IHTMLAnchorElement;
      string url = anchor?.href;

      if (url.IsNullOrEmpty())
        return;

      if (providers.IsNull() || providers.Count == 0)
        return;

      var provider = MatchContentProvider(url);
      if (provider.IsNull())
        return;

      var parentWdw = Application.Current.Dispatcher.Invoke(() => ContentUtils.GetFocusedHtmlWindow());
      if (parentWdw.IsNull())
        return;

      var linkElement = anchor as IHTMLElement2;

      // Add mouseleave event to cancel early
      var rcts = new RemoteCancellationTokenSource();
      AnchorElementMouseLeave = new HtmlEvent();
      linkElement.SubscribeTo(EventType.onmouseleave, AnchorElementMouseLeave);
      AnchorElementMouseLeave.OnEvent += (sender, args) => rcts?.Cancel();

      try
      {
        await OpenNewPopupWdw((IHTMLWindow4)parentWdw, url, provider, rcts.Token, x, y);
      }
      catch (RemotingException) { }

    }

    private void DispatchEvents()
    {
      while (true)
      {
        EventQueue.DataAvailableEvent.WaitOne(3000);
        while (EventQueue.TryDequeue(out var action))
        {
          action();
        }
      }
    }

    private IMouseoverContentProvider MatchContentProvider(string url)
    {

      if (url.IsNullOrEmpty())
        return null;

      if (providers.IsNull() || providers.Count == 0)
        return null;

      // Find the matching provider
      foreach (var keyValuePair in providers)
      {
        var regexes = keyValuePair.Value.urlRegexes;
        var provider = keyValuePair.Value.provider;

        if (regexes.Any(r => new Regex(r).Match(url).Success))
        {
          return provider;
        }
      }

      return null;

    }

    private async Task OpenNewPopupWdw(IHTMLWindow4 parentWdw, string url, IMouseoverContentProvider provider, RemoteCancellationToken ct, int x, int y)
    {

      if (url.IsNullOrEmpty() || provider.IsNull() || parentWdw.IsNull())
        return;

      // Get html
      PopupContent content = await provider.FetchHtml(ct, url);
      if (content.IsNull() || content.Html.IsNullOrEmpty())
        return;


      // Create Popup
      await Application.Current.Dispatcher.BeginInvoke((Action)(() =>
      {
        try
        {

          var htmlDoc = ((IHTMLWindow2)parentWdw).document;
          if (htmlDoc.IsNull())
            return;

          var htmlCtrlBody = htmlDoc.body;
          if (htmlCtrlBody.IsNull())
            return;

          var popup = parentWdw.CreatePopup();
          popup.OnShow += (sender, args) => _mouseoverSvc?.InvokeOnShow(sender, args);
          if (popup.IsNull())
            return;

          var popupDoc = popup.GetDocument();
          if (popupDoc.IsNull())
            return;

          // Popup Styling
          popupDoc.body.style.border = "solid black 1px";
          //popupDoc.body.style.overflow = "scroll";
          popupDoc.body.style.margin = "7px";
          popupDoc.body.innerHTML = content.Html;

          // For icons I created an Images folder in the Plugins\Development\PluginName folder
          var outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);

          // Extract Button
          if (content.AllowExtract)
          {

            // Create Extract button
            var extractBtn = popupDoc.createElement("<button>");
            extractBtn.id = "extract-btn";

            var iconPath = Path.Combine(outPutDirectory, "Images\\SMExtract.png");
            string icon_path = new Uri(iconPath).LocalPath;

            extractBtn.innerHTML = $"<span><img src='{icon_path}' width='16px' height='16px' margin='5px'></span>";
            //extractBtn.innerText = "Extract";
            extractBtn.style.margin = "10px";

            // Add extract button
            ((IHTMLDOMNode)popupDoc.body).appendChild((IHTMLDOMNode)extractBtn);

            // Add click event
            ExtractButtonClick = new HtmlEvent();
            ((IHTMLElement2)extractBtn).SubscribeTo(EventType.onclick, ExtractButtonClick);
            ExtractButtonClick.OnEvent += (sender, e) => ExtractButtonClick_OnEvent(sender, e, content, popup);

          }

          // Goto Element Button
          if (content.AllowGotoInSM)
          {

            var iconPath = Path.Combine(outPutDirectory, "Images\\GotoElement.jpg");
            string icon_path = new Uri(iconPath).LocalPath;

            // Create goto element button
            var gotobutton = popupDoc.createElement("<button>");
            gotobutton.id = "goto-btn";
            gotobutton.innerHTML = $"<span><img src='{icon_path}' width='16px' height='16px' margin='5px'></span>";
            gotobutton.style.margin = "10px";

            // Add Goto button
            ((IHTMLDOMNode)popupDoc.body).appendChild((IHTMLDOMNode)gotobutton);

            // Add click event
            GotoElementButtonClick = new HtmlEvent();
            ((IHTMLElement2)gotobutton).SubscribeTo(EventType.onclick, GotoElementButtonClick);
            GotoElementButtonClick.OnEvent += (sender, e) => GotoElementButtonClick_OnEvent(sender, e, content.SMElementId);

          }

          // Edit Button
          if (content.AllowEdit)
          {

            var iconPath = Path.Combine(outPutDirectory, "Images\\Editor.png");
            string icon_path = new Uri(iconPath).LocalPath;

            var editBtn = popupDoc.createElement("<button>");
            editBtn.id = "edit-btn";

            editBtn.innerHTML = $"<span><img src='{icon_path}' width='16px' height='16px' margin='5px'></span>";
            editBtn.style.margin = "10px";

            // Create open button
            ((IHTMLDOMNode)popupDoc.body).appendChild((IHTMLDOMNode)editBtn);

            // Add click event
            EditButtonClick = new HtmlEvent();
            ((IHTMLElement2)editBtn).SubscribeTo(EventType.onclick, EditButtonClick);
            EditButtonClick.OnEvent += (sender, args) => EditButtonClick_OnEvent(sender, args, content.EditUrl);

          }

          if (content.AllowOpenInBrowser)
          {

            var iconPath = Path.Combine(outPutDirectory, "Images\\web.png");
            string icon_path = new Uri(iconPath).LocalPath;

            // Create browser button
            var browserBtn = popupDoc.createElement("<button>");
            browserBtn.id = "browser-btn";

            browserBtn.innerHTML = $"<span><img src='{icon_path}' width='16px' height='16px' margin='5px'></span>";
            browserBtn.style.margin = "10px";

            // Create open button
            ((IHTMLDOMNode)popupDoc.body).appendChild((IHTMLDOMNode)browserBtn);

            // Add click event
            PopupBrowserButtonClick = new HtmlEvent();
            ((IHTMLElement2)browserBtn).SubscribeTo(EventType.onclick, PopupBrowserButtonClick);
            PopupBrowserButtonClick.OnEvent += (sender, args) => PopupBrowserButtonClick_OnEvent(sender, args, content.BrowserQuery);

          }

          int maxheight = 400;
          int width = 300;
          int height = popup.GetOffsetHeight(width) + 10;

          if (height > maxheight)
          {
            height = maxheight;
            popupDoc.body.style.overflow = "scroll";
          }
          else
          {
            popupDoc.body.style.overflow = "";
          }

          popup.Show(x, y, width, height);

          // Popup links
          PopupWindowLinkClick = new HtmlEvent();
          ((IHTMLElement2)popupDoc.body).SubscribeTo(EventType.onclick, PopupWindowLinkClick);
          PopupWindowLinkClick.OnEvent += (sender, args) => PopupWindowLinkClick_OnEvent(sender, args, x, y, width, height, popup);

        }
        catch (RemotingException) { }
        catch (UnauthorizedAccessException) { }

      }));

    }

    private void EditButtonClick_OnEvent(object sender, IControlHtmlEventArgs e, string url)
    {

      if (url.IsNullOrEmpty())
        return;

      try
      {
        Process.Start(url);
      }
      catch (Exception) { }

    }

    private void PopupWindowLinkClick_OnEvent(object sender, IControlHtmlEventArgs obj, int x, int y, int w, int h, HtmlPopup popup)
    {

      var ev = obj.EventObj;


      var srcElement = ev.srcElement;
      if (srcElement.tagName.ToLower() != "a")
        return;

      var linkElement = srcElement as IHTMLAnchorElement;
      string url = linkElement?.href;
      if (url.IsNullOrEmpty())
        return;

      var provider = MatchContentProvider(url);
      if (provider.IsNull())
        return;

      bool ctrlPressed = ev.ctrlKey;
      Action action = null;

      if (ctrlPressed)
      {

        if (popup.IsNull())
          return;

        x = x + w;
        var doc = popup.GetDocument();
        var wdw = Application.Current.Dispatcher.Invoke(() => doc.parentWindow);

        action = () =>
        {
          // TODO: This will mess up the events
          OpenNewPopupWdw((IHTMLWindow4)wdw, url, provider, new RemoteCancellationToken(new CancellationToken()), x, y);
        };

      }
      else 
      {
        var wdw = Application.Current.Dispatcher.Invoke(() => ContentUtils.GetFocusedHtmlWindow());

        action = () =>
        {
          OpenNewPopupWdw((IHTMLWindow4)wdw, url, provider, new RemoteCancellationToken(new CancellationToken()), x, y);
        };
      }

      EventQueue.Enqueue(action);

    }

    private void GotoElementButtonClick_OnEvent(object sender, IControlHtmlEventArgs e, int elementId)
    {

      if (elementId < 0)
        return;

      if (Svc.SM.Registry.Element[elementId].IsNull())
        return;

      // Pass to the event queue to avoid deadlock issues
      Action action = () =>
      {

        if (!Svc.SM.UI.ElementWdw.GoToElement(elementId))
          LogTo.Warning($"Failed to GoToElement with id {elementId}");

      };

      EventQueue.Enqueue(action);

    }

    private void PopupBrowserButtonClick_OnEvent(object sender, IControlHtmlEventArgs e, string query)
    {

      if (_popupWindowSvc.IsNull())
        return;

      Action action = () =>
      {

        try
        {
          _popupWindowSvc.Open(query, ContentType.Article);
        }
        catch (RemotingException) { }

      };

      EventQueue.Enqueue(action);

    }

    private void ExtractButtonClick_OnEvent(object sender, IControlHtmlEventArgs e, PopupContent content, HtmlPopup popup)
    {

      ExtractType type = ExtractType.Full;

      if (popup.IsNull() || content.IsNull())
        return;

      var htmlDoc = popup.GetDocument();
      if (htmlDoc.IsNull())
        return;
      
      var sel = htmlDoc?.selection;
      var selObj = sel?.createRange() as IHTMLTxtRange;

      if (selObj.IsNull() || selObj.htmlText.IsNullOrEmpty())
        type = ExtractType.Full;
      else
        type = ExtractType.Partial;

      // Create an action to be passed to the event thread
      // Running on the main UI thread causes deadlock

      Action action = () => {

        // Extract the whole popup document
        if (type == ExtractType.Full)
          CreateSMExtract(content.Html, content.References);

        // Extract selected html
        else if (type == ExtractType.Partial)
          CreateSMExtract(selObj.htmlText, content.References);

      };

      EventQueue.Enqueue(action);


    }

    private void CreateSMExtract(string extract, References refs)
    {

      if (extract.IsNullOrEmpty())
      {
        LogTo.Error("Failed to CreateSMElement beacuse extract text was null");
        return;
      }

      var contents = new List<ContentBase>();
      contents.Add(new TextContent(true, extract));
      var currentElement = Svc.SM.UI.ElementWdw.CurrentElement;

      if (currentElement == null)
      {
        LogTo.Error("Failed to CreateSMElement beacuse element was null");
        return;
      }

      // Config.Defautl
      double priority = 30;

      bool ret = Svc.SM.Registry.Element.Add(
        out var value,
        ElemCreationFlags.ForceCreate,
        new ElementBuilder(ElementType.Topic, contents.ToArray())
          .WithParent(currentElement)
          .WithLayout("Article")
          .WithPriority(priority)
          .WithReference(r =>
            r.WithLink(refs.Link)
             .WithSource(refs.Source)
             .WithTitle(refs.Title)
          )
          .DoNotDisplay()
      );

      if (ret)
      {
        LogTo.Debug("Successfully created SM Element");
      }
      else
      {
        LogTo.Error("Failed to CreateSMElement");
      }
    }

    public bool RegisterProvider(string name, List<string> urlRegexes, IMouseoverContentProvider provider)
    {

      if (string.IsNullOrEmpty(name))
        return false;

      if (provider == null)
        return false;

      if (providers.ContainsKey(name))
        return false;

      providers[name] = new ContentProviderInfo(urlRegexes, provider);
      return true;

    }

    #endregion

    #region Methods
    #endregion
  }
}
