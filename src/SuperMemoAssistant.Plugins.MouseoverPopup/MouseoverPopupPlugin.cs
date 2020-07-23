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
  using Ganss.Text;
  using global::MouseoverPopup.Interop;
  using mshtml;
  using SuperMemoAssistant.Extensions;
  using SuperMemoAssistant.Interop.Plugins;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Contents;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
  using SuperMemoAssistant.Interop.SuperMemo.Core;
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Builders;
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Models;
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Types;
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
    private Dictionary<string, ContentProvider> providers { get; set; } = new Dictionary<string, ContentProvider>();

    /// <summary>
    /// Service that providers can call to register themselves.
    /// </summary>
    private MouseoverService _mouseoverSvc = new MouseoverService();

    /// <summary>
    /// Service for communication between mouseover popup and popup window.
    /// </summary>
    private IPopupWindowSvc _popupWindowSvc { get; set; }

    // HtmlDoc events
    private HtmlEvent AnchorElementMouseLeave { get; set; }
    private HtmlEvent AnchorElementMouseEnter { get; set; }

    // Popup window events
    private HtmlEvent PopupAnchorElementClick { get; set; }
    private HtmlEvent PopupExtractButtonClick { get; set; }
    private HtmlEvent PopupBrowserButtonClick { get; set; }
    private HtmlEvent PopupGotoButtonClick { get; set; }
    private HtmlEvent PopupEditButtonClick { get; set; }

    // Used to run certain actions off of the UI thread
    // eg. item creation - otherwise will result in deadlock
    private EventfulConcurrentQueue<Action> EventQueue = new EventfulConcurrentQueue<Action>();

    public MouseoverPopupCfg Config;

    // True after Dispose is called
    private bool HasExited = false;

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

      Svc.SM.UI.ElementWdw.OnElementChanged += new ActionProxy<SMDisplayedElementChangedEventArgs>(OnElementChanged);

      // Start a new thread to handle events away from the UI thread.
      _ = Task.Factory.StartNew(DispatchEvents, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);

    }

    /// <summary>
    /// Runs the actions added to the event queue away from the UI thread.
    /// </summary>
    private void DispatchEvents()
    {

      while (!HasExited)
      {
        EventQueue.DataAvailableEvent.WaitOne(3000);
        while (EventQueue.TryDequeue(out var action))
        {
          try
          {
            action();
          }
          catch (RemotingException) { }
          catch (Exception e) 
          {
            LogTo.Error($"Exception {e} caught in event dispatcher thread");
          }
        }
      }
    }


    /// <summary>
    /// Register a provider that does NOT support keyword scanning.
    /// </summary>
    /// <param name="name"></param>
    /// <param name="urlRegexes"></param>
    /// <param name="provider"></param>
    /// <returns></returns>
    public bool RegisterProvider(string name, string[] urlRegexes, IMouseoverContentProvider provider)
    {

      if (string.IsNullOrEmpty(name))
      {

        LogTo.Warning("Failed to RegisterProvider because provider name was null or empty.");
        return false;

      }

      if (provider.IsNull())
      {

        LogTo.Warning("Failed to RegisterProvider because provider was null.");
        return false;

      }

      if (providers.ContainsKey(name))
      {

        LogTo.Warning($"Failed to RegisterProvider because provider with name {name} already exists.");
        return false;

      }

      providers[name] = new ContentProvider(urlRegexes, provider);
      return true;

    }

    /// <summary>
    /// Register a provider that DOES support keyword scanning.
    /// </summary>
    /// <param name="name"></param>
    /// <param name="urlRegexes"></param>
    /// <param name="keywordScanningOptions"></param>
    /// <param name="provider"></param>
    /// <returns></returns>
    public bool RegisterProvider(string name, string[] urlRegexes, KeywordScanningOptions keywordScanningOptions, IMouseoverContentProvider provider)
    {

      if (string.IsNullOrEmpty(name))
      {

        LogTo.Warning("Failed to RegisterProvider because provider name was null or empty");
        return false;

      }

      if (provider.IsNull())
      {

        LogTo.Warning("Failed to RegisterProvider because provider was null.");
        return false;

      }

      if (providers.ContainsKey(name))
      {

        LogTo.Warning($"Failed to RegisterProvider because provider with name {name} already exists");
        return false;

      }

      providers[name] = new ContentProvider(urlRegexes, keywordScanningOptions, provider);
      return true;

    }

    protected override void Dispose(bool disposing)
    {

      HasExited = true;
      base.Dispose(disposing);

    }

    private void OnElementChanged(SMDisplayedElementChangedEventArgs obj)
    {

      var matchingProviders = ProviderMatching.MatchProvidersAgainstCurrentElement(providers);
      if (matchingProviders.IsNull() || !matchingProviders.Any())
        return;

      var searchKeywords = Keywords.CreateKeywords(matchingProviders);
      if (searchKeywords.IsNull())
        return;

      // Adds links to matched keywords in the current element
      Keywords.ScanAndAddLinks(matchingProviders, searchKeywords);

      // Subscribe to mouseover and mouseleave events
      SubscribeToHtmlDocEvents(matchingProviders);

    }

    // TODO: Matching providers like this won't handle text changes well
    // Only refreshes on element changed events
    private void SubscribeToHtmlDocEvents(Dictionary<string, ContentProvider> matchedProviders)
    {

      var htmlDoc = ContentUtils.GetFocusedHtmlDocument();
      if (htmlDoc.IsNull())
        return;

      AnchorElementMouseEnter = new HtmlEvent();

      // Add mouseenter events to anchor elements
      var all = htmlDoc.body?.all as IHTMLElementCollection;
      all?.Cast<IHTMLElement>()
         ?.Where(x => x.tagName.ToLowerInvariant() == "a")
         ?.ForEach(x => ((IHTMLElement2)x).SubscribeTo(EventType.onmouseenter, AnchorElementMouseEnter));

      AnchorElementMouseEnter.OnEvent += (sender, args) => AnchorElementMouseEnterEvent(sender, args, matchedProviders);

    }

    private async void AnchorElementMouseEnterEvent(object sender, IControlHtmlEventArgs obj, Dictionary<string, ContentProvider> potentialProviders)
    {

      var ev = obj.EventObj;
      if (ev.IsNull())
        return;

      // Coordinates
      var x = ev.screenX;
      var y = ev.screenY;

      var anchor = obj.EventObj.srcElement as IHTMLAnchorElement;
      string url = anchor?.href;
      string innerText = ((IHTMLElement)anchor)?.innerText;
      if (url.IsNullOrEmpty() || anchor.IsNull() || innerText.IsNullOrEmpty())
        return;

      var matchedProviders = ProviderMatching.MatchProvidersAgainstMouseoverLink(url, innerText, potentialProviders);
      if (matchedProviders.IsNull() || !matchedProviders.Any())
        return;

      var parentWdw = Application.Current.Dispatcher.Invoke(() => ContentUtils.GetFocusedHtmlWindow());
      if (parentWdw.IsNull())
        return;

      var linkElement = anchor as IHTMLElement2;
      if (linkElement.IsNull())
        return;

      // Add mouseleave event to cancel early
      var remoteCancellationTokenSource = new RemoteCancellationTokenSource();
      CancelTaskEarlyOnMouseLeave(linkElement, remoteCancellationTokenSource);

      try
      {

        // Open a menu to choose a provider if multiple available
        if (matchedProviders.Count > 1)
        {
          await OpenChooseProviderMenu((IHTMLWindow4)parentWdw, url, innerText, matchedProviders, x, y);
        }

        // Directly open a popup window
        else
        {
          await OpenNewPopupWdw((IHTMLWindow4)parentWdw, url, matchedProviders.First().Value.provider, remoteCancellationTokenSource.Token, x, y);
        }

      }
      catch (RemotingException) { }

    }

    /// <summary>
    /// Cancel the oustanding Task on mouseleave.
    /// </summary>
    /// <param name="element"></param>
    /// <param name="remoteCancellationTokenSource"></param>
    private void CancelTaskEarlyOnMouseLeave(IHTMLElement2 element, RemoteCancellationTokenSource remoteCancellationTokenSource)
    {

      if (element.IsNull())
        return;

      AnchorElementMouseLeave = new HtmlEvent();
      element.SubscribeTo(EventType.onmouseleave, AnchorElementMouseLeave);
      AnchorElementMouseLeave.OnEvent += (sender, args) => remoteCancellationTokenSource?.Cancel();

    }

    /// <summary>
    /// If multiple providers match, open a selection menu so the user can pick one.
    /// </summary>
    /// <param name="parentWdw"></param>
    /// <param name="url"></param>
    /// <param name="text"></param>
    /// <param name="providers"></param>
    /// <param name="x"></param>
    /// <param name="y"></param>
    /// <returns></returns>
    private async Task OpenChooseProviderMenu(IHTMLWindow4 parentWdw, string url, string text, Dictionary<string, ContentProvider> providers, int x, int y)
    {

      if (url.IsNullOrEmpty())
        return;

      await Application.Current.Dispatcher.BeginInvoke((Action)(() =>
      {

        var htmlDoc = ((IHTMLWindow2)parentWdw).document;
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

        string choices = string.Empty;

        foreach (var provider in providers)
        {
          choices += $"<li><a href='{provider.Value.keywordScanningOptions.urlKeywordMap[text.ToLowerInvariant()]}'>{text} ({provider.Key})</li>";
        }

        popupDoc.body.innerHTML = $@"
            <html>
              <body>
                <h1>Mouseover Popup Menu</h1>
                <ul>
                  {choices}
                </ul>
              </body>
            </html>";

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
        PopupAnchorElementClick = new HtmlEvent();
        ((IHTMLElement2)popupDoc.body).SubscribeTo(EventType.onclick, PopupAnchorElementClick);
        PopupAnchorElementClick.OnEvent += (sender, args) => PopupWindowLinkClick_OnEvent(sender, args, x, y, width, height, popup);

      }));

    }

    /// <summary>
    /// Open a new popup mouseover popup.
    /// </summary>
    /// <param name="parentWdw"></param>
    /// <param name="url"></param>
    /// <param name="provider"></param>
    /// <param name="ct"></param>
    /// <param name="x"></param>
    /// <param name="y"></param>
    /// <returns></returns>
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
            PopupExtractButtonClick = new HtmlEvent();
            ((IHTMLElement2)extractBtn).SubscribeTo(EventType.onclick, PopupExtractButtonClick);
            PopupExtractButtonClick.OnEvent += (sender, e) => ExtractButtonClick_OnEvent(sender, e, content, popup);

          }

          // Goto Element Button
          if (content.SMElementId > -1)
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
            PopupGotoButtonClick = new HtmlEvent();
            ((IHTMLElement2)gotobutton).SubscribeTo(EventType.onclick, PopupGotoButtonClick);
            PopupGotoButtonClick.OnEvent += (sender, e) => GotoElementButtonClick_OnEvent(sender, e, content.SMElementId);

          }

          // Edit Button
          if (!content.EditUrl.IsNullOrEmpty())
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
            PopupEditButtonClick = new HtmlEvent();
            ((IHTMLElement2)editBtn).SubscribeTo(EventType.onclick, PopupEditButtonClick);
            PopupEditButtonClick.OnEvent += (sender, args) => EditButtonClick_OnEvent(sender, args, content.EditUrl);

          }

          if (!content.BrowserQuery.IsNullOrEmpty())
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
          PopupAnchorElementClick = new HtmlEvent();
          ((IHTMLElement2)popupDoc.body).SubscribeTo(EventType.onclick, PopupAnchorElementClick);
          PopupAnchorElementClick.OnEvent += (sender, args) => PopupWindowLinkClick_OnEvent(sender, args, x, y, width, height, popup);

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
      if (srcElement.tagName.ToLowerInvariant() != "a")
        return;

      var linkElement = srcElement as IHTMLAnchorElement;
      string url = linkElement?.href;
      if (url.IsNullOrEmpty())
        return;

      string innerText = ((IHTMLElement)linkElement).innerText;

      var matchedProviders = ProviderMatching.MatchProvidersAgainstMouseoverLink(url, innerText, providers);
      if (matchedProviders.IsNull())
        return;

      bool ctrlPressed = ev.ctrlKey;
      Action action = null;

      if (ctrlPressed)
      {

        if (popup.IsNull())
          return;

        x += w;
        var doc = popup.GetDocument();
        var wdw = Application.Current.Dispatcher.Invoke(() => doc.parentWindow);

        if (matchedProviders.Count == 1)
        {

          action = () =>
          {
            OpenNewPopupWdw((IHTMLWindow4)wdw, url, matchedProviders.First().Value.provider, new RemoteCancellationToken(new CancellationToken()), x, y);
          };

        }
        else
        {

          action = () =>
          {
            OpenChooseProviderMenu((IHTMLWindow4)wdw, url, innerText, matchedProviders, x, y);
          };

        }
      }
      else 
      {
        var wdw = Application.Current.Dispatcher.Invoke(() => ContentUtils.GetFocusedHtmlWindow());

        if (matchedProviders.Count == 1)
        {

          action = () =>
          {
            OpenNewPopupWdw((IHTMLWindow4)wdw, url, matchedProviders.First().Value.provider, new RemoteCancellationToken(new CancellationToken()), x, y);
          };

        }
        else
        {

          action = () =>
          {
            OpenChooseProviderMenu((IHTMLWindow4)wdw, url, innerText, matchedProviders, x, y);
          };

        }

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

      // Open inside popupWindow if available
      if (!_popupWindowSvc.IsNull())
      {

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

      // Open in the user's default browser
      else
      {
        try
        {
          Process.Start(query);
        }
        catch (Exception) { }
      }

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

      bool ret = Svc.SM.Registry.Element.Add(
        out var value,
        ElemCreationFlags.ForceCreate,
        new ElementBuilder(ElementType.Topic, contents.ToArray())
          .WithParent(currentElement)
          .WithLayout("Article")
          .WithPriority(Config.DefaultPriority)
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


    #endregion

    #region Methods
    #endregion
  }
}
