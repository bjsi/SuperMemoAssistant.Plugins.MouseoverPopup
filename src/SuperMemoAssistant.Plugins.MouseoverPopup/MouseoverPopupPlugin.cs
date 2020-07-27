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
      if (!matchingProviders.IsNull() && matchingProviders.Any())
      {

        // Create keyword trie
        var searchKeywords = Keywords.CreateKeywords(matchingProviders);
        if (searchKeywords.IsNull())
          return;

        // Adds links to matched keywords in the current element
        Keywords.ScanAndAddLinks(matchingProviders, searchKeywords);

      }

      // Subscribe to mouseover and mouseleave events
      // TODO: Pass all providers or just matching?
      SubscribeToHtmlDocEvents(providers);

    }

    private void SubscribeToHtmlDocEvents(Dictionary<string, ContentProvider> matchedProviders)
    {

      var htmlDoc = ContentUtils.GetFocusedHtmlDocument();
      if (htmlDoc.IsNull())
        return;

      AnchorElementMouseEnter = new HtmlEvent();

      // Add mouseenter events to links
      var links = htmlDoc.links;
      if (links.IsNull())
        return;

      foreach (IHTMLElement2 link in links)
      {
         link.SubscribeTo(EventType.onmouseenter, AnchorElementMouseEnter);
      }

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

          Action action = () =>
          {
            OpenChooseProviderMenu((IHTMLWindow4)parentWdw, url, innerText, matchedProviders, x, y);
          };

          EventQueue.Enqueue(action);

        }

        // Directly open a popup window
        else
        {

          Action action = () =>
          {
            OpenNewPopupWdw((IHTMLWindow4)parentWdw, url, innerText, matchedProviders.First().Value, remoteCancellationTokenSource.Token, x, y);
          };

          EventQueue.Enqueue(action);

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
        if (popup.IsNull())
          return;
        popup.OnShow += (sender, args) => _mouseoverSvc?.InvokeOnShow(sender, args);

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

      }));

    }

    /// <summary>
    /// Open a new popup mouseover popup.
    /// </summary>
    /// <returns></returns>
    private async Task OpenNewPopupWdw(IHTMLWindow4 parentWdw, string url, string innerText, ContentProvider providerInfo, RemoteCancellationToken ct, int x, int y)
    {

      if (url.IsNullOrEmpty() || parentWdw.IsNull() || innerText.IsNullOrEmpty())
        return;

      if (providerInfo.IsNull() || providerInfo.provider.IsNull())
        return;

      var provider = providerInfo.provider;

      // Check url or keywords

      PopupContent content = null;
      if (providerInfo.urlRegexes.Any(r => new Regex(r).Match(url).Success))
      {
        content = await provider.FetchHtml(ct, url);
      }
      else
      {
        if (providerInfo.keywordScanningOptions.urlKeywordMap.TryGetValue(innerText, out var href))
        {
          content = await provider.FetchHtml(ct, href);
        }
        else
          return;
      }

      if (content.IsNull() || content.Html.IsNullOrEmpty())
        return;

      await Application.Current.Dispatcher.BeginInvoke((Action)(() =>
      {
        try
        {

          var htmlDoc = ((IHTMLWindow2)parentWdw).document;
          if (htmlDoc.IsNull())
            return;

          var popup = parentWdw.CreatePopup();
          if (popup.IsNull())
            return;

          popup.OnShow += (sender, args) => _mouseoverSvc?.InvokeOnShow(sender, args);
          popup.AddContent(content.Html);

          // Extract Button
          if (content.AllowExtract)
          {

            popup.AddExtractButton();
            popup.OnExtractButtonClick += (sender, args) => ExtractButtonClick_OnEvent(sender, args, content, popup);

          }

          // SM Element Goto Button
          if (content.SMElementId > -1)
          {

            popup.AddGotoButton();
            popup.OnGotoButtonClick += (sender, args) => GotoElementButtonClick_OnEvent(sender, args, content.SMElementId);

          }

          // Edit Button
          if (!content.EditUrl.IsNullOrEmpty())
          {

            popup.AddEditButton();
            popup.OnEditButtonClick += (sender, args) => EditButtonClick_OnEvent(sender, args, content.EditUrl);

          }

          // Browser Button
          if (!content.BrowserQuery.IsNullOrEmpty())
          {

            popup.AddBrowserButton();
            popup.OnBrowserButtonClick += (sender, args) => PopupBrowserButtonClick_OnEvent(sender, args, content.BrowserQuery);

          }

          int width = 300;
          int height = CalculatePopupHeight(popup, width);

          // Link Click events
          popup.OnLinkClick += (sender, args) => PopupWindowLinkClick_OnEvent(sender, args, x, y, width, height, popup);

          popup.Show(x, y, width, height);

        }
        catch (RemotingException) { }
        catch (UnauthorizedAccessException) { }
        catch (Exception ex)
        {
          LogTo.Error($"Exception {ex} while opening new popup window");
        }

      }));

    }

    private int CalculatePopupHeight(HtmlPopup popup, int width)
    {

      if (popup.IsNull())
        return -1;

      var popupDoc = popup.GetDocument();

      int maxheight = 500;
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

      return height;

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

    private async void PopupWindowLinkClick_OnEvent(object sender, IControlHtmlEventArgs obj, int x, int y, int w, int h, HtmlPopup popup)
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

            OpenNewPopupWdw((IHTMLWindow4)wdw, url, innerText, matchedProviders.First().Value, new RemoteCancellationToken(new CancellationToken()), x, y);
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
            OpenNewPopupWdw((IHTMLWindow4)wdw, url, innerText, matchedProviders.First().Value, new RemoteCancellationToken(new CancellationToken()), x, y);
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
