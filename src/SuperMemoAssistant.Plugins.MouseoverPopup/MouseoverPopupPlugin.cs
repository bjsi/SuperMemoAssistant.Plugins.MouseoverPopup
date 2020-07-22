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
    private Dictionary<string, ContentProviderInfo> providers { get; set; } = new Dictionary<string, ContentProviderInfo>();

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

    // Fast keyword search data structure
    private AhoCorasick Keywords { get; set; }

    private const int MaxTextLength = 2000000000;

    public MouseoverPopupCfg Config;

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

      // Runs the actions added to the event queue
      _ = Task.Factory.StartNew(DispatchEvents, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);

    }

    protected override void Dispose(bool disposing)
    {

      HasExited = true;
      base.Dispose(disposing);

    }

    private void OnElementChanged(SMDisplayedElementChangedEventArgs obj)
    {

      var matchingProviders = MatchProvidersAgainstCurrentElement();
      if (matchingProviders.IsNull() || !matchingProviders.Any())
        return;

      CreateKeywords(matchingProviders);

      ScanForKeywords(matchingProviders);

      SubscribeToHtmlDocEvents(matchingProviders);

    }

    private Dictionary<string, ContentProviderInfo> MatchProvidersAgainstCurrentElement()
    {

      var ret = new Dictionary<string, ContentProviderInfo>();

      if (providers.IsNull() || !providers.Any())
        return ret;

      foreach (var provider in providers)
      {

        if (MatchProviderAgainstCurrentElement(provider.Value))
          ret.Add(provider.Key, provider.Value);

      }

      return ret;

    }

    private void SubscribeToHtmlDocEvents(Dictionary<string, ContentProviderInfo> matchedProviders)
    {

      var htmlDoc = ContentUtils.GetFocusedHtmlDocument();
      if (htmlDoc.IsNull())
        return;

      MouseEnterEvent = new HtmlEvent();

      var all = htmlDoc.body?.all as IHTMLElementCollection;
      all?.Cast<IHTMLElement>()
         ?.Where(x => x.tagName.ToLowerInvariant() == "a")
         ?.ForEach(x => ((IHTMLElement2)x).SubscribeTo(EventType.onmouseenter, MouseEnterEvent));


      // TODO: Matching providers like this won't handle text changes well
      // Only refreshes on element changed events

      MouseEnterEvent.OnEvent += (sender, args) => MouseEnterEvent_OnEvent(sender, args, matchedProviders);

    }

    private bool MatchAgainstCurrentReferences(ReferenceRegexes refRegexes)
    {

      var htmlCtrl = ContentUtils.GetFirstHtmlCtrl();
      string text = htmlCtrl?.Text;
      if (text.IsNullOrEmpty())
        return false;

      var refs = ReferenceParser.GetReferences(htmlCtrl?.Text);
      if (refs.IsNull())
        return false;

      else if (MatchAgainstRegexes(refs.Source, refRegexes.SourceRegexes))
        return true;

      else if (MatchAgainstRegexes(refs.Link, refRegexes.LinkRegexes))
        return true;

      if (MatchAgainstRegexes(refs.Title, refRegexes.TitleRegexes))
        return true;

      else if (MatchAgainstRegexes(refs.Author, refRegexes.AuthorRegexes))
        return true;

      return false;

    }

    private bool MatchAgainstRegexes(string input, string[] regexes)
    {

      if (input.IsNullOrEmpty())
        return false;

      if (regexes.IsNull() || !regexes.Any())
        return false;

      if (regexes.Any(r => new Regex(r).Match(input).Success))
        return true;

      return false;

    }

    private bool MatchAgainstCategoryPath(IElement element, string[] regexes)
    {

      if (element.IsNull())
        return false;

      var cur = element.Parent;
      while (!cur.IsNull())
      {
        if (cur.Type == ElementType.ConceptGroup)
        {

          // TODO: Check that this works
          var concept = Svc.SM.Registry.Concept[cur.Id];
          string name = concept.Name;

          if (!concept.IsNull() && regexes.Any(x => new Regex(x).Match(name).Success))
            return true;

        }
        cur = cur.Parent;
      }

      return false;

    }

    private bool MatchProviderAgainstCurrentElement(ContentProviderInfo providerInfo)
    {

      var element = Svc.SM.UI.ElementWdw.CurrentElement;
      if (element.IsNull())
        return false;

      var referenceRegexes = providerInfo.referenceRegexes;
      var categoryPathRegexes = providerInfo.CategoryPathRegexes;

      return MatchAgainstCategoryPath(element, categoryPathRegexes) || MatchAgainstCurrentReferences(referenceRegexes);

    }

    private void ScanForKeywords(Dictionary<string, ContentProviderInfo> providers)
    {

      var htmlCtrls = ContentUtils.GetHtmlCtrls();
      if (htmlCtrls.IsNull() || !htmlCtrls.Any())
        return;

      foreach (KeyValuePair<int, IControlHtml> kvpair in htmlCtrls)
      {

        var htmlCtrl = kvpair.Value;
        var text = htmlCtrl?.Text?.ToLowerInvariant();
        var htmlDoc = htmlCtrl?.GetDocument();
        if (text.IsNullOrEmpty() || htmlDoc.IsNull())
          continue;

        var matches = Keywords.Search(text);
        if (!matches.Any())
          continue;

        var orderedMatches = matches.OrderBy(x => x.Index);
        var selObj = htmlDoc.selection?.createRange() as IHTMLTxtRange;
        if (selObj.IsNull())
          continue;

        foreach (var match in orderedMatches)
        {

          string word = match.Word;
          if (selObj.findText(word, Flags: 2)) // Match whole words only TODO: Test this
          {

            var parentEl = selObj.parentElement();
            if (!parentEl.IsNull())
            {
              if (parentEl.tagName.ToLowerInvariant() == "a")
                continue;
            }
            else
            {

              string href = null;
              foreach (var provider in providers)
              {
                if (provider.Value.keywordUrlMap.TryGetValue(word, out href))
                  break;
              }

              if (href.IsNull())
                return;

              //selObj.pasteHTML($"<a href='{href}'>{selObj.text}<a>");

            }

          }

          selObj.collapse(false);
          selObj.moveEnd("character", MaxTextLength);

        }
      }
    }

    private void CreateKeywords(Dictionary<string, ContentProviderInfo> providers)
    {

      Keywords = new AhoCorasick();

      foreach (var provider in providers)
      {
        
        var words = provider.Value.keywordUrlMap?.Keys;
        if (words.IsNull() || !words.Any())
          continue;

        Keywords.Add(words);

      }

    }

    private async void MouseEnterEvent_OnEvent(object sender, IControlHtmlEventArgs obj, Dictionary<string, ContentProviderInfo> potentialProviders)
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

      string innerText = ((IHTMLElement)anchor).innerText;

      var matchedProviders = MatchProvidersAgainstMouseoverLink(url, innerText, potentialProviders);
      if (matchedProviders.IsNull() || !matchedProviders.Any())
        return;

      var parentWdw = Application.Current.Dispatcher.Invoke(() => ContentUtils.GetFocusedHtmlWindow());
      if (parentWdw.IsNull())
        return;

      var linkElement = anchor as IHTMLElement2;
      if (linkElement.IsNull())
        return;

      // Add mouseleave event to cancel early
      var remoteCancellationToken = new RemoteCancellationTokenSource();
      AnchorElementMouseLeave = new HtmlEvent();
      linkElement.SubscribeTo(EventType.onmouseleave, AnchorElementMouseLeave);
      AnchorElementMouseLeave.OnEvent += (sender, args) => remoteCancellationToken?.Cancel();

      try
      {

        if (matchedProviders.Count > 1)
          await OpenChooseProviderMenu((IHTMLWindow4)parentWdw, url, innerText, matchedProviders, x, y);
        else
          await OpenNewPopupWdw((IHTMLWindow4)parentWdw, url, matchedProviders.First().Value.provider, remoteCancellationToken.Token, x, y);

      }
      catch (RemotingException) { }

    }

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

    private Dictionary<string, ContentProviderInfo> MatchProvidersAgainstMouseoverLink(string url, string text, Dictionary<string, ContentProviderInfo> potentialProviders)
    {

      if (url.IsNullOrEmpty() || text.IsNullOrEmpty() || potentialProviders.IsNull() || !potentialProviders.Any())
        return null;

      var ret = new Dictionary<string, ContentProviderInfo>();

      foreach (var provider in potentialProviders)
      {

        var regexes = provider.Value.urlRegexes;

        //var provider = keyValuePair.Value.provider;

        if (regexes.Any(r => new Regex(r).Match(url).Success)
         || provider.Value.keywordUrlMap.Keys.Any(x => x == text))
        {
          ret.Add(provider.Key, provider.Value);
        }

      }

      return null;

    }

    private async Task OpenChooseProviderMenu(IHTMLWindow4 parentWdw, string url, string text, Dictionary<string, ContentProviderInfo> providers, int x, int y)
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
          choices += $"<li><a href='{provider.Value.keywordUrlMap[text.ToLowerInvariant()]}'>{text} ({provider.Key})</li>";
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
        PopupWindowLinkClick = new HtmlEvent();
        ((IHTMLElement2)popupDoc.body).SubscribeTo(EventType.onclick, PopupWindowLinkClick);
        PopupWindowLinkClick.OnEvent += (sender, args) => PopupWindowLinkClick_OnEvent(sender, args, x, y, width, height, popup);

      }));

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
      if (srcElement.tagName.ToLowerInvariant() != "a")
        return;

      var linkElement = srcElement as IHTMLAnchorElement;
      string url = linkElement?.href;
      if (url.IsNullOrEmpty())
        return;

      string innerText = ((IHTMLElement)linkElement).innerText;

      var matchedProviders = MatchProvidersAgainstMouseoverLink(url, innerText, providers);
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

    public bool RegisterProvider(string name, List<string> urlRegexes, Dictionary<string, string> keywordUrlMap, ReferenceRegexes referenceRegexes, string[] categoryPathRegexes, IMouseoverContentProvider provider)
    {

      if (string.IsNullOrEmpty(name))
        return false;

      if (provider == null)
        return false;

      if (providers.ContainsKey(name))
        return false;

      providers[name] = new ContentProviderInfo(urlRegexes, keywordUrlMap, referenceRegexes, categoryPathRegexes, provider);
      return true;

    }

    #endregion

    #region Methods
    #endregion
  }
}
