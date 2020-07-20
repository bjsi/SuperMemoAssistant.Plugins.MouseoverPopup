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
  using System.Linq;
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
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Builders;
  using SuperMemoAssistant.Interop.SuperMemo.Elements.Models;
  using SuperMemoAssistant.Plugins.MouseoverPopup.Models;
  using SuperMemoAssistant.Plugins.MouseoverPopup.UI;
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
    private MouseoverService _mouseoverSvc { get; set; }

    /// <summary>
    /// HtmlDocumentEvents
    /// TODO: Refactor
    /// </summary>
    private HTMLControlEvents htmlDocEvents { get; set; }

    // Popup window events
    private HtmlEvent ExtractButtonClick { get; set; }
    private HtmlEvent PopupBrowserButtonClick { get; set; }
    private HtmlEvent GotoElementButtonClick { get; set; }

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

      _mouseoverSvc = new MouseoverService();

      PublishService<IMouseoverSvc, MouseoverService>(_mouseoverSvc);

      SubscribeToMouseoverEvents();

      // Runs the actions added to the event queue
      _ = Task.Factory.StartNew(DispatchEvents, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);

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

    private void SubscribeToMouseoverEvents()
    {

      var events = new List<EventInitOptions>
      {
        new EventInitOptions(EventType.onmouseenter, _ => true, x => ((IHTMLElement)x).tagName.ToLower() == "a"),
      };

      htmlDocEvents = new HTMLControlEvents(events);
      htmlDocEvents.OnMouseEnterEvent += HtmlDocEvents_OnMouseOverEvent;
    }

    private async void HtmlDocEvents_OnMouseOverEvent(object sender, IHTMLControlEventArgs obj)
    {

      var ev = obj.EventObj;

      // Coordinates
      var x = ev.screenX;
      var y = ev.screenY;

      var linkElement = obj.EventObj.srcElement as IHTMLAnchorElement;
      string url = linkElement?.href;

      if (url.IsNullOrEmpty())
        return;

      if (providers.IsNull() || providers.Count == 0)
        return;

      // Find the matching provider
      foreach (var keyValuePair in providers)
      {
        var regexes = keyValuePair.Value.urlRegexes;
        var provider = keyValuePair.Value.provider;

        if (regexes.Any(r => new Regex(r).Match(url).Success))
        {
          await OpenNewPopupWdw(url, provider, new RemoteCancellationToken(new CancellationToken()), x, y);
          break;
        }
      }
    }

    private async Task OpenNewPopupWdw(string url, IContentProvider provider, RemoteCancellationToken ct, int x, int y)
    {

      if (url.IsNullOrEmpty() || provider.IsNull())
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

          var wdw = ContentUtils.GetFocusedHtmlWindow() as IHTMLWindow4;
          if (wdw.IsNull())
            return;

          var htmlDoc = ((IHTMLWindow2)wdw).document;
          if (htmlDoc.IsNull())
            return;

          var htmlCtrlBody = htmlDoc.body;
          if (htmlCtrlBody.IsNull())
            return;

          var popup = wdw.CreatePopup();
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

          // Extract Button
          if (content.AllowExtract)
          {

            // TODO: Add extract icon to button
            // Create Extract button
            var extractBtn = popupDoc.createElement("<button>");
            extractBtn.id = "extract-btn";
            extractBtn.innerText = "Extract";
            extractBtn.style.margin = "5px";

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

            // Create goto element button
            var gotobutton = popupDoc.createElement("<button>");
            gotobutton.id = "goto-btn";
            gotobutton.innerText = "Goto Element";
            gotobutton.style.margin = "5px";

            // Add Goto button
            ((IHTMLDOMNode)popupDoc.body).appendChild((IHTMLDOMNode)gotobutton);

            // Add click event
            GotoElementButtonClick = new HtmlEvent();
            ((IHTMLElement2)gotobutton).SubscribeTo(EventType.onclick, GotoElementButtonClick);
            GotoElementButtonClick.OnEvent += (sender, e) => GotoElementButtonClick_OnEvent(sender, e, content.SMElementId);

          }

          if (content.AllowOpenInBrowser)
          {

            // Create browser button
            var browserBtn = popupDoc.createElement("<button>");
            browserBtn.id = "browser-btn";
            browserBtn.innerText = "Popup Browser";
            browserBtn.style.margin = "5px";


            // Create open button
            ((IHTMLDOMNode)popupDoc.body).appendChild((IHTMLDOMNode)browserBtn);

            // Add click event
            PopupBrowserButtonClick = new HtmlEvent();
            ((IHTMLElement2)browserBtn).SubscribeTo(EventType.onclick, PopupBrowserButtonClick);
            PopupBrowserButtonClick.OnEvent += PopupBrowserButtonClick_OnEvent;

          }

          int maxheight = 400;
          int width = 300;
          int height = popup.GetOffsetHeight(width) + 10;

          if (height > maxheight)
          {
            height = maxheight;
            popupDoc.body.style.overflow = "scroll";
          }

          popup.Show(x, y, width, height);

        }
        catch (RemotingException) { }
        catch (UnauthorizedAccessException) { }

      }));

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

    private void PopupBrowserButtonClick_OnEvent(object sender, IControlHtmlEventArgs e)
    {
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

      if (!selObj.IsNull() && !selObj.htmlText.IsNullOrEmpty())
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

    public bool RegisterProvider(string name, List<string> urlRegexes, IContentProvider provider)
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
