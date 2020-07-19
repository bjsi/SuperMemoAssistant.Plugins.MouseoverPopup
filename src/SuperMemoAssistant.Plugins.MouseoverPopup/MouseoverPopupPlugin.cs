﻿#region License & Metadata

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
    private HTMLControlEvents htmlDocEvents { get; set; }
    private PopupWdw CurrentWdw { get; set; } = null;

    private HtmlEvent ExtractButtonClick { get; set; }
    private HtmlEvent PopupBrowserButtonClick { get; set; }
    private HtmlEvent GotoElementButtonClick { get; set; }


    private IHTMLPopup CurrentPopup { get; set; }

    public MouseoverPopupCfg Config;
    public PopupContent CurrentContent { get; set; }

    private EventfulConcurrentQueue<Action> EventQueue = new EventfulConcurrentQueue<Action>();

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

    private async void HtmlDocEvents_OnMouseOverEvent(object sender, IHTMLControlEventArgs e)
    {
      var ev = e.EventObj;
      var x = ev.clientX;
      var y = ev.clientY;

      var link = e.EventObj.srcElement as IHTMLAnchorElement;
      string url = link?.href;
      if (url.IsNullOrEmpty())
        return;

      if (CurrentWdw != null && !CurrentWdw.IsClosed)
        return;

      foreach (var keyValuePair in providers)
      {
        var regexes = keyValuePair.Value.urlRegexes;
        var provider = keyValuePair.Value.provider;

        if (regexes.Any(r => new Regex(r).Match(url).Success))
        {
          await OpenNewPopupWdw(url, provider, new RemoteCancellationToken(new CancellationToken()), x, y);
        }
      }
    }

    private async Task OpenNewPopupWdw(string url, IContentProvider provider, RemoteCancellationToken ct, int x, int y)
    {

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

          var body = ((IHTMLWindow2)wdw).document?.body;
          CurrentPopup = wdw?.createPopup() as IHTMLPopup;
          var doc = ((IHTMLDocument2)CurrentPopup?.document);
          if (doc.IsNull())
            return;

          doc.body.style.border = "solid black 1px";
          doc.body.style.overflow = "scroll";
          doc.body.style.margin = "5px";
          doc.body.innerHTML = content.Html;

          if (content.AllowExtract)
          {

            // TODO: Add extract icon to button
            // Create Extract button
            var extractBtn = doc.createElement("<button>");
            extractBtn.id = "extract-btn";
            extractBtn.innerText = "Extract";
            extractBtn.style.margin = "5px";

            // Add extract button
            ((IHTMLDOMNode)doc.body).appendChild((IHTMLDOMNode)extractBtn);

            // Add click event
            ExtractButtonClick = new HtmlEvent();
            ((IHTMLElement2)extractBtn).SubscribeTo(EventType.onclick, ExtractButtonClick);
            ExtractButtonClick.OnEvent += ExtractButtonClick_OnEvent;

          }

          if (content.AllowGotoInSM)
          {

            // Create goto element button
            var gotobutton = doc.createElement("<button>");
            gotobutton.id = "goto-btn";
            gotobutton.innerText = "Goto Element";
            gotobutton.style.margin = "5px";

            // Add Goto button
            ((IHTMLDOMNode)doc.body).appendChild((IHTMLDOMNode)gotobutton);

            // Add click event
            GotoElementButtonClick = new HtmlEvent();
            ((IHTMLElement2)gotobutton).SubscribeTo(EventType.onclick, GotoElementButtonClick);
            GotoElementButtonClick.OnEvent += (sender, e) => GotoElementButtonClick_OnEvent(sender, e, content.SMElementId);

          }

          if (content.AllowOpenInBrowser)
          {

            // Create browser button
            var browserBtn = doc.createElement("<button>");
            browserBtn.id = "browser-btn";
            browserBtn.innerText = "Popup Browser";
            browserBtn.style.margin = "5px";

            // Create open button
            ((IHTMLDOMNode)doc.body).appendChild((IHTMLDOMNode)browserBtn);

            // Add click event
            PopupBrowserButtonClick = new HtmlEvent();
            ((IHTMLElement2)browserBtn).SubscribeTo(EventType.onclick, PopupBrowserButtonClick);
            PopupBrowserButtonClick.OnEvent += PopupBrowserButtonClick_OnEvent;

          }

          // TODO: How to size to content?
          CurrentContent = content;
          CurrentPopup.Show(x, y, 350, 350, body);

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

    private void ExtractButtonClick_OnEvent(object sender, IControlHtmlEventArgs e)
    {
      Action action = () => {

        if (CurrentPopup.IsNull())
          return;

        var htmlDoc = CurrentPopup?.document as IHTMLDocument2;

        var sel = htmlDoc?.selection;
        var selObj = sel?.createRange() as IHTMLTxtRange;
        if (selObj.IsNull() || selObj.text.IsNullOrEmpty())
        {
          // Extract the whole popup document
          CreateSMExtract(CurrentContent.Html);
        }
        else
        {
          // Extract selected html
          CreateSMExtract(selObj.htmlText);
        }

      };

      EventQueue.Enqueue(action);


    }

    private void CreateSMExtract(string extract)
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
            r.WithLink(CurrentContent.References.Link)
             .WithSource(CurrentContent.References.Source)
             .WithTitle(CurrentContent.References.Title)
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
