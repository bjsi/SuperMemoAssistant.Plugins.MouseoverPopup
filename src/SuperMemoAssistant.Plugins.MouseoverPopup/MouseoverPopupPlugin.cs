using Anotar.Serilog;
using MouseoverPopupInterfaces;
using mshtml;
using SuperMemoAssistant.Interop.SuperMemo.Content.Contents;
using SuperMemoAssistant.Interop.SuperMemo.Core;
using SuperMemoAssistant.Interop.SuperMemo.Elements.Builders;
using SuperMemoAssistant.Interop.SuperMemo.Elements.Models;
using SuperMemoAssistant.Plugins.MouseoverPopup.Models;
using SuperMemoAssistant.Services;
using SuperMemoAssistant.Services.IO.HotKeys;
using SuperMemoAssistant.Services.Sentry;
using SuperMemoAssistant.Services.UI.Configuration;
using SuperMemoAssistant.Sys.Remoting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

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
// Created On:   4/29/2021 5:43:15 PM
// Modified By:  james

#endregion




namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
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
    private MouseoverService MouseoverSvc { get; set; } = new MouseoverService();

    // HtmlDoc events
    private HtmlEvent AnchorElementMouseLeave { get; set; }
    private HtmlEvent AnchorElementMouseEnter { get; set; }

    // Used to run certain actions off of the UI thread
    // eg. item creation - otherwise will result in deadlock
    private EventfulConcurrentQueue<Action> JobQueue = new EventfulConcurrentQueue<Action>();

    public MouseoverPopupCfg Config { get; private set; }

    // True after Dispose is called
    private bool HasExited = false;

    #endregion




    #region Methods Impl

    private void LoadConfig()
    {
      Config = Svc.Configuration.Load<MouseoverPopupCfg>() ?? new MouseoverPopupCfg();
    }



    /// <inheritdoc />
    protected override void OnSMStarted(bool wasSMAlreadyStarted)
    {
      LoadConfig();

      PublishService<IMouseoverSvc, MouseoverService>(MouseoverSvc);

      Svc.SM.UI.ElementWdw.OnElementChanged += new ActionProxy<SMDisplayedElementChangedEventArgs>(OnElementChanged);

      // Start a new thread to handle events away from the UI thread.
      _ = Task.Factory.StartNew(HandleJobs, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);

      base.OnSMStarted(wasSMAlreadyStarted);
    }

    #endregion


    #region Methods


    // TODO: Dispose.

    /// <summary>
    /// Runs the actions added to the event queue away from the UI thread.
    /// </summary>
    private void HandleJobs()
    {

      while (!HasExited)
      {
        JobQueue.DataAvailableEvent.WaitOne(3000);
        while (JobQueue.TryDequeue(out var action))
        {
          try
          {
            action();
          }
          catch (RemotingException) { }
          catch (Exception e)
          {
            LogTo.Error($"Exception {e} caught in job queue thread");
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
      LogTo.Debug("Recieved call to RegisterProvider from: " + name);

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

    private void OnElementChanged(SMDisplayedElementChangedEventArgs obj)
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

      AnchorElementMouseEnter.OnEvent += (sender, args) => AnchorElementMouseEnterEvent(sender, args, providers);

    }

    private async void AnchorElementMouseEnterEvent(object sender, IControlHtmlEventArgs obj, Dictionary<string, ContentProvider> potentialProviders)
    {

      // Get data from the IHTMLEventObj / IHTMLElements

      var ev = obj.EventObj;
      if (ev.IsNull())
        return;

      bool ctrlPressed = ev.ctrlKey;
      if (Config.RequireCtrlKey && !ctrlPressed)
        return;

      // Coordinates
      var x = ev.screenX;
      var y = ev.screenY;

      var anchor = obj.EventObj.srcElement as IHTMLAnchorElement;
      string url = anchor?.href;
      string innerText = ((IHTMLElement)anchor)?.innerText;
      if (url.IsNullOrEmpty() || anchor.IsNull() || innerText.IsNullOrEmpty())
        return;

      try
      {

        Action action = () =>
        {

          var matchedProvider = ProviderMatching.MatchProvidersAgainstMouseoverLink(url, innerText, potentialProviders);
          if (matchedProvider.IsNull())
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
          OpenNewPopupWdw((IHTMLWindow4)parentWdw, url, innerText, matchedProvider, remoteCancellationTokenSource.Token, x, y);

        };

        JobQueue.Enqueue(action);

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

      // Check whether url or keywords matched and fetch the corresponding content
      // Time the response and add some delay if necessary.
      // This prevents many popups opening when you move the mouse over many links
      // It also makes the interval between mouseover and popup opening more consistent

      var start = DateTime.Now;
      PopupContent content = null;
      if (providerInfo.urlRegexes.Any(r => new Regex(r).Match(url).Success))
      {
        content = await provider.FetchHtml(ct, url);
      }

      // Add delay if necessary

      var responseTime = DateTime.Now - start;
      if (responseTime.TotalMilliseconds < 400)
      {
        await Task.Delay(TimeSpan.FromMilliseconds(400 - responseTime.TotalMilliseconds)).ConfigureAwait(false);
      }

      if (ct.Token().IsCancellationRequested)
        return;

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

          popup.AddContent(content.Html);

          // Extract Button
          if (content.AllowExtract)
          {

            popup.AddExtractButton();
            popup.OnExtractButtonClick += (sender, args) => ExtractButtonClick_OnEvent(sender, args, content, popup);

          }

          // SM Element Goto Button
          //if (content.SMElementId > -1)
          //{

          //  popup.AddGotoButton();
          //  popup.OnGotoButtonClick += (sender, args) => GotoElementButtonClick_OnEvent(sender, args, content.SMElementId);

          //}

          // Browser Button
          if (!content.BrowserQuery.IsNullOrEmpty())
          {

            popup.AddBrowserButton();
            popup.OnBrowserButtonClick += (sender, args) => PopupBrowserButtonClick_OnEvent(sender, args, content.BrowserQuery);

          }

          int width = 400;
          int height = CalculatePopupHeight(popup, width);

          if (ct.Token().IsCancellationRequested)
            return;

          // Link Click events

          var opts = new HtmlPopupOptions(x, y, width, height);
          popup.OnLinkClick += (sender, args) => PopupWindowLinkClick_OnEvent(sender, args, opts, popup);
          popup.Show(opts);

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

    private async void PopupWindowLinkClick_OnEvent(object sender, IControlHtmlEventArgs obj, HtmlPopupOptions opts, HtmlPopup popup)
    {

      // Get data from event obj / IHTMLElements

      var ev = obj.EventObj;
      if (ev.IsNull())
        return;

      bool ctrlPressed = ev.ctrlKey;

      if (popup.IsNull() || opts.IsNull())
        return;

      var srcElement = ev.srcElement;
      if (srcElement.tagName.ToLowerInvariant() != "a")
        return;

      var linkElement = srcElement as IHTMLAnchorElement;
      string url = linkElement?.href;
      if (url.IsNullOrEmpty() || linkElement.IsNull())
        return;

      string innerText = ((IHTMLElement)linkElement).innerText;

      // Create a job to run on the job queue
      Action action = () =>
      {

        var matchedProvider = ProviderMatching.MatchProvidersAgainstMouseoverLink(url, innerText, providers);
        if (matchedProvider.IsNull())
          return;

        // Keep current window open, open new window to the right of the current window
        if (ctrlPressed)
        {
          opts.x += opts.width;
          var doc = popup.GetDocument();
          var wdw = Application.Current.Dispatcher.Invoke(() => doc.parentWindow);

          OpenNewPopupWdw((IHTMLWindow4)wdw, url, innerText, matchedProvider, new RemoteCancellationToken(new CancellationToken()), opts.x, opts.y);
        }

        // Replace the current window with a new window
        else
        {
          var wdw = Application.Current.Dispatcher.Invoke(() => ContentUtils.GetFocusedHtmlWindow());

          action = () =>
          {
            OpenNewPopupWdw((IHTMLWindow4)wdw, url, innerText, matchedProvider, new RemoteCancellationToken(new CancellationToken()), opts.x, opts.y);
          };
        }
      };

      JobQueue.Enqueue(action);

    }

    private void GotoElementButtonClick_OnEvent(object sender, IControlHtmlEventArgs e, int elementId)
    {
      JobQueue.Enqueue(() => GotoElement(elementId));
    }

    private void GotoElement(int elementId)
    {

      if (elementId < 0)
        return;

      if (Svc.SM.Registry.Element[elementId].IsNull())
        return;

      if (!Svc.SM.UI.ElementWdw.GoToElement(elementId))
        LogTo.Warning($"Failed to GoToElement with id {elementId}");

    }

    private void PopupBrowserButtonClick_OnEvent(object sender, IControlHtmlEventArgs e, string query)
    {
      JobQueue.Enqueue(() => OpenUrlInBrowser(query));
    }

    /// <summary>
    /// Open in PopupWindow if available else default browser
    /// </summary>
    /// <param name="query"></param>
    private void OpenUrlInBrowser(string query)
    {
      if (query.IsNullOrEmpty())
        return;

      OpenUrlInUserDefaultBrowser(query);
    }

    private void OpenUrlInUserDefaultBrowser(string query)
    {
      try
      {
        Process.Start(query);
      }
      catch (Exception ex)
      {
        LogTo.Warning($"Exception {ex} thrown while attempting to open {query} in user's default browser");
      }
    }


    [LogToErrorOnException]
    private void ExtractButtonClick_OnEvent(object sender, IControlHtmlEventArgs e, PopupContent content, HtmlPopup popup)
    {

      // Create an action to be passed to the event thread
      // Running on the main UI thread causes deadlock

      Action action = () =>
      {

        try
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

          // Extract the whole popup document
          if (type == ExtractType.Full)
            CreateSMExtract(content.Html, content.References, popup);

          // Extract selected html
          else if (type == ExtractType.Partial)
            CreateSMExtract(selObj.htmlText, content.References, popup);

        }
        catch (UnauthorizedAccessException) { }
        catch (COMException) { }
      };

      JobQueue.Enqueue(action);
    }

    [LogToErrorOnException]
    private void CreateSMExtract(string extract, References refs, HtmlPopup popup)
    {

      try
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
          ElemCreationFlags.None,
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
          string msg = "Failed to CreateSMElement";
          MessageBox.Show(msg);
          LogTo.Error(msg);
        }

      }
      catch (RemotingException) { }

    }

    // Set HasSettings to true, and uncomment this method to add your custom logic for settings
    /// <inheritdoc />
    public override void ShowSettings()
    {
      ConfigurationWindow.ShowAndActivate("MouseoverPopup", HotKeyManager.Instance, Config);
    }


    #endregion

    #region Methods
    #endregion
  }

}
