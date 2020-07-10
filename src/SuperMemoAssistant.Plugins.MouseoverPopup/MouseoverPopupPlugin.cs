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
  using System.Diagnostics.CodeAnalysis;
  using System.Linq;
  using System.Text.RegularExpressions;
  using System.Threading;
  using System.Windows;
  using Anotar.Serilog;
  using global::MouseoverPopup.Interop;
  using mshtml;
  using SuperMemoAssistant.Extensions;
  using SuperMemoAssistant.Interop.Plugins;
  using SuperMemoAssistant.Plugins.MouseoverPopup.Models;
  using SuperMemoAssistant.Plugins.MouseoverPopup.UI;
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
    public override bool HasSettings => false;
    private Dictionary<string, ContentProviderInfo> providers { get; set; } = new Dictionary<string, ContentProviderInfo>();
    private MouseoverService _mouseoverSvc { get; set; }
    private HTMLControlEvents htmlDocEvents { get; set; }
    private PopupWdw CurrentWdw { get; set; }

    #endregion

    #region Methods Impl

    /// <inheritdoc />
    protected override void PluginInit()
    {
      _mouseoverSvc = new MouseoverService();
      PublishService<IMouseoverSvc, MouseoverService>(_mouseoverSvc);
      SubscribeToMouseoverEvents();
    }

    private void SubscribeToMouseoverEvents()
    {
      var events = new List<EventInitOptions>
      {
        new EventInitOptions(EventType.onmouseenter, _ => true, x => ((IHTMLElement)x).tagName.ToLower() == "a")
      };

      htmlDocEvents = new HTMLControlEvents(events);
      htmlDocEvents.OnMouseEnterEvent += HtmlDocEvents_OnMouseEnterEvent;
    }

    private void HtmlDocEvents_OnMouseEnterEvent(object sender, IHTMLControlEventArgs e)
    {

      if (CurrentWdw != null || !CurrentWdw.IsClosed)
        return;

      var link = e.EventObj.srcElement as IHTMLAnchorElement;
      string url = link?.href;
      
      // TODO: Sub to mouseleave immediately??
      // Cancel if window already open
      
      foreach (var keyValuePair in providers)
      {
        var regexes = keyValuePair.Value.urlRegexes;
        var provider = keyValuePair.Value.provider;

        if (regexes.Any(r => new Regex(r).Match(url).Success))
        {
          RemoteCancellationTokenSource ct = new RemoteCancellationTokenSource();
          SubscribeToMouseLeaveEvent(link, ct);
          OpenNewPopupWdw(url, provider, ct.Token);
        }
      }
    }

    private void OpenNewPopupWdw(string url, IContentProvider provider, RemoteCancellationToken ct)
    {
      Application.Current.Dispatcher.Invoke(() => 
      {
        var wdw = new PopupWdw(url, provider, ct);
        wdw.ShowAndActivate();
        CurrentWdw = wdw;
      });
    }

    private void SubscribeToMouseLeaveEvent(IHTMLAnchorElement link, RemoteCancellationTokenSource ct)
    {
      if (link == null)
      {
        LogTo.Warning("Failed to subscribe to MouseLeaveEvent");
        return;
      }

      ((IHTMLElement2)link).attachEvent("onmouseleave", new HtmlElementEvent(ct.Cancel));

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
