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
  using System.Collections.Generic;
  using System.Diagnostics.CodeAnalysis;
  using System.Linq;
  using global::MouseoverPopup.Interop;
  using SuperMemoAssistant.Interop.Plugins;
  using SuperMemoAssistant.Plugins.MouseoverPopup.Models;
  using SuperMemoAssistant.Services.Sentry;

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

    #endregion


    #region Methods Impl

    /// <inheritdoc />
    protected override void PluginInit()
    {

      _mouseoverSvc = new MouseoverService();
      PublishService<IMouseoverSvc, MouseoverService>(_mouseoverSvc);

    }

    public bool RegisterProvider(string name, Func<string, bool> predicate, IContentProvider provider)
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

    // Uncomment to register an event handler for element changed events
    // [LogToErrorOnException]
    // public void OnElementChanged(SMDisplayedElementChangedEventArgs e)
    // {
    //   try
    //   {
    //     Insert your logic here
    //   }
    //   catch (RemotingException) { }
    // }

    #endregion
  }
}
