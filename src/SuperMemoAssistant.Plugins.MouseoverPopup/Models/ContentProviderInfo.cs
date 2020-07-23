using MouseoverPopup.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup.Models
{

  [Serializable]
  public class ContentProvider
  {

    public string[] urlRegexes { get; set; }
    public IMouseoverContentProvider provider { get; set; }
    public KeywordScanningOptions keywordScanningOptions { get; set; }

    public ContentProvider(string[] urlRegexes, IMouseoverContentProvider provider)
    {

      this.urlRegexes = urlRegexes;
      this.provider = provider;

    }

    public ContentProvider(string[] urlRegexes, KeywordScanningOptions keywordScanningOptions, IMouseoverContentProvider provider)
    {

      this.urlRegexes = urlRegexes;
      this.keywordScanningOptions = keywordScanningOptions;
      this.provider = provider;

    }

  }
}
