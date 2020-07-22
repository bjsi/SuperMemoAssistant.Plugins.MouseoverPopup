using MouseoverPopup.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup.Models
{
  class ContentProviderInfo
  {

    public List<string> urlRegexes { get; set; }
    public IMouseoverContentProvider provider { get; set; }
    public Dictionary<string, string> keywordUrlMap { get; set; }

    // Category Path Regex
    public string[] CategoryPathRegexes { get; set; }

    // Reference Regexes
    public ReferenceRegexes referenceRegexes { get; set; }

    public ContentProviderInfo(List<string> urlRegexes, Dictionary<string, string> keywordUrlMap, ReferenceRegexes regexes, string[] categoryPathRegexes, IMouseoverContentProvider provider)
    {

      this.urlRegexes = urlRegexes;
      this.provider = provider;
      this.keywordUrlMap = keywordUrlMap;
      this.referenceRegexes = regexes;
      this.CategoryPathRegexes = categoryPathRegexes;

    }
  }
}
