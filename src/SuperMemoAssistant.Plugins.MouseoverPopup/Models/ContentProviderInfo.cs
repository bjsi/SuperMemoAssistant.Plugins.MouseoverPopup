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
    public IContentProvider provider { get; set; }
    public ContentProviderInfo(List<string> urlRegexes, IContentProvider provider)
    {
      this.urlRegexes = urlRegexes;
      this.provider = provider;
    }
  }
}
