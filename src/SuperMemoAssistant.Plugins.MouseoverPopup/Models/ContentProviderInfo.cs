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
    public Func<string, bool> predicate { get; set; }
    public IContentProvider provider { get; set; }
    public ContentProviderInfo(Func<string, bool> predicate, IContentProvider provider)
    {
      this.predicate = predicate;
      this.provider = provider;
    }
  }
}
