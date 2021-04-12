using MouseoverPopupInterfaces;
using System;
using static MouseoverPopupInterfaces.PopupContent;

namespace SuperMemoAssistant.Plugins.MouseoverPopup.Models
{

  [Serializable]
  public class ContentProvider
  {
    public string[] urlRegexes { get; set; }
    public IMouseoverContentProvider provider { get; set; }

    public ContentProvider(string[] urlRegexes, IMouseoverContentProvider provider)
    {
      this.urlRegexes = urlRegexes;
      this.provider = provider;
    }
  }
}
