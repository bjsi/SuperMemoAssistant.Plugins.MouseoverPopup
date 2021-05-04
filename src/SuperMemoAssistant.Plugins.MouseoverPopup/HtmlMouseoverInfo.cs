using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public class HtmlMouseoverInfo
  {
    public bool ctrlKey { get; }
    public int x { get; }
    public int y { get; }
    public string url { get; }
    public string innerText { get; }

    public HtmlMouseoverInfo(IHTMLEventObj ev)
    {
      ctrlKey = ev.ctrlKey;
      x = ev.screenX;
      y = ev.screenY;
      var anchor = ev.srcElement as IHTMLAnchorElement;
      url = anchor?.href;
      innerText = ((IHTMLElement)anchor)?.innerText;
    }
  }
}
