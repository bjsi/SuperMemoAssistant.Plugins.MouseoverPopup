using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public static class PopupEx
  {
    public static IHTMLPopup CreatePopup()
    {
      try
      {

        var wdw = ContentUtils.GetFocusedHtmlWindow() as IHTMLWindow4;
        var popup = wdw?.createPopup() as IHTMLPopup;

        // Styling
        var doc = popup?.document as IHTMLDocument2;
        if (!doc.IsNull())
          doc.body.style.border = "solid black 1px";

        return popup;

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return null;
    }


  }
}
