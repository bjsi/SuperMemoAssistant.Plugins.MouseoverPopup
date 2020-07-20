using MouseoverPopup.Interop;
using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{

  public class HtmlPopup
  {

    private IHTMLPopup _popup { get; set; }

    public event EventHandler<HtmlPopupEventArgs> OnShow;

    public HtmlPopup(IHTMLWindow4 wdw)
    {
      _popup = wdw?.createPopup() as IHTMLPopup;
    }

    public void Show(int screenX, int screenY, int w, int h)
    {

      if (_popup.IsNull())
        return;

      _popup.Show(screenX, screenY, w, h, null);
      OnShow?.Invoke(this, new HtmlPopupEventArgs(screenX, screenY, w, h));

    }

    public bool IsOpen()
    {
      return _popup.IsNull()
        ? false
        : _popup.isOpen;
    }

    public void Hide()
    {
      _popup.Hide();
    }

    public IHTMLDocument2 GetDocument()
    {
      return _popup?.document as IHTMLDocument2;
    }

  }

  public static class PopupEx
  {
    public static HtmlPopup CreatePopup(this IHTMLWindow4 wdw)
    {
      return new HtmlPopup(wdw);
    }
  }
}
