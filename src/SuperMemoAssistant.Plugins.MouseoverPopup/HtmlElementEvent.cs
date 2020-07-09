using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  public class HtmlElementEvent
  {

    private Action action { get; set; }
    
    public HtmlElementEvent(Action action)
    {
      this.action = action;
    }

    [DispId(0)]
    public void handler(IHTMLEventObj e)
    {
      action();
    }
  }
}
