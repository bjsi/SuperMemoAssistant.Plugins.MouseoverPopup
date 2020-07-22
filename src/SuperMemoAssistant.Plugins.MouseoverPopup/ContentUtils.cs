using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public static class ContentUtils
  {

    /// <summary>
    /// Get the IHTMLWindow2 object for the currently focused HtmlControl
    /// </summary>
    /// <returns>IHTMLWindow2 object or null</returns>
    public static IHTMLWindow2 GetFocusedHtmlWindow()
    {

      try
      {
        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
        var htmlDoc = htmlCtrl?.GetDocument();
        if (htmlDoc == null)
          return null;

        return htmlDoc.parentWindow;
      }
      catch (UnauthorizedAccessException) { }

      return null;

    }

    /// <summary>
    /// Get the IHTMLDocument2 object representing the focused html control of the current element.
    /// </summary>
    /// <returns>IHTMLDocument2 object or null</returns>
    public static IHTMLDocument2 GetFocusedHtmlDocument()
    {

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
      return htmlCtrl?.GetDocument();

    }

    public static IControlHtml GetFirstHtmlCtrl()
    {

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      return ctrlGroup?.GetFirstHtmlControl()?.AsHtml();

    }

    public static Dictionary<int, IControlHtml> GetHtmlCtrls()
    {

      var ret = new Dictionary<int, IControlHtml>();

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      if (ctrlGroup.IsNull())
        return ret;

      for (int i = 0; i < ctrlGroup.Count; i++)
      {
        var htmlCtrl = ctrlGroup[i].AsHtml();
        if (htmlCtrl.IsNull())
          continue;
        ret.Add(i, htmlCtrl);
      }

      return ret;

    }
  }
}
