using Anotar.Serilog;
using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public static class ContentUtils
  {

    /// <summary>
    /// Get the IHTMLWindow2 object for the currently focused HtmlControl
    /// </summary>
    /// <returns>IHTMLWindow2 object or null</returns>
    public static IHTMLWindow4 GetFocusedHtmlWindow()
    {
      try
      {
        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
        var htmlDoc = htmlCtrl?.GetDocument();
        return htmlDoc?.parentWindow as IHTMLWindow4;
      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

      return null;
    }

    /// <summary>
    /// Get the IHTMLDocument2 object representing the focused html control of the current element.
    /// </summary>
    /// <returns>IHTMLDocument2 object or null</returns>
    public static IHTMLDocument2 GetFocusedHtmlDocument()
    {

      try
      {

        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
        return htmlCtrl?.GetDocument();

      }
      catch (UnauthorizedAccessException) { }

      return null;
    }

    public static IControlHtml GetFirstHtmlCtrl()
    {

      try
      {

        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        return ctrlGroup?.GetFirstHtmlControl()?.AsHtml();

      }
      catch (UnauthorizedAccessException) { }

      return null;


    }

    public static List<IControlHtml> GetHtmlCtrls()
    {

      try
      {

        var ret = new List<IControlHtml>();

        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        if (ctrlGroup.IsNull())
          return ret;

        for (int i = 0; i < ctrlGroup.Count; i++)
        {
          var htmlCtrl = ctrlGroup[i].AsHtml();
          if (htmlCtrl.IsNull())
            continue;
          ret.Add(htmlCtrl);
        }

        return ret;

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

      return null;


    }
  }
}
