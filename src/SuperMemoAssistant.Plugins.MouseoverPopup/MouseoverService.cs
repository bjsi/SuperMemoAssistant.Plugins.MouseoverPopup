﻿using MouseoverPopup.Interop;
using PluginManager.Interop.Sys;
using SuperMemoAssistant.Interop.Plugins;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public class MouseoverService : PerpetualMarshalByRefObject, IMouseoverSvc
  {
    public bool RegisterProvider(string name, Func<string, bool> predicate, IContentProvider provider)
    {
      return Svc<MouseoverPopupPlugin>.Plugin.RegisterProvider(name, predicate, provider);
    }
  }
}
