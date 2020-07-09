using Forge.Forms.Annotations;
using Newtonsoft.Json;
using SuperMemoAssistant.Services.UI.Configuration;
using SuperMemoAssistant.Sys.ComponentModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  [Form(Mode = DefaultFields.None)]
  [Title("Dictionary Settings",
         IsVisible = "{Env DialogHostContext}")]
  [DialogAction("cancel",
        "Cancel",
        IsCancel = true)]
  [DialogAction("save",
        "Save",
        IsDefault = true,
        Validates = true)]
  class MouseoverPopupCfg : CfgBase<MouseoverPopupCfg>, INotifyPropertyChangedEx
  {

    [Field(Name = "Block SM default url click behaviour?")]
    public bool BlockUrlMouseClick { get; set; } = true;

    [Field(Name = "Highlight targetted urls")]
    public bool HighlightUrls { get; set; } = true;

    [JsonIgnore]
    public bool IsChanged { get; set; }

    public override string ToString()
    {
      return "Mouseover Popup";
    }

    public event PropertyChangedEventHandler PropertyChanged;
  }
}
