using Forge.Forms.Annotations;
using Newtonsoft.Json;
using SuperMemoAssistant.Services.UI.Configuration;
using SuperMemoAssistant.Sys.ComponentModel;
using System.ComponentModel;

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
  public class MouseoverPopupCfg : CfgBase<MouseoverPopupCfg>, INotifyPropertyChangedEx
  {

    [Title("Mouseover Popup Plugin")]
    [Heading("By Jamesb | Experimental Learning")]

    [Heading("Features")]
    [Text(@"- Quickly preview the content of urls for different sites by hovering over links in SuperMemo.
- Publishes itself as a SuperMemoAssistant Service to integrate with content provider plugins such as Mouseover Wiki
- Easily extend Mouseover Popup's capabilities by creating a new content provider.")]

    [Heading("Support")]
    [Text("If you would like to support my projects, check out my Patreon or buy me a coffee.")]

    [Action("patreon", "Patreon", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("coffee", "Coffee", Placement = Placement.Before, LinePosition = Position.Left)]

    [Heading("Links")]
    [Action("github", "GitHub", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("feedback", "Feedback Site", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("blog", "Blog", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("youtube", "YouTube", Placement = Placement.Before, LinePosition = Position.Left)]
    [Action("twitter", "Twitter", Placement = Placement.Before, LinePosition = Position.Left)]

    [Heading("Settings")]

    [Field(Name = "Default Popup Extract Priority (%)")]
    [Value(Must.BeGreaterThanOrEqualTo,
           0,
           StrictValidation = true)]
    [Value(Must.BeLessThanOrEqualTo,
           100,
           StrictValidation = true)]
    public double DefaultPriority { get; set; } = 30;

    [Field(Name = "Require ctrl key to be pressed to open links?")]
    public bool RequireCtrlKey { get; set; } = false;


    [Field]

    [JsonIgnore]
    public bool IsChanged { get; set; }

    public override string ToString()
    {
      return "Mouseover Popup";
    }

    public event PropertyChangedEventHandler PropertyChanged;
  }
}
