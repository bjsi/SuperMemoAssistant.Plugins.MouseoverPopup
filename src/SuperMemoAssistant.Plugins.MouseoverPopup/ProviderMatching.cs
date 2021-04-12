using SuperMemoAssistant.Plugins.MouseoverPopup.Models;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{

  /// <summary>
  /// Methods for matching content providers against the current element.
  /// </summary>
  public static class ProviderMatching
  {

    /// <summary>
    /// Find the first provider that matches against the url
    /// </summary>
    /// <param name="url"></param>
    /// <param name="text"></param>
    /// <param name="potentialProviders"></param>
    /// <returns></returns>
    public static ContentProvider MatchProvidersAgainstMouseoverLink(string url, string text, Dictionary<string, ContentProvider> potentialProviders)
    {

      if (url.IsNullOrEmpty() || text.IsNullOrEmpty() || potentialProviders.IsNull() || !potentialProviders.Any())
        return null;

      foreach (var provider in potentialProviders)
      {

        var regexes = provider.Value.urlRegexes;

        if (regexes.Any(r => new Regex(r).Match(url).Success))
        {
          return provider.Value;
        }

      }

      return null;

    }

  }
}
