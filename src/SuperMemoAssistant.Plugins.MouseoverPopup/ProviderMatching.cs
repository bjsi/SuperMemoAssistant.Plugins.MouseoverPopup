using MouseoverPopup.Interop;
using SuperMemoAssistant.Interop.SuperMemo.Elements.Models;
using SuperMemoAssistant.Interop.SuperMemo.Elements.Types;
using SuperMemoAssistant.Plugins.MouseoverPopup.Models;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{

  /// <summary>
  /// Methods for matching content providers against the current element.
  /// </summary>
  public static class ProviderMatching
  {

    /// <summary>
    /// Finds the providers that can supply content for the current element.
    /// </summary>
    /// <param name="providers"></param>
    /// <returns></returns>
    public static Dictionary<string, ContentProvider> MatchProvidersAgainstCurrentElement(Dictionary<string, ContentProvider> providers)
    {

      var ret = new Dictionary<string, ContentProvider>();

      if (providers.IsNull() || !providers.Any())
        return ret;

      foreach (var provider in providers)
      {

        if (provider.Value.keywordScanningOptions.IsNull())
          continue;

        if (MatchProviderAgainstCurrentElement(provider.Value))
          ret.Add(provider.Key, provider.Value);

      }

      return ret;

    }

    /// <summary>
    /// Returns true if the category path or the current references match
    /// </summary>
    /// <param name="provider"></param>
    /// <returns></returns>
    private static bool MatchProviderAgainstCurrentElement(ContentProvider provider)
    {

      var element = Svc.SM.UI.ElementWdw.CurrentElement;
      if (element.IsNull() || provider.IsNull() || provider.keywordScanningOptions.IsNull())
        return false;

      var referenceRegexes = provider.keywordScanningOptions.referenceRegexes;
      var categoryPathRegexes = provider.keywordScanningOptions.categoryPathRegexes;


      return MatchAgainstCategoryPath(element, categoryPathRegexes) || MatchAgainstCurrentReferences(referenceRegexes);

    }

    private static bool MatchAgainstCurrentReferences(ReferenceRegexes refRegexes)
    {

      var htmlCtrl = ContentUtils.GetFirstHtmlCtrl();
      string text = htmlCtrl?.Text;
      if (text.IsNullOrEmpty())
        return false;

      var refs = ReferenceParser.GetReferences(htmlCtrl?.Text);
      if (refs.IsNull())
        return false;

      else if (MatchAgainstRegexes(refs.Source, refRegexes.SourceRegexes))
        return true;

      else if (MatchAgainstRegexes(refs.Link, refRegexes.LinkRegexes))
        return true;

      if (MatchAgainstRegexes(refs.Title, refRegexes.TitleRegexes))
        return true;

      else if (MatchAgainstRegexes(refs.Author, refRegexes.AuthorRegexes))
        return true;

      return false;

    }

    /// <summary>
    /// Return true if any regex in the array matches.
    /// </summary>
    /// <param name="input"></param>
    /// <param name="regexes"></param>
    /// <returns></returns>
    private static bool MatchAgainstRegexes(string input, string[] regexes)
    {

      if (input.IsNullOrEmpty())
        return false;

      if (regexes.IsNull() || !regexes.Any())
        return false;

      if (regexes.Any(r => new Regex(r).Match(input).Success))
        return true;

      return false;

    }

    /// <summary>
    /// Returns true if any of the regexes in the array match the category path
    /// </summary>
    /// <param name="element"></param>
    /// <param name="regexes"></param>
    /// <returns></returns>
    private static bool MatchAgainstCategoryPath(IElement element, string[] regexes)
    {

      if (element.IsNull() || regexes.IsNull() || !regexes.Any())
        return false;

      var cur = element.Parent;
      while (!cur.IsNull())
      {
        if (cur.Type == ElementType.ConceptGroup)
        {

          // TODO: Check that this works
          var concept = Svc.SM.Registry.Concept[cur.Id];
          string name = concept?.Name;

          if (!name.IsNullOrEmpty() && regexes.Any(x => new Regex(x).Match(name).Success))
            return true;

        }
        cur = cur.Parent;
      }

      return false;

    }

    /// <summary>
    /// Find all providers against the url and the inner text of the anchor element.
    /// </summary>
    /// <param name="url"></param>
    /// <param name="text"></param>
    /// <param name="potentialProviders"></param>
    /// <returns></returns>
    public static Dictionary<string, ContentProvider> MatchProvidersAgainstMouseoverLink(string url, string text, Dictionary<string, ContentProvider> potentialProviders)
    {

      if (url.IsNullOrEmpty() || text.IsNullOrEmpty() || potentialProviders.IsNull() || !potentialProviders.Any())
        return null;

      var ret = new Dictionary<string, ContentProvider>();

      foreach (var provider in potentialProviders)
      {

        var regexes = provider.Value.urlRegexes;

        if (regexes.Any(r => new Regex(r).Match(url).Success)
         || provider.Value.keywordScanningOptions.keywordMap.Keys.Any(x => x == text)) // TODO: Should be lower?
        {
          ret.Add(provider.Key, provider.Value);
        }

      }

      return ret;

    }

  }
}
