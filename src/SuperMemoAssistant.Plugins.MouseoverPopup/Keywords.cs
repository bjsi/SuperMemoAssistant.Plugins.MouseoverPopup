using Ganss.Text;
using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Plugins.MouseoverPopup.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.MouseoverPopup
{
  public static class Keywords
  {

    private const int MAX_TEXT_LENGTH = 2000000000;

    /// <summary>
    /// TODO: Find a way to skip element references
    /// Scan the current element for keywords.
    /// </summary>
    /// <param name="providers"></param>
    public static void ScanAndAddLinks(Dictionary<string, ContentProvider> providers, AhoCorasick keywords)
    {

      var htmlCtrls = ContentUtils.GetHtmlCtrls();
      if (htmlCtrls.IsNull() || !htmlCtrls.Any())
        return;

      foreach (KeyValuePair<int, IControlHtml> kvpair in htmlCtrls)
      {

        var htmlCtrl = kvpair.Value;
        var htmlDoc = htmlCtrl?.GetDocument();
        var text = htmlDoc?.body?.innerText;
        if (text.IsNullOrEmpty() || htmlDoc.IsNull())
          continue;

        // Find matching keywords in the current htmlCtrl
        var matches = keywords.Search(text);
        if (matches.IsNull() || !matches.Any())
          continue;

        // TODO: Sort by longest matches, then index or vice versa
        // Order the keywords by index
        var orderedMatches = matches.OrderBy(x => x.Index);
        var selObj = htmlDoc.selection?.createRange() as IHTMLTxtRange;
        if (orderedMatches.IsNull() || !orderedMatches.Any() || selObj.IsNull())
          continue;

        // If the matched keyword is not already a link,
        // wrap it in an anchor element with the first matching url

        foreach (var match in orderedMatches)
        {

          string word = match.Word;
          if (word.IsNullOrEmpty())
            continue;

          if (selObj.findText(word, Flags: 2)) // Match whole words only
          {

            // Skip keywords that are already links

            var parentEl = selObj.parentElement();
            if (!parentEl.IsNull())
            {
              if (parentEl.tagName.ToLowerInvariant() == "a")
                continue;
            }

            // Add the first url that matches from the providers

            string href = null;
            foreach (var provider in providers)
            {
              if (provider.Value.keywordScanningOptions.urlKeywordMap.TryGetValue(word, out href))
                break;
            }

            if (href.IsNullOrEmpty())
              continue;

            // Wrap in a link
            selObj.pasteHTML($"<a href='{href}'>{selObj.text}<a>");

          }

          // Since the keywords are index ordered, can collapse to the end and continue searching

          selObj.collapse(false);
          selObj.moveEnd("character", MAX_TEXT_LENGTH);

        }
      }
    }

    /// <summary>
    /// Create the keyword search datastructure.
    /// </summary>
    /// <param name="providers"></param>
    public static AhoCorasick CreateKeywords(Dictionary<string, ContentProvider> providers)
    {

      var ret = new AhoCorasick();

      if (providers.IsNull() || !providers.Any())
        return ret;

      foreach (var provider in providers)
      {
        
        var words = provider.Value.keywordScanningOptions.urlKeywordMap?.Keys;
        if (words.IsNull() || !words.Any())
          continue;

        ret.Add(words);

      }

      return ret;

    }
  }
}
