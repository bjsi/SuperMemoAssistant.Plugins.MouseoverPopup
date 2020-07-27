using Anotar.Serilog;
using Ganss.Text;
using HtmlAgilityPack;
using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Plugins.MouseoverPopup.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
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

      try
      {

        var htmlCtrls = ContentUtils.GetHtmlCtrls();
        if (htmlCtrls.IsNull() || !htmlCtrls.Any())
          return;

        foreach (KeyValuePair<int, IControlHtml> kvpair in htmlCtrls)
        {

          var htmlCtrl = kvpair.Value;
          var htmlDoc = htmlCtrl?.GetDocument();
          var text = htmlDoc?.body?.innerText
            ?.Replace("\r\n", " ")
            ?.ToLowerInvariant();

          if (text.IsNullOrEmpty() || htmlDoc.IsNull())
            continue;


          // Find matching keywords in the current htmlCtrl
          var matches = keywords
            ?.Search(text)
            ?.Where(x => x.Word.Length > 2);

          // Don't add links to the references section of a component
          int referencesIdx = text.IndexOf("#supermemo reference:");
          if (referencesIdx != -1)
          {

            matches = matches
              ?.Where(x => x.Index < referencesIdx);

          }

          if (matches.IsNull() || !matches.Any())
            continue;

          // Order the keywords by index
          var orderedMatches = matches
            ?.OrderBy(x => x.Index)
            ?.ThenByDescending(x => x.Word.Length);

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

            var replaceDuplicate = selObj.duplicate();

            // mshtml is so buggy with newlines
            // need to check the htmlText if it begins / ends with <BR>

            if (replaceDuplicate.findText(word))
            {

              // skip non-full words
              var isWordDuplicate = replaceDuplicate.duplicate();

              if (isWordDuplicate.moveStart("character", -1) == -1)
              {
                var fstText = isWordDuplicate.text.First();
                var html = isWordDuplicate.htmlText;

                if (!char.IsWhiteSpace(fstText) && !char.IsPunctuation(fstText))
                {

                  bool htmlStartContainsBR = false;
                  var doc = new HtmlDocument();
                  doc.LoadHtml(html);
                  var firstNode = doc.DocumentNode.SelectNodes("//text()").FirstOrDefault();
                  if (!firstNode.IsNull())
                  {
                    var idx = firstNode.InnerStartIndex;
                    if (html.Substring(0, idx + 1).Contains("<BR>"))
                    {
                      htmlStartContainsBR = true;
                    }
                  }

                  if (!htmlStartContainsBR)
                    continue;

                }

              }

              if (isWordDuplicate.moveEnd("character", 2) == 2)
              {
                var lstText = isWordDuplicate.text[isWordDuplicate.text.Length - 2];
                var html = isWordDuplicate.htmlText;

                if (!char.IsWhiteSpace(lstText) && !char.IsPunctuation(lstText))
                {

                  bool htmlEndContainsBR = false;
                  var doc = new HtmlDocument();
                  doc.LoadHtml(html);
                  var lastNode = doc.DocumentNode.SelectNodes("//text()").LastOrDefault();
                  if (!lastNode.IsNull())
                  {
                    var idx = lastNode.InnerStartIndex + lastNode.InnerLength;
                    if (html.Substring(idx).Contains("<BR>"))
                    {
                      htmlEndContainsBR = true;
                    }
                  }

                  if (!htmlEndContainsBR)
                    continue;
                }
              }

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
              selObj.setEndPoint("StartToEnd", replaceDuplicate);
              replaceDuplicate.pasteHTML($"<a href='{href}'>{replaceDuplicate.text}<a>");
            }
          }
        }
      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }
      catch (Exception ex)
      {
        LogTo.Error($"Exception {ex} while executing ScanAndAddLinks");
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
