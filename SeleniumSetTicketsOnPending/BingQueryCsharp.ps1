﻿<#
Endpoints
https://api.cognitive.microsoft.com/bing/v7.0/suggestions

https://api.cognitive.microsoft.com/bing/v7.0/entities

https://api.cognitive.microsoft.com/bing/v7.0/images

https://api.cognitive.microsoft.com/bing/v7.0/localbusinesses

https://api.cognitive.microsoft.com/bing/v7.0/news

https://api.cognitive.microsoft.com/bing/v7.0/spellcheck

https://api.cognitive.microsoft.com/bing/v7.0/urlpreview

https://api.cognitive.microsoft.com/bing/v7.0/videos

https://api.cognitive.microsoft.com/bing/v7.0/images/visualsearch

https://api.cognitive.microsoft.com/bing/v7.0

Key 1: 27aa28ad1786420b87baee4a57ddb03c

Key 2: 73fdc0c495c340deaa3cadecad7b7d76
#>
<#
Console.OutputEncoding = System.Text.Encoding.UTF8;
$BingQuery = "Trump"
$MyBingApiKey = '73fdc0c495c340deaa3cadecad7b7d76'
$WebSearch = Invoke-RestMethod -Uri `
"https://api.cognitive.microsoft.com/bing/v5.0/search?q=$BingQuery&count=3&mkt=en-us" `
-Headers @{ "Ocp-Apim-Subscription-Key" = $MyBingApiKey } 
#>

$code = @"
using System;
using System.Text;
using System.Net;
using System.IO;
using System.Collections.Generic;

namespace BingSearchApisQuickstart
{

    class Program
    {
        // **********************************************
        // *** Update or verify the following values. ***
        // **********************************************

        // Replace the accessKey string value with your valid access key.
        const string accessKey = "73fdc0c495c340deaa3cadecad7b7d76";

        // Verify the endpoint URI.  At this writing, only one endpoint is used for Bing
        // search APIs.  In the future, regional endpoints may be available.  If you
        // encounter unexpected authorization errors, double-check this value against
        // the endpoint for your Bing Web search instance in your Azure dashboard.
        const string uriBase = "https://api.cognitive.microsoft.com/bing/v7.0/search";

        const string searchTerm = "Microsoft Cognitive Services";

        // Used to return search results including relevant headers
        struct SearchResult
        {
            public String jsonResult;
            public Dictionary<String, String> relevantHeaders;
        }

        static void Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            if (accessKey.Length == 32)
            {
                Console.WriteLine("Searching the Web for: " + searchTerm);

                SearchResult result = BingWebSearch(searchTerm);

                Console.WriteLine("\nRelevant HTTP Headers:\n");
                foreach (var header in result.relevantHeaders)
                    Console.WriteLine(header.Key + ": " + header.Value);

                Console.WriteLine("\nJSON Response:\n");
                Console.WriteLine(JsonPrettyPrint(result.jsonResult));
            }
            else
            {
                Console.WriteLine("Invalid Bing Search API subscription key!");
                Console.WriteLine("Please paste yours into the source code.");
            }

            Console.Write("\nPress Enter to exit ");
            Console.ReadLine();
        }

        /// <summary>
        /// Performs a Bing Web search and return the results as a SearchResult.
        /// </summary>
        static SearchResult BingWebSearch(string searchQuery)
        {
            // Construct the URI of the search request
            var uriQuery = uriBase + "?q=" + Uri.EscapeDataString(searchQuery);

            // Perform the Web request and get the response
            WebRequest request = HttpWebRequest.Create(uriQuery);
            request.Headers["Ocp-Apim-Subscription-Key"] = accessKey;
            HttpWebResponse response = (HttpWebResponse)request.GetResponseAsync().Result;
            string json = new StreamReader(response.GetResponseStream()).ReadToEnd();

            // Create result object for return
            var searchResult = new SearchResult()
            {
                jsonResult = json,
                relevantHeaders = new Dictionary<String, String>()
            };

            // Extract Bing HTTP headers
            foreach (String header in response.Headers)
            {
                if (header.StartsWith("BingAPIs-") || header.StartsWith("X-MSEdge-"))
                    searchResult.relevantHeaders[header] = response.Headers[header];
            }

            return searchResult;
        }

        /// <summary>
        /// Formats the given JSON string by adding line breaks and indents.
        /// </summary>
        /// <param name="json">The raw JSON string to format.</param>
        /// <returns>The formatted JSON string.</returns>
        static string JsonPrettyPrint(string json)
        {
            if (string.IsNullOrEmpty(json))
                return string.Empty;

            json = json.Replace(Environment.NewLine, "").Replace("\t", "");

            StringBuilder sb = new StringBuilder();
            bool quote = false;
            bool ignore = false;
            char last = ' ';
            int offset = 0;
            int indentLength = 2;

            foreach (char ch in json)
            {
                switch (ch)
                {
                    case '"':
                        if (!ignore) quote = !quote;
                        break;
                    case '\\':
                        if (quote && last != '\\') ignore = true;
                        break;
                }

                if (quote)
                {
                    sb.Append(ch);
                    if (last == '\\' && ignore) ignore = false;
                }
                else
                {
                    switch (ch)
                    {
                        case '{':
                        case '[':
                            sb.Append(ch);
                            sb.Append(Environment.NewLine);
                            sb.Append(new string(' ', ++offset * indentLength));
                            break;
                        case '}':
                        case ']':
                            sb.Append(Environment.NewLine);
                            sb.Append(new string(' ', --offset * indentLength));
                            sb.Append(ch);
                            break;
                        case ',':
                            sb.Append(ch);
                            sb.Append(Environment.NewLine);
                            sb.Append(new string(' ', offset * indentLength));
                            break;
                        case ':':
                            sb.Append(ch);
                            sb.Append(' ');
                            break;
                        default:
                            if (quote || ch != ' ') sb.Append(ch);
                            break;
                    }
                }
                last = ch;
            }

            return sb.ToString().Trim();
        }
    }
}
"@

if (-not ([System.Management.Automation.PSTypeName]'MyNameSpace.DotnetPiper').Type)
{
 Add-Type -TypeDefinition $code -Language CSharp;
}<[MyNameSpace.DotnetPiper]::new().DotnetPiperInstanceMethod()[MyNameSpace.DotnetPiperExtension]::DotnetPiperExtentionMethod($this.DotnetPiper ,"q") 
#>

[BingSearchApisQuickstart]::[Program].Invoke(2222)










 <# 
class Program  
{  
    static void Main(string[] args)  
    {  
        DotnetPiper dp = new DotnetPiper();  
        string nameStr = dp.DotnetPiperExtentionMethod("Sachin Kalia");  
        Console.WriteLine(nameStr);  
        Console.ReadLine();  
  
                }  
       }  #>