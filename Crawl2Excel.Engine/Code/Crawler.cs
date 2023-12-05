using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Net;
using System.Security;
using System.Text;
using System.Threading;
using Abot2.Crawler;
using Abot2.Poco;
using AngleSharp.Io;
using Crawl2Excel.Engine.Models;
using CSSParser;
using CSSParser.ContentProcessors;
using Microsoft.Extensions.Logging;

namespace Crawl2Excel.Engine.Code
{
	public class Crawler
	{
		private readonly Uri url;
		private readonly FileInfo excelResult;
		private PoliteWebCrawler? abot;
		private BlockingCollection<CrawledPageResult> pages = new BlockingCollection<CrawledPageResult>();
		private ConcurrentDictionary<int, string> crawledPages = new ConcurrentDictionary<int, string>();
		private List<ParsedLink> innerLinks = new List<ParsedLink>();
		private int crawledPagesCount = 0;
		
		public CrawlerOptions Options { get; private set; } = new CrawlerOptions();
		public CrawlResult Result { get; private set; } = new CrawlResult();

		public Crawler(string startUrl, string resultExcelFile)
		{
			url = new Uri(startUrl);
			if (!string.IsNullOrEmpty(resultExcelFile)) 
			{
				excelResult = new FileInfo(resultExcelFile);
			}
			else
			{
				excelResult = new FileInfo(Path.Combine(Environment.CurrentDirectory, $"{url.Host}.xlsx"));
			}
		}

		public async Task Start()
		{
			var textContentTypes = "text/html,text/plain,text/csv,text/xml";
			var appContentTypes = "application/pdf,application/msword,application/vnd.ms-excel,application/vnd.ms-powerpoint,application/octet-stream,application/json,application/rtf,application/xhtml+xml,application/xml,application/zip";
			var config = new CrawlConfiguration
			{
				MaxPagesToCrawl = 10000,
				IsExternalPageCrawlingEnabled = false,
				IsExternalPageLinksCrawlingEnabled = false,
				HttpRequestTimeoutInSeconds = 600,
				DownloadableContentTypes = $"{textContentTypes},{appContentTypes}",			
			};

			if (excelResult.Exists)
			{
				if (Options.OverwriteResultExcelIfExists)
				{
					excelResult.Delete();
				}
				else
				{
					throw new Exception($"Crawl2Excel error: result file {excelResult.FullName} allready exists.");
				}
			}

			var crawler = new PoliteWebCrawler(config);
			crawler.PageCrawlStarting += Crawler_PageCrawlStarting;
			crawler.PageCrawlCompleted += Crawler_PageCrawlCompleted;
			crawler.PageCrawlDisallowed += Crawler_PageCrawlDisallowed;
			crawler.ShouldCrawlPageDecisionMaker = (page, ctx) =>
			{
				return DecideToCrawlUrl(page?.Uri?.LocalPath);
			};

			Console.Clear();
			Console.WriteLine("Crawling: " + url.ToString());
			Result = await crawler.CrawlAsync(url);

			await FetchAndParseInnerLinks(url);

			if (pages.Any())
			{
				WriteResultsToExcel(pages.ToList());
			}
		}

		private readonly object displayLock = new object();
		private void DisplayCurrentCrawledPagesCount()
		{
			lock (displayLock)
			{
				Console.SetCursorPosition(0, 2);
				Console.WriteLine($"Pages Crawled: {crawledPagesCount}   ");
			}
		}

		private CrawlDecision DecideToCrawlUrl(string url)
		{
			// don't wont data:application URIs
			if ((url ?? "").ToLower().Contains($"data:application/"))
			{
				return new CrawlDecision { Allow = false, Reason = "data:application" };
			}

			// don't wont data:image URIs
			if ((url ?? "").ToLower().Contains($"data:image/"))
			{
				return new CrawlDecision { Allow = false, Reason = "data:image" };
			}

			return new CrawlDecision { Allow = true };
		}

		private void Crawler_PageCrawlDisallowed(object sender, PageCrawlDisallowedArgs e)
		{
			var ignoreReasons = new List<string> {
				"data:application",
				"data:image"
			};

			if (ignoreReasons.Contains(e.DisallowedReason))
			{
				return;
			}
			var result = new CrawledPageResult();
			result.Url = e.PageToCrawl.Uri.ToString();
			result.Referer = e.PageToCrawl.ParentUri.ToString();
			result.Status = 500;
			result.TimeMiliseconds = 0;
			result.Size = 0;
			result.Error = e.DisallowedReason;
			AddCrawlResult(result);
		}

		private void Crawler_PageCrawlCompleted(object sender, PageCrawlCompletedArgs e)
		{
			if (e.CrawledPage != null)
			{
				StoreCrawlData(e.CrawledPage);
			}
		}

		private void Crawler_PageCrawlStarting(object sender, PageCrawlStartingArgs e)
		{
		}

		private void StoreCrawlData(CrawledPage crawledPage, string error = "")
		{
			string url = crawledPage.Uri?.ToString()?.ToLower();
			if (string.IsNullOrEmpty(url))
			{
				return;
			}
			int hash = url.GetHashCode();
			if (crawledPages.ContainsKey(hash))
			{
				return;
			}
			crawledPages.TryAdd(hash, "");

			var result = new CrawledPageResult();
			result.Url = url;
			result.Referer = crawledPage.ParentUri.ToString();
			if (crawledPage.HttpResponseMessage != null)
			{
				result.Status = (int)crawledPage.HttpResponseMessage.StatusCode;
				result.TimeMiliseconds = (long)crawledPage.Elapsed;
				result.Size = 0;
				if (crawledPage.Content.Bytes != null)
				{
					result.Size = crawledPage.Content.Bytes.Length;
				}
				else if (crawledPage.HttpResponseMessage.Content.Headers.ContentLength.HasValue)
				{
					result.Size = crawledPage.HttpResponseMessage.Content.Headers.ContentLength.Value;
				}
				result.Error = error;

				var infoWorker = new PageInfoWorker(crawledPage);
				result.PageInfo = infoWorker.GetInfo();
				if (result.PageInfo.IsHtml)
				{
					result.Seo = infoWorker.GetSeoData();
					result.OpenGraph = infoWorker.GetOpenGraphData();
				}
			}
			else
			{
				result.Status = (int)HttpStatusCode.ServiceUnavailable;
				result.Error = "No connection!";
			}

			if (crawledPage.HttpRequestException != null)
			{
				if (!string.IsNullOrEmpty(error))
				{
					result.Error += Environment.NewLine;
				}
				result.Error += crawledPage.HttpRequestException.Message;
			}
			
			AddCrawlResult(result);

			var contentParser = new HtmlContentParser(crawledPage);
			contentParser.ParseContent();
			innerLinks.AddRange(contentParser.ParsedLinks);
		}

		private async Task FetchAndParseInnerLinks(Uri rootUri)
		{
			var allreadyParsed = new HashSet<string>();

			foreach (var link in innerLinks)
			{
				await ProcessInnerLink(rootUri, allreadyParsed, link);
			}
		}

		private async Task ProcessInnerLink(Uri rootUri, HashSet<string> allreadyParsed, ParsedLink link)
		{
			var linkUri = new Uri(rootUri.Scheme + "://" + link.Url.Replace("http://", "").Replace("https://", ""));
			string sameSchemeUrl = linkUri.GetLeftPart(UriPartial.Path);

			if (!allreadyParsed.Contains(sameSchemeUrl) && !string.IsNullOrEmpty(linkUri.LocalPath) && linkUri.LocalPath != "/")
			{
				allreadyParsed.Add(sameSchemeUrl);

				var result = new CrawledPageResult();
				result.Url = sameSchemeUrl;
				result.Referer = link.Referer;
				result.PageInfo.ContentType = link.ContentType;

				byte[] content = new byte[0];
				for (int i = 0; i < 3; i++)
				{
					content = await DownloadFileAndFilResult(link.Url, result);
					if (content.Length > 0)
					{
						break;
					}

					Crawl2ExcelLogger.Logger.LogError($"Error downloading file: {link.Url}, status: {result.Status}, Error: {result.Error}");
					// retrying gateway errors
					if (result.Status == 503 || result.Status == 502 || result.Status == 503)
					{
						Crawl2ExcelLogger.Logger.LogInformation($"Waiting 2 seconds and trying again for url: {link.Url}");
						Thread.Sleep(2000);
					}
				}

				// parsam CSS datoteke i izvlačim sve url-ove
				if (link.ContentType == MimeTypeNames.Css && content.Length > 0)
				{
					string cssContent = Encoding.UTF8.GetString(content);

					var parsedCss = Parser.ParseCSS(cssContent);
					int srcProperyIndex = -1;
					foreach (var pc in parsedCss)
					{
						if (pc.CharacterCategorisation == CharacterCategorisationOptions.SelectorOrStyleProperty && pc.Value == "src")
						{
							srcProperyIndex = pc.IndexInSource;
						}

						if (pc.CharacterCategorisation == CharacterCategorisationOptions.Value && srcProperyIndex != -1 && pc.IndexInSource > srcProperyIndex && pc.Value != null)
						{
							// url value = url("/assets/FontAwesome/webfonts/fa-brands-400.eot?")
							string url = pc.Value.Trim().Replace("url(\"", "").Replace("\")", "");
							if (DecideToCrawlUrl(url).Allow)
							{
								// može biti full url ili relativni, ali može početi i sa ../
								Crawl2ExcelLogger.Logger.LogInformation($"Found css src url: {pc.Value.Trim()}, clean url extracted: {url}");
							}
							srcProperyIndex = -1;
						}
					}
				}

				AddCrawlResult(result);
			}
		}

		private async Task<byte[]> DownloadFileAndFilResult(string url, CrawledPageResult result)
		{
			using (var client = new HttpClient())
			{
				try
				{
					var sw = new Stopwatch();
					sw.Start();
					var content = await client.GetAsync(url);
					var data = await content.Content.ReadAsByteArrayAsync();
					sw.Stop();

					result.Status = (int)content.StatusCode;
					result.TimeMiliseconds = sw.ElapsedMilliseconds;
					result.Size = data.Length;
					return data;
				}
				catch (Exception ex)
				{
					result.Error = ex.Message;
					return new byte[0];
				}

			}
		}

		private void AddCrawlResult(CrawledPageResult result)
		{
			pages.Add(result);
			crawledPagesCount++;
			DisplayCurrentCrawledPagesCount();
			int takePages = 1000;
			if (pages.Count > takePages)
			{
				var storeResults = new List<CrawledPageResult>();
				for (int i = 0; i < takePages; i++)
				{
					storeResults.Add(pages.Take());
				}
				//Parallel.Invoke(() => WriteResultsToExcel(storeResults));
				WriteResultsToExcel(storeResults);
			}
		}

		private void WriteResultsToExcel(List<CrawledPageResult> results)
		{
			var excelWriter = new ExcelWriter(excelResult);
			excelWriter.WriteResults(results);
		}
	}

	public class CrawlerOptions
	{
		public bool OverwriteResultExcelIfExists { get; set;}
		public bool LaunchExcelWhenFinished { get; set;}	
	}
}
