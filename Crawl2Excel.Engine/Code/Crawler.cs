using System;
using System.Collections.Concurrent;
using System.Net;
using Abot2.Crawler;
using Abot2.Poco;
using Crawl2Excel.Engine.Models;

namespace Crawl2Excel.Engine.Code
{
	public class Crawler
	{
		private readonly Uri url;
		private readonly FileInfo excelResult;
		private PoliteWebCrawler? abot;
		private BlockingCollection<CrawledPageResult> pages = new BlockingCollection<CrawledPageResult>();
		private ConcurrentDictionary<int, string> crawledPages = new ConcurrentDictionary<int, string>();

		public CrawlerOptions Options { get; private set; } = new CrawlerOptions();
		public CrawlResult Result { get; private set; } = new CrawlResult();

		public Crawler(string startUrl, string? resultExcelFile)
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
			var config = new CrawlConfiguration
			{
				MaxPagesToCrawl = 10000,
				IsExternalPageCrawlingEnabled = false,
				IsExternalPageLinksCrawlingEnabled = false,
				HttpRequestTimeoutInSeconds = 600
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
				/// don't wont data:application URIs
				if ((page?.Uri?.LocalPath ?? "").ToLower().Contains($"data:application/"))
				{
					return new CrawlDecision { Allow = false, Reason = "data:application" };
				}

				/// don't wont data:image URIs
				if ((page?.Uri?.LocalPath ?? "").ToLower().Contains($"data:image/"))
				{
					return new CrawlDecision { Allow = false, Reason = "data:image" };
				}

				return new CrawlDecision { Allow = true };
			};

			Console.Clear();
			Console.WriteLine("Crawling: " + url.ToString());
			Result = await crawler.CrawlAsync(url);

			if (pages.Any())
			{
				WriteResultsToExcel(pages.ToList());
			}
		}

		private void Crawler_PageCrawlDisallowed(object? sender, PageCrawlDisallowedArgs e)
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

		private readonly object _consoleLock = new object();
		private void Crawler_PageCrawlCompleted(object? sender, PageCrawlCompletedArgs e)
		{
			if (e.CrawledPage != null)
			{
				//Parallel.Invoke(() => StoreCrawlData(e.CrawledPage));
				StoreCrawlData(e.CrawledPage);
				lock (_consoleLock)
				{
					Console.SetCursorPosition(0, 2);
					Console.WriteLine($"Pages Crawled: {e.CrawlContext.CrawledCount}");
				}
			}
		}

		private void Crawler_PageCrawlStarting(object? sender, PageCrawlStartingArgs e)
		{
		}

		private void StoreCrawlData(CrawledPage crawledPage, string error = "")
		{
			string? url = crawledPage.Uri?.ToString()?.ToLower();
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
				result.Seo = infoWorker.GetSeoData();
				result.OpenGraph = infoWorker.GetOpenGraphData();
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
		}

		private static readonly object _flushLock = new object();
		
		private void AddCrawlResult(CrawledPageResult result)
		{
			pages.Add(result);
			lock (_consoleLock)
			{
				Console.SetCursorPosition(0, 3);
				Console.WriteLine($"Pages in buffer: {pages.Count}    ");
			}

			lock (_flushLock)
			{
				int takePages = 20;
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
