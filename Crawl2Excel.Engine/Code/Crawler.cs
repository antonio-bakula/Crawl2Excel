using System;
using Abot2.Crawler;
using Crawl2Excel.Engine.Models;

namespace Crawl2Excel.Engine.Code
{
	public class Crawler
	{
		private readonly Uri url;
		private readonly FileInfo excelResult;
		private PoliteWebCrawler? abot;

		public CrawlerOptions Options { get; private set; } = new CrawlerOptions();

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
			//abot = new PoliteWebCrawler(null, null, null, null, null, new Abot.Core.CSQueryHyperlinkParser(), null, null, null);
		}
	}

	public class CrawlerOptions
	{
		public bool OverwriteResultExcelIfExists { get; set;}
		public bool LaunchExcelWhenFinished { get; set;}	
	}
}
