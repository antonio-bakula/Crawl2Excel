using System;

namespace Crawl2Excel.Engine.Models
{

	public class PageInfo
	{
		public bool IsHtml => this.ContentType?.ToLower().Contains("html") ?? false;
		public string? Charset { get; set; }
		public string? Lang { get; set; }
		public string? ContentType { get; set; }
	}

	public class PageSeoInfo
	{
		public string? Title { get; set; }
		public string? Description { get; set; }
		public string? Keywords { get; set; }
	}

	public class PageOpenGraphInfo
	{
		public string? Title { get; set; }
		public string? Description { get; set; }
		public string? Type { get; set; }
		public string? Url { get; set; }
		public string? Image { get; set; }
		public string? SiteName { get; set; }
	}

}
