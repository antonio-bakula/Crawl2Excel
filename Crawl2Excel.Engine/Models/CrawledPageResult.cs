using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crawl2Excel.Engine.Models
{
	public class CrawledPageResult
	{
		public string? Url { get; set; }
		public string? Referer { get; set; }
		public int Status { get; set; }
		public long TimeMiliseconds { get; set; }
		public long Size { get; set; }
		public string? Error { get; set; }
		public PageInfo PageInfo { get; set; } = new PageInfo();
		public PageSeoInfo Seo { get; set; } = new PageSeoInfo();
		public PageOpenGraphInfo OpenGraph { get; set; } = new PageOpenGraphInfo();
	}
}
