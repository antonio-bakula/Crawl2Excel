using Abot2.Poco;
using HeyRed.Mime;

namespace Crawl2Excel.Engine.Code
{
	public class HtmlContentParser
	{
		private CrawledPage page;

		public List<ParsedLink> ParsedLinks { get; private set; } = new List<ParsedLink>();

		public HtmlContentParser(CrawledPage crawledPage)
		{
			page = crawledPage;
		}

		public void ParseContent()
		{
			foreach (var image in page.AngleSharpHtmlDocument.Images)
			{
				if (IsLocalUrl(image.Source))
				{
					string fixedUrl = FixUrl(image.Source);	
					ParsedLinks.Add(new ParsedLink(fixedUrl, page.Uri.ToString()));
				}
			}

			foreach (var script in page.AngleSharpHtmlDocument.Scripts)
			{
				if (IsLocalUrl(script.Source))
				{
					string fixedUrl = FixUrl(script.Source);
					ParsedLinks.Add(new ParsedLink(fixedUrl, page.Uri.ToString()));
				}
			}

			foreach (var css in page.AngleSharpHtmlDocument.StyleSheets)
			{
				if (IsLocalUrl(css.Href))
				{
					string fixedUrl = FixUrl(css.Href);
					ParsedLinks.Add(new ParsedLink(fixedUrl, page.Uri.ToString()));
				}
			}

			var links = page.AngleSharpHtmlDocument.GetElementsByTagName("link");
			foreach (var link in links)
			{
				string linkUrl = link.GetAttribute("href");
				if (IsLocalUrl(linkUrl))
				{
					string fixedUrl = FixUrl(linkUrl);
					ParsedLinks.Add(new ParsedLink(fixedUrl, page.Uri.ToString()));
				}
			}
		}

		private bool IsLocalUrl(string url)
		{
			if (string.IsNullOrEmpty(url) || url.StartsWith("//") || url.ToLower().StartsWith("http://") || url.ToLower().StartsWith("https://"))
			{
				return false;
			}
			
			if (url.StartsWith("about:///"))
			{
				return true;
			}
			
			Uri uri = new Uri(url, UriKind.RelativeOrAbsolute);
			return !uri.IsAbsoluteUri || uri.Host == page.Uri.Host;

		}

		private string FixUrl(string url)
		{
			string fxed = url.Replace("about://", "");
			if (!fxed.StartsWith("/"))
			{
				fxed = "/" + fxed;
			}
			return page.Uri.GetLeftPart(UriPartial.Authority) + fxed;
		}
	}

	public class ParsedLink
	{
		public string Url { get; set; }
		public string ContentType { get; set; }
		public string Referer { get; set; }

		public ParsedLink(string url, string referer)
		{
			this.Url = url;
			this.ContentType = MimeTypesMap.GetMimeType(url);
			this.Referer = referer;
		}
	}
}
