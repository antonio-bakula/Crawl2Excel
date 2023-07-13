using System;
using Abot2.Poco;
using AngleSharp.Dom;
using Crawl2Excel.Engine.Models;

namespace Crawl2Excel.Engine.Code
{
	public class PageInfoWorker
	{
		private CrawledPage page;

		public PageInfoWorker(CrawledPage crawledPage)
		{
			page = crawledPage;
		}

		public PageInfo GetInfo()
		{
			var info = new PageInfo();
			info.Charset = page.AngleSharpHtmlDocument.CharacterSet;
			info.Lang = page.AngleSharpHtmlDocument.Head.Language;
			return info;
		}

		public PageSeoInfo GetSeoData()
		{
			var result = new PageSeoInfo();
			result.Title = page.AngleSharpHtmlDocument.Title;
			var metaTags = page.AngleSharpHtmlDocument.Head.GetElementsByTagName("meta");
			result.Description = metaTags.GetMetaTagContent("description");
			result.Keywords = metaTags.GetMetaTagContent("keywords");
			return result;
		}

		public PageOpenGraphInfo GetOpenGraphData()
		{
			var result = new PageOpenGraphInfo();
			var metaTags = page.AngleSharpHtmlDocument.Head.GetElementsByTagName("meta");

			result.Title = metaTags.GetOpenGraphPropertyValue("og:title");
			result.Title = metaTags.GetOpenGraphPropertyValue("og:title");
			result.Description = metaTags.GetOpenGraphPropertyValue("og:description");
			result.Type = metaTags.GetOpenGraphPropertyValue("og:type");
			result.Url = metaTags.GetOpenGraphPropertyValue("og:url");
			result.Image = metaTags.GetOpenGraphPropertyValue("og:image");
			result.SiteName = metaTags.GetOpenGraphPropertyValue("og:site_name");
			return result;
		}

	}

	public static class ElementListExtension
	{
		public static IElement? GetElementByAttributeValue(this IEnumerable<IElement> elements, string name, string value)
		{
			return elements.FirstOrDefault(e => e.Attributes.Any(a => a.Name == name && a.Value == value));
		}

		public static string? GetMetaTagContent(this IEnumerable<IElement> elements, string name)
		{
			return elements.GetElementByAttributeValue("name", name)?.GetAttribute("content");
		}

		public static string? GetOpenGraphPropertyValue(this IEnumerable<IElement> elements, string propertyName)
		{
			return elements.GetElementByAttributeValue("property", propertyName)?.GetAttribute("content");
		}

	}
}
