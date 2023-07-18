using Crawl2Excel.Engine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crawl2Excel
{
	public class CmdLineParameters : AbstractCommandLineParameters
	{
		[CommandLineParameter(
			Index = 0, 
			Mandatory = true,
			Example = "https://www.antoniob.com/", 
			Description = "The URL of the website where the crawl begins."
		)]
		public required string CrawlStartUrl { get; set; }

		[CommandLineParameter(
			Index = 1, 
			Example = "c:\\temp\\antoniob-crawl-results.xlsx",
			Description = "The Excel file with the results that the crawler will create, if omitted excel file will be created in current directory."
		)]
		public string? ExcelFileName { get; set; }

		public override void CheckParsedParameters()
		{
			if (!string.IsNullOrEmpty(this.CrawlStartUrl) && !Uri.IsWellFormedUriString(this.CrawlStartUrl, UriKind.Absolute))
			{
				ParsingErrors.Add($"CrawlStartUrl error, {this.CrawlStartUrl} is not valid URI");
			}
		}

		public override IEnumerable<string> GetHelpText()
		{
			var result = new List<string>();
			result.Add("Crawl2Excel");
			result.Add("Crawler that crawls web site and extracts SEO data from web pages and saves the results in an MS Excel file.");
			result.Add("");

			result.Add("Parameters:");
			result.AddRange(GetParametersDescription());
			result.Add("");

			result.Add("Switches:");
			result.Add("-help: Prints Help for application.");
			result.Add("-le: Launches Excel and opens the file with the results after the crawl is finished.");
			result.Add("-owr: Overwrites result Excel file if exists.");
			result.Add("");

			result.Add("Usage examples:");
			result.Add("Crawl2Excel https://www.antoniob.com -owr -le");
			result.Add("Crawl2Excel https://www.antoniob.com antoniob.xlsx -owr -le");
			result.Add("Crawl2Excel https://www.antoniob.com c:\\temp\\antoniob.xlsx");

			return result;
		}
	}
}
