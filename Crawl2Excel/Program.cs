using Crawl2Excel;
using Crawl2Excel.Engine;
using Crawl2Excel.Engine.Code;
using System.Diagnostics;

var sw = new Stopwatch();
sw.Start();
var parameters = CommandLineParser.Parse<CmdLineParameters>(args);
if (parameters.Switches.Contains("help"))
{
	var help = parameters.GetHelpText();
	foreach (var line in help)
	{
		Console.WriteLine(line);
	}
	return 0;
}

if (!parameters.ParametersValid)
{
	foreach (var error in parameters.ParsingErrors)
	{
		Console.WriteLine(error);
	}
	Console.WriteLine("");
	Console.WriteLine("See help with switch -help (Crawl2Excel -help)");
	return 0;
}

var crawler = new Crawler(parameters.CrawlStartUrl, parameters.ExcelFileName);
crawler.Options.OverwriteResultExcelIfExists = parameters.Switches.Contains("owr");
crawler.Options.LaunchExcelWhenFinished = parameters.Switches.Contains("le");
await crawler.Start();
sw.Stop();
Console.WriteLine(sw.Elapsed);
return 1;


