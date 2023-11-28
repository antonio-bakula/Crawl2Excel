using Crawl2Excel;
using Crawl2Excel.Engine;
using Crawl2Excel.Engine.Code;
using Karambolo.Extensions.Logging.File;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
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

var services = new ServiceCollection();

services.AddLogging(builder =>
{
	builder.AddFile(o => 
	{
		o.RootPath = AppContext.BaseDirectory;
		o.FileAccessMode = LogFileAccessMode.OpenTemporarily;
		o.FileEncoding = System.Text.Encoding.UTF8;
		o.DateFormat = "dd.MM.yyyy HH:mm:ss";
		o.CounterFormat = "000";
		o.MaxFileSize = 1024 * 1024 * 10;
		o.Files = new[]
		{
			new LogFileOptions
			{
				Path = "Crawl2Excel-<counter>.log",
			}
		};
	});
});


var sp = services.BuildServiceProvider();
var loggerFactory = sp.GetService<ILoggerFactory>();
Crawl2ExcelLogger.Logger = loggerFactory.CreateLogger("Crawl2Excel");

Crawl2ExcelLogger.Logger.LogInformation($"Crawl2Excel started. Start Url: {parameters.CrawlStartUrl}");

var crawler = new Crawler(parameters.CrawlStartUrl, parameters.ExcelFileName);
crawler.Options.OverwriteResultExcelIfExists = parameters.Switches.Contains("owr");
crawler.Options.LaunchExcelWhenFinished = parameters.Switches.Contains("le");
await crawler.Start();
sw.Stop();
Console.WriteLine(sw.Elapsed);
sp.Dispose();
return 1;


