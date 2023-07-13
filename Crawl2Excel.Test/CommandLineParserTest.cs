using Crawl2Excel.Engine.Models;
using Crawl2Excel.Engine;

namespace Crawl2Excel.Test
{
	[TestClass]
	public class CommandLineParserTest
	{
		[TestMethod]
		public void ParseTest1()
		{
			var pms = CreateParameters("pm0 pm1 -mu -gaa -go_crazy");
			Assert.IsNotNull(pms);
			Assert.AreEqual<string>("pm0", pms.Parameter0);
			Assert.IsNotNull(pms.Parameter1);
			Assert.AreEqual<string>("pm1", pms.Parameter1);
			Assert.IsTrue(pms.Switches.Contains("mu"));
			Assert.IsTrue(pms.Switches.Contains("gaa"));
			Assert.IsTrue(pms.Switches.Contains("go_crazy"));
		}

		[TestMethod]
		public void ParseTest2()
		{
			var pms = CreateParameters("pm0 -gaa -go_crazy");
			Assert.IsNotNull(pms);
			Assert.AreEqual<string>("pm0", pms.Parameter0);
			Assert.IsNull(pms.Parameter1);
			Assert.IsFalse(pms.Switches.Contains("mu"));
			Assert.IsTrue(pms.Switches.Contains("gaa"));
			Assert.IsTrue(pms.Switches.Contains("go_crazy"));
		}

		private TestCmdLineParameters CreateParameters(string commandLine)
		{
			return CommandLineParser.Parse<TestCmdLineParameters>(commandLine.Split(' '));
		}
	}

	public class TestCmdLineParameters : AbstractCommandLineParameters
	{
		[CommandLineParameter(
			Index = 0,
			Mandatory = true,
			Example = "pm0",
			Description = "This is the description of the first parameter."
		)]
		public required string Parameter0 { get; set; }

		[CommandLineParameter(
			Index = 1,
			Example = "pm1",
			Description = "This is the description of the second parameter."
		)]
		public string? Parameter1 { get; set; }

		public override void CheckParsedParameters()
		{ 
		
		}

	}
}