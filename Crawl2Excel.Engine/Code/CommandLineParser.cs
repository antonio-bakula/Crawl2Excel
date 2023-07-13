using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crawl2Excel.Engine
{
	public class CommandLineParameterAttribute : Attribute
	{
		public int Index { get; set; }

		public bool Mandatory { get; set; }

		public string? Description { get; set; }

		public string? Example { get; set; }

		public CommandLineParameterAttribute()
		{
		}
	}

	public static class CommandLineParser
	{
		public static T Parse<T>(string[] args) where T : AbstractCommandLineParameters
		{
			var result = Activator.CreateInstance<T>();

			result.CheckParsedParameters();
			return result;
		}
	}

	public abstract class AbstractCommandLineParameters
	{
		public List<string> Switches { get; protected set; } = new List<string>();
		public List<string> ParsingErrors {get; protected set; } = new List<string>();

		public bool ParametersValid => ParsingErrors.Any() == false;

		public virtual IEnumerable<string> GetHelpText()
		{
			return new List<string>();
		}

		public abstract void CheckParsedParameters();

		protected IEnumerable<string> GetParametersDescription()
		{
			return new List<string>();
		}

	}

}
