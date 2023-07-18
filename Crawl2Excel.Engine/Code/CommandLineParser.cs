using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
			Type resultType = result.GetType();

			var allParameters = resultType.GetProperties()
				.Where(p => p.GetCustomAttribute<CommandLineParameterAttribute>() != null)
				.Select(p => new 
				{ 
					Property = p, 
					Attribute = p.GetCustomAttribute<CommandLineParameterAttribute>() 
				});

			// fill parameters
			int index = 0;
			foreach (var arg in args.Where(a => !string.IsNullOrEmpty(a)))
			{
				if (arg.StartsWith("-"))
				{
					result.Switches.Add(arg.Substring(1));
				}
				else
				{
					var myParam = allParameters.FirstOrDefault(i => i.Attribute?.Index == index);
					if (myParam != null)
					{
						myParam.Property.SetValue(result, arg);
					}
					index++;
				}
			}

			// check mandatory parameters
			var mandatoryParameters = allParameters.Where(i => i.Attribute?.Mandatory ?? false);
			foreach (var mpar in mandatoryParameters)
			{
				var value = mpar.Property?.GetValue(result);
				if (value == null)
				{
					result.ParsingErrors.Add($"Mandatory parameter {mpar.Property?.Name} not set.");
				}
			}

			result.CheckParsedParameters();
			return result;
		}
	}

	public abstract class AbstractCommandLineParameters
	{
		public List<string> Switches { get; protected set; } = new List<string>();
		public List<string> ParsingErrors { get; protected set; } = new List<string>();

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
