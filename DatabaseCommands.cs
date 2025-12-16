using System.Diagnostics;
using System.Reflection;
using WindowsInput;
using WindowsInput.Native;

namespace NaturalCommands
{
	public class DatabaseCommands
	{
		IInputSimulator inputSimulator = new InputSimulator();
		
		private string FormatDictation(string dictation, string howToFormatDictation)
		{
			if (howToFormatDictation == "Do Nothing")
			{
				return dictation;
			}
			string[] stringSeparators = new string[] { " " };

			string result = "";
			List<string> words = dictation.Split(stringSeparators, StringSplitOptions.None).ToList();
			if (howToFormatDictation == "Camel")
			{
				var counter = 0; string value = "";
				foreach (var word in words)
				{
					counter++;
					if (counter != 1)
					{
						value = value + word.Substring(0, 1).ToUpper() + word.Substring(1).ToLower();
					}
					else
					{
						value += word.ToLower();
					}
					result = value;
				}
			}
			else if (howToFormatDictation == "Variable")
			{
				string value = "";
				foreach (var word in words)
				{
					if (word.Length > 0)
					{
						value = value + word.Substring(0, 1).ToUpper() + word.Substring(1).ToLower();
					}
				}
				result = value;
			}
			else if (howToFormatDictation == "dot notation")
			{
				string value = "";
				foreach (var word in words)
				{
					if (word.Length > 0)
					{
						value = value + word.Substring(0, 1).ToUpper() + word.Substring(1).ToLower() + ".";
					}
				}
				result = value.Substring(0, value.Length - 1);
			}
			else if (howToFormatDictation == "Title")
			{
				string value = "";
				foreach (var word in words)
				{
					if (word.Length > 0)
					{
						value = value + word.Substring(0, 1).ToUpper() + word.Substring(1).ToLower() + " ";
					}
				}
				result = value;
			}
			else if (howToFormatDictation == "Upper")
			{
				string value = "";
				foreach (var word in words)
				{
					value = value + word.ToUpper() + " ";
				}
				result = value;
			}
			else if (howToFormatDictation == "Lower")
			{
				string value = "";
				foreach (var word in words)
				{
					value = value + word.ToLower() + " ";
				}
				result = value;
			}
			return result;
		}
		
		
		public static string RemovePunctuation(string rawResult)
		{
			rawResult = rawResult.Replace(",", "");
			rawResult = rawResult.Replace(";", "");
			rawResult = rawResult.Replace(":", "");
			rawResult = rawResult.Replace("?", "");
			if (rawResult.EndsWith(".")) { rawResult = rawResult.Substring(startIndex: 0, length: rawResult.Length - 1); }
			return rawResult;
		}

	}
}
