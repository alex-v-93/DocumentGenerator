using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace OpenXml.DocumentGenerator
{
	public class OpenXmlDocumentGenerator
	{
		// #{2}[A-z]+\.+[A-z]+|#{2}[A-z]+
		// #{2}[A-z]+\.+[A-z]+|#{2}[A-z]+[0-9]|#{2}[A-z]+ 
		//
		private static readonly Regex _tagsRegex = new(@"#{2}[A-z]+\.+[A-z]+|#{2}[A-z]+[0-9]|#{2}[A-z]+");
		public Stream GenerateDocument(JObject data, Stream template)
		{
			var result = new MemoryStream();
			template.CopyTo(result);
			result.Seek(0, SeekOrigin.Begin);
			using var doc = WordprocessingDocument.Open(result, true);
			var document = doc.MainDocumentPart.Document;

			ReplaceTags(document, tag => GetValueByTag(tag, data));

			doc.Save();

			return result;
		}

		private string GetValueByTag(string tagPath, JObject data)
		{
			var token = data.SelectToken(tagPath);
			return token?.Value<string>() ?? "";
		}

		private void ReplaceTags(Document document, Func<string, string> valueByTagFunc)
		{
			var paragraphs = document.Body.Elements<Paragraph>().Where(p => _tagsRegex.IsMatch(p.InnerText));
			paragraphs = paragraphs.Union(GetParagraphs(document.Body.Elements<Table>()));
			
			foreach(var paragraph in paragraphs)
			{
				foreach(var run in paragraph.Elements<Run>())
				{
					foreach(var text in run.Elements<Text>())
					{
						ReplaceText(text, valueByTagFunc);
					}
				}
			}
		}

		private IEnumerable<Paragraph> GetParagraphs(IEnumerable<Table> tables)
		{
			foreach(var table in tables)
			{
				foreach(var row in table.Elements<TableRow>())
				{
					if (_tagsRegex.IsMatch(row.InnerText) == false)
					{
						continue;
					}
					foreach (var cell in row.Elements<TableCell>())
					{
						if (_tagsRegex.IsMatch(cell.InnerText) == false)
						{
							continue;
						}
						foreach (var paragraph in cell.Elements<Paragraph>())
						{
							yield return paragraph;
						}
					}
				}
			}
		}

		private void ReplaceText(Text text, Func<string, string> valueByTagFunc)
		{
			var matches = _tagsRegex.Matches(text.Text);
			for (var i = 0; i < matches.Count; i++)
			{
				var match = matches[i];
				var path = match.Value.Substring(2);
				if(string.IsNullOrWhiteSpace(path))
				{
					continue;
				}
				text.Text = text.Text.Replace(match.Value, valueByTagFunc(path));
			}
		}
	}
}
