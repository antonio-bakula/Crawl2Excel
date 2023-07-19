using Abot2.Poco;
using Crawl2Excel.Engine.Models;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crawl2Excel.Engine.Code
{
	public class ExcelWriter
	{
		private FileInfo file;

		public ExcelWriter(FileInfo excelFile) 
		{
			file = excelFile;
		}

		public void WriteResults(List<CrawledPageResult> results)
		{
			int row = 1;
			file.Refresh();
			if (!file.Exists)
			{
				using (var excel = new SLDocument())
				{
					excel.RenameWorksheet("Sheet1", "CrawlResults");
					var headerStyle = excel.CreateStyle();
					headerStyle.Font.Bold = true;
					headerStyle.Font.FontColor = Color.Black;
					headerStyle.Fill.SetPattern(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid, Color.LightGray, Color.White);
					var headerItems = GetHeaderItems();
					int col = 1;
					foreach (var item in headerItems)
					{
						excel.SetCellValue(row, col, item.Title);
						excel.SetCellStyle(row, col, headerStyle);
						col++;
					}
					row++;
					excel.SaveAs(file.FullName);
				}
			}

			using (var excel = new SLDocument(file.FullName, "CrawlResults"))
			{
				row = excel.GetWorksheetStatistics().EndRowIndex + 1;

				foreach (var cp in results)
				{
					int col = 1;
					excel.SetCellValue(row, col++, cp.Url);
					excel.SetCellValue(row, col++, cp.Referer);
					excel.SetCellValue(row, col++, cp.Status);
					excel.SetCellValue(row, col++, cp.TimeMiliseconds);
					excel.SetCellValue(row, col++, cp.Size);

					// PAGE INFO
					excel.SetCellValue(row, col++, cp.PageInfo.Charset);
					excel.SetCellValue(row, col++, cp.PageInfo.Lang);
					excel.SetCellValue(row, col++, cp.PageInfo.ContentType);

					// SEO
					excel.SetCellValue(row, col++, cp.Seo.Title?.Replace(Environment.NewLine, " ") ?? string.Empty);
					excel.SetCellValue(row, col++, cp.Seo.Description?.Replace(Environment.NewLine, " ") ?? string.Empty);
					excel.SetCellValue(row, col++, cp.Seo.Keywords?.Replace(Environment.NewLine, " ") ?? string.Empty);

					// OPEN GRAPH
					excel.SetCellValue(row, col++, cp.OpenGraph.Title?.Replace(Environment.NewLine, " ") ?? string.Empty);
					excel.SetCellValue(row, col++, cp.OpenGraph.Description?.Replace(Environment.NewLine, " ") ?? string.Empty);
					excel.SetCellValue(row, col++, cp.OpenGraph.Type?.Replace(Environment.NewLine, " ") ?? string.Empty);
					excel.SetCellValue(row, col++, cp.OpenGraph.Url?.Replace(Environment.NewLine, " ") ?? string.Empty);
					excel.SetCellValue(row, col++, cp.OpenGraph.Image?.Replace(Environment.NewLine, " ") ?? string.Empty);
					excel.SetCellValue(row, col++, cp.OpenGraph.SiteName?.Replace(Environment.NewLine, " ") ?? string.Empty);

					excel.SetCellValue(row, col++, cp.Error);
					row++;
				}
				excel.Save();
			}
		}

		private List<ExcelColumnInfo> GetHeaderItems()
		{
			var items = new List<ExcelColumnInfo>();
			items.Add(new ExcelColumnInfo { Title = "Url", Width = 100 });
			items.Add(new ExcelColumnInfo { Title = "Referer", Width = 50 });
			items.Add(new ExcelColumnInfo { Title = "Status", AutoFit = true });
			items.Add(new ExcelColumnInfo { Title = "Time (ms)", AutoFit = true });
			items.Add(new ExcelColumnInfo { Title = "Size (bytes)", AutoFit = true, NumberFormat = "#,##0" });

			// PAGE INFO
			items.Add(new ExcelColumnInfo { Title = "Charset", Width = 100 });
			items.Add(new ExcelColumnInfo { Title = "Lang", Width = 100 });
			items.Add(new ExcelColumnInfo { Title = "ContentType", Width = 100 });

			// SEO
			items.Add(new ExcelColumnInfo { Title = "SeoTitle", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 50 });
			items.Add(new ExcelColumnInfo { Title = "SeoDescription", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 50 });
			items.Add(new ExcelColumnInfo { Title = "SeoKeywords", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 50 });

			// OPEN GRAPH
			items.Add(new ExcelColumnInfo { Title = "OgTitle", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 50 });
			items.Add(new ExcelColumnInfo { Title = "OgDescription", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 50 });
			items.Add(new ExcelColumnInfo { Title = "OgType", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 50 });
			items.Add(new ExcelColumnInfo { Title = "OgUrl", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 50 });
			items.Add(new ExcelColumnInfo { Title = "OgImage", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 50 });
			items.Add(new ExcelColumnInfo { Title = "OgSiteName", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 50 });

			items.Add(new ExcelColumnInfo { Title = "Error", AutoFit = true });
			return items;
		}

	}

	public class ExcelColumnInfo
	{
		public string? Title { get; set; }
		public double? Width { get; set; }
		public bool AutoFit { get; set; }
		public double? AutoFitMinWidth { get; set; }
		public double? AutoFitMaxWidth { get; set; }
		public string? NumberFormat { get; set; }
	}
}
