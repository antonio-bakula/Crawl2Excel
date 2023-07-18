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
			SLDocument excel;
			int row = 1;
			if (!file.Exists)
			{
				excel = new SLDocument();
				excel.AddWorksheet("CrawlResults");
				var headerStyle = excel.CreateStyle();
				headerStyle.Font.Bold = true;
				headerStyle.Fill.SetPattern(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid, Color.Black, Color.LightGray);
				var headerItems = GetHeaderItems();
				int col = 1;
				foreach (var item in headerItems )
				{
					excel.SetCellValue(row, col, item.Title);
					col++;
				}
				excel.SetRowStyle(row, headerStyle);
				row++;
			}
			else
			{
				excel = new SLDocument(file.FullName, "CrawlResults");
				row = excel.GetWorksheetStatistics().EndRowIndex + 1;
			}

			foreach (var cp in results)
			{
				excel.SetCellValue(row, 1, cp.Url);
				excel.SetCellValue(row, 2, cp.Referer);
				excel.SetCellValue(row, 3, cp.Status);
				excel.SetCellValue(row, 4, cp.TimeMiliseconds);
				excel.SetCellValue(row, 5, cp.Size);

				// PAGE INFO
				excel.SetCellValue(row, 6, cp.PageInfo.Charset);
				excel.SetCellValue(row, 7, cp.PageInfo.Lang);

				// SEO
				excel.SetCellValue(row, 8, cp.Seo.Title?.Replace(Environment.NewLine, " ") ?? string.Empty);
				excel.SetCellValue(row, 9, cp.Seo.Description?.Replace(Environment.NewLine, " ") ?? string.Empty);
				excel.SetCellValue(row, 10, cp.Seo.Keywords?.Replace(Environment.NewLine, " ") ?? string.Empty);

				// OPEN GRAPH
				excel.SetCellValue(row, 11, cp.OpenGraph.Title?.Replace(Environment.NewLine, " ") ?? string.Empty);
				excel.SetCellValue(row, 12, cp.OpenGraph.Description?.Replace(Environment.NewLine, " ") ?? string.Empty);
				excel.SetCellValue(row, 13, cp.OpenGraph.Type?.Replace(Environment.NewLine, " ") ?? string.Empty);
				excel.SetCellValue(row, 14, cp.OpenGraph.Url?.Replace(Environment.NewLine, " ") ?? string.Empty);
				excel.SetCellValue(row, 15, cp.OpenGraph.Image?.Replace(Environment.NewLine, " ") ?? string.Empty);
				excel.SetCellValue(row, 16, cp.OpenGraph.SiteName?.Replace(Environment.NewLine, " ") ?? string.Empty);

				excel.SetCellValue(row, 17, cp.Error);
				row++;
			}
			if (file.Exists)
			{
				excel.Save();
			}
			else
			{
				excel.SaveAs(file.FullName);
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
