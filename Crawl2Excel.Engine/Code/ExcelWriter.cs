using Crawl2Excel.Engine.Models;
using OfficeOpenXml;
using System;

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
			int col = 1;
			if (!File.Exists(file.FullName))
			{
				using (var pack = new ExcelPackage())
				{
					var ws = pack.Workbook.Worksheets.Add("CrawlResults");
					var headers = GetHeaderItems();
					col = 1;
					foreach (var header in headers)
					{
						ws.Cells[row, col].Style.Font.Bold = true;
						ws.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
						ws.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
						ws.Cells[row, col].Value = header.Title;
						col++;
					}
					pack.SaveAs(file);
				}
			}

			using (var pack = new ExcelPackage(file))
			{
				var ws = pack.Workbook.Worksheets["CrawlResults"];
				row = ws.Dimension.End.Row + 1;
				foreach (var cp in results)
				{
					col = 1;
					foreach (var value in GetDataItems(cp))
					{
						ws.Cells[row, col].Value = value;
						col++;
					}
					row++;
				}

				var headers = GetHeaderItems();
				col = 1;
				foreach (var header in headers)
				{
					if (header.AutoFit)
					{
						if (header.AutoFitMinWidth.HasValue && header.AutoFitMaxWidth.HasValue)
						{
							ws.Column(col).AutoFit(header.AutoFitMinWidth.Value, header.AutoFitMaxWidth.Value);
						}
						else if (header.AutoFitMinWidth.HasValue)
						{
							ws.Column(col).AutoFit(header.AutoFitMinWidth.Value);
						}
						else
						{
							ws.Column(col).AutoFit();
						}
					}
					else if (header.Width.HasValue)
					{
						ws.Column(col).Width = header.Width.Value;
					}

					if (!string.IsNullOrEmpty(header.NumberFormat))
					{
						ws.Column(col).Style.Numberformat.Format = header.NumberFormat;
					}
					col++;
				}
				ws.View.FreezePanes(2, 1);
				pack.Save();
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
			items.Add(new ExcelColumnInfo { Title = "Charset", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 30 });
			items.Add(new ExcelColumnInfo { Title = "Lang", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 30 });
			items.Add(new ExcelColumnInfo { Title = "ContentType", AutoFit = true, AutoFitMinWidth = 10, AutoFitMaxWidth = 32 });

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

		private List<object?> GetDataItems(CrawledPageResult cp)
		{
			var items = new List<object?>();
			items.Add(cp.Url);
			items.Add(cp.Referer);
			items.Add(cp.Status);
			items.Add(cp.TimeMiliseconds);
			items.Add(cp.Size);

			// PAGE INFO
			items.Add(cp.PageInfo.Charset);
			items.Add(cp.PageInfo.Lang);
			items.Add(cp.PageInfo.ContentType);
				
			// SEO
			items.Add(cp.Seo.Title?.Replace(Environment.NewLine, " ") ?? string.Empty);
			items.Add(cp.Seo.Description?.Replace(Environment.NewLine, " ") ?? string.Empty);
			items.Add(cp.Seo.Keywords?.Replace(Environment.NewLine, " ") ?? string.Empty);
				
			// OPEN GRAPH
			items.Add(cp.OpenGraph.Title?.Replace(Environment.NewLine, " ") ?? string.Empty);
			items.Add(cp.OpenGraph.Description?.Replace(Environment.NewLine, " ") ?? string.Empty);
			items.Add(cp.OpenGraph.Type?.Replace(Environment.NewLine, " ") ?? string.Empty);
			items.Add(cp.OpenGraph.Url?.Replace(Environment.NewLine, " ") ?? string.Empty);
			items.Add(cp.OpenGraph.Image?.Replace(Environment.NewLine, " ") ?? string.Empty);
			items.Add(cp.OpenGraph.SiteName?.Replace(Environment.NewLine, " ") ?? string.Empty);

			items.Add(cp.Error);
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
