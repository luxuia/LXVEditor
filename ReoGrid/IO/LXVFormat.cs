/*****************************************************************************
 * 
 * ReoGrid - .NET Spreadsheet Control
 * 
 * https://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Author: Jingwood <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2023 Jingwood <jingwood at unvell.com>
 * Copyright (c) 2012-2023 unvell inc. All rights reserved.
 * 
 ****************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using unvell.ReoGrid;

namespace unvell.ReoGrid.IO
{
	internal static class LXVFormat
	{
		public const int DEFAULT_READ_BUFFER_LINES = 512;
		
		private static Regex lineRegex = new Regex("\\s*(\\\"(?<item>[^\\\"]*)\\\"|(?<item>[^,]*))\\s*,?", RegexOptions.Compiled);

		public static string Read(StreamReader sr, Worksheet sheet, RangePosition targetRange, 
			Encoding encoding = null, int bufferLines = DEFAULT_READ_BUFFER_LINES, bool autoSpread = true)
		{
			targetRange = sheet.FixRange(targetRange);

			string[] lines = new string[bufferLines];
			List<object>[] bufferLineList = new List<object>[bufferLines];

			for (int i = 0; i < bufferLineList.Length; i++)
			{
				bufferLineList[i] = new List<object>(256);
			}

#if DEBUG
			var sw = System.Diagnostics.Stopwatch.StartNew();
#endif

			int row = targetRange.Row;
			int totalReadLines = 0;

			sheet.SuspendDataChangedEvents();
            int maxCols = 0;
            string nextline = null;

            try {
                bool finished = false;
                while (!finished) {
                    int readLines = 0;


                    for (; readLines < lines.Length; readLines++) {
                        var line = sr.ReadLine();
                        if (line == null) {
                            finished = true;
                            break;
                        }
                        if (line.StartsWith("---")) {
                            finished = true;
                            nextline = line;
                            break;
                        }

                        lines[readLines] = line;

                        totalReadLines++;
                        if (!autoSpread && totalReadLines > targetRange.Rows) {
                            finished = true;
                            break;
                        }
                    }

                    if (autoSpread && row + readLines > sheet.RowCount) {
                        int appendRows = bufferLines - (sheet.RowCount % bufferLines);
                        if (appendRows <= 0) appendRows = bufferLines;
                        sheet.AppendRows(appendRows);
                    }

                    for (int i = 0; i < readLines; i++) {
                        var line = lines[i];

                        var toBuffer = bufferLineList[i];
                        toBuffer.Clear();

                        var items = line.Split(new string[] { Worksheet.LXV_SEP }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var item in items) {
                            toBuffer.Add(item);

                            if (toBuffer.Count >= targetRange.Cols) break;
                        }

                        if (maxCols < toBuffer.Count) maxCols = toBuffer.Count;

                        if (autoSpread && maxCols >= sheet.ColumnCount) {
                            sheet.SetCols(maxCols + 1);
                        }
                    }

                    sheet.SetRangeData(row, targetRange.Col, readLines, maxCols, bufferLineList);
                    row += readLines;
                }
            } finally {
                sheet.ResumeDataChangedEvents();
            }

            sheet.RaiseRangeDataChangedEvent(new RangePosition(
                targetRange.Row, targetRange.Col, maxCols, totalReadLines));

            return nextline;
        }
    }

    #region LXV File Provider
    internal class LXVFileFormatProvider : IFileFormatProvider
	{
		public bool IsValidFormat(string file)
		{
			return System.IO.Path.GetExtension(file).Equals(".lxv", StringComparison.CurrentCultureIgnoreCase);
		}

		public bool IsValidFormat(Stream s)
		{
			throw new NotSupportedException();
		}

		public void Load(IWorkbook workbook, Stream stream, Encoding encoding, object arg)
		{
			bool autoSpread = true;
			int bufferLines = LXVFormat.DEFAULT_READ_BUFFER_LINES;
			RangePosition targetRange = RangePosition.EntireRange;

			LXVFormatArgument csvArg = arg as LXVFormatArgument;

			if (csvArg != null)
			{
				autoSpread = csvArg.AutoSpread;
				bufferLines = csvArg.BufferLines;
				targetRange = csvArg.TargetRange;
			}
		
			workbook.Worksheets.Clear();
			using (var sr = new StreamReader(stream, encoding)) {
				var speline = sr.ReadLine();

                string nextline = null;
                bool isfinish = false;
                while (!isfinish) {
                    var sheetnameline = "";
                    if (nextline != null) {
                        if (nextline.StartsWith("---sheet")) {
                            sheetnameline = nextline;
                            nextline = null;
                        } else {
                            // DO_FINAL_DATA
                        }
                    } else {
                       sheetnameline  = sr.ReadLine();
                    }
                    if (sheetnameline != null) {
                        var sheetname = sheetnameline.Replace("---sheet:", "");

                        var sheet = new Worksheet(workbook as Workbook, sheetname);
                        workbook.Worksheets.Add(sheet);
                        nextline = LXVFormat.Read(sr, sheet, targetRange, encoding, bufferLines, autoSpread);

                        if (nextline == null) {
                            isfinish = true;
                        }
                    }
                }
			}
		}

		public void SaveSheet(Worksheet worksheet, StreamWriter sw) {

			var maxRow = worksheet.RowCount;
			var maxCol = 0;
            // 检查最大列
            for (int r = 0; r <= maxRow; r++) {
                for (int c = 0; c <= worksheet.ColumnCount;) {
                    var cell = worksheet.GetCell(r, c);
                    if (cell == null || !cell.IsValidCell) {
                        c++;
                    } else {
                        c += cell.Colspan;
                        var data = cell.Data;

                        if (data is string str) {
                        } else {
                            str = Convert.ToString(data);
                        }

                        if (!string.IsNullOrWhiteSpace(str)) {
                            maxCol = Math.Max(maxCol, c + 1);
                        }
                    }
                }
            }

			{
                StringBuilder sb = new StringBuilder();

                sb.Append("---sheet:" + worksheet.Name);


                for (int r = 0; r <= maxRow; r++) {
                    if (sb.Length > 0) {
                        sw.WriteLine(sb.ToString());
                        sb.Length = 0;
                    }

                    for (int c = 0; c <= maxCol;) {
                        if (sb.Length > 0) {
                            sb.Append(Worksheet.LXV_SEP);
                        }

                        var cell = worksheet.GetCell(r, c);
                        if (cell == null || !cell.IsValidCell) {
                            c++;
                        } else {
                            var data = cell.Data;

                            if (data is string str) {
                            } else {
                                str = Convert.ToString(data);
                            }

                            str = str.Replace("\r", "");
                            str = str.Replace("\n", "\\n");
                            str = str.Replace("\t", "\\t");

                            sb.Append(str);

                            c += cell.Colspan;
                        }
                    }
                }

                if (sb.Length > 0) {
                    sw.WriteLine(sb.ToString());
                    sb.Length = 0;
                }
            }
        }

		public void Save(IWorkbook workbook, Stream stream, Encoding encoding, object arg)
		{
		
			if (encoding == null) encoding = Encoding.Default;

            using (var sw = new StreamWriter(stream, encoding)) {
                sw.WriteLine("---sep:" + Worksheet.LXV_SEP);

				foreach (var sheet in workbook.Worksheets) {
					SaveSheet(sheet, sw);
                }
            }
        }
        //int fromRow = 0, fromCol = 0, toRow = 0, toCol = 0;

        //if (args != null)
        //{
        //	object arg;
        //	if (args.TryGetValue("fromRow", out arg)) fromRow = (int)arg;
        //	if (args.TryGetValue("fromCol", out arg)) fromCol = (int)arg;
        //	if (args.TryGetValue("toRow", out arg)) toRow = (int)arg;
        //	if (args.TryGetValue("toCol", out arg)) toCol = (int)arg;
        //}
    }

	/// <summary>
	/// Arguments for loading and saving LXV format.
	/// </summary>
	public class LXVFormatArgument
	{
		/// <summary>
		/// Determines whether or not allow to expand worksheet to load more data from LXV file. (Default is True)
		/// </summary>
		public bool AutoSpread { get; set; }

		/// <summary>
		/// Determines how many rows read from LXV file one time. (Default is LXVFormat.DEFAULT_READ_BUFFER_LINES = 512)
		/// </summary>
		public int BufferLines { get; set; }


		/// <summary>
		/// Determines where to start import LXV data on worksheet.
		/// </summary>
		public RangePosition TargetRange { get; set; }

		/// <summary>
		/// Create the argument object instance.
		/// </summary>
		public LXVFormatArgument()
		{
			this.AutoSpread = true;
			this.BufferLines = LXVFormat.DEFAULT_READ_BUFFER_LINES;
			this.TargetRange = RangePosition.EntireRange;
		}
	}
	#endregion // LXV File Provider

}
