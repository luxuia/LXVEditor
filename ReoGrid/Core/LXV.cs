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
using System.IO;
using System.Text;

using unvell.ReoGrid.DataFormat;
using unvell.ReoGrid.Interaction;
using unvell.ReoGrid.IO;

namespace unvell.ReoGrid
{
	partial class Worksheet
	{
		public static string LXV_SEP = "∤";

		#region Load

		/// <summary>
		/// Load LXV file into worksheet.
		/// </summary>
		/// <param name="path">File contains LXV data.</param>
		public void LoadLXV(string path)
		{
			LoadLXV(path, RangePosition.EntireRange);
		}

		/// <summary>
		/// Load LXV file into worksheet.
		/// </summary>
		/// <param name="path">File contains LXV data.</param>
		/// <param name="targetRange">The range used to fill loaded LXV data.</param>
		public void LoadLXV(string path, RangePosition targetRange)
		{
			LoadLXV(path, Encoding.Default, targetRange);
		}

		/// <summary>
		/// Load LXV file into worksheet.
		/// </summary>
		/// <param name="path">Path to load LXV file.</param>
		/// <param name="encoding">Encoding used to read and decode plain-text from file.</param>
		public void LoadLXV(string path, Encoding encoding)
		{
			LoadLXV(path, encoding, RangePosition.EntireRange);
		}

		/// <summary>
		/// Load LXV file into worksheet.
		/// </summary>
		/// <param name="path">Path to load LXV file.</param>
		/// <param name="encoding">Encoding used to read and decode plain-text from file.</param>
		/// <param name="targetRange">The range used to fill loaded LXV data.</param>
		public void LoadLXV(string path, Encoding encoding, RangePosition targetRange)
		{
			using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				LoadLXV(fs, encoding, targetRange);
			}
		}

		/// <summary>
		/// Load LXV data from stream into worksheet.
		/// </summary>
		/// <param name="s">Input stream to read LXV data.</param>
		public void LoadLXV(Stream s)
		{
			LoadLXV(s, Encoding.Default);
		}

		/// <summary>
		/// Load LXV data from stream into worksheet.
		/// </summary>
		/// <param name="s">Input stream to read LXV data.</param>
		/// <param name="targetRange">The range used to fill loaded LXV data.</param>
		public void LoadLXV(Stream s, RangePosition targetRange)
		{
			LoadLXV(s, Encoding.Default, targetRange);
		}

		/// <summary>
		/// Load LXV data from stream into worksheet.
		/// </summary>
		/// <param name="s">Input stream to read LXV data.</param>
		/// <param name="encoding">Text encoding used to read and decode plain-text from stream.</param>
		public void LoadLXV(Stream s, Encoding encoding)
		{
			LoadLXV(s, encoding, RangePosition.EntireRange);
		}

		/// <summary>
		/// Load LXV data from stream into worksheet.
		/// </summary>
		/// <param name="s">Input stream to read LXV data.</param>
		/// <param name="encoding">Text encoding used to read and decode plain-text from stream.</param>
		/// <param name="targetRange">The range used to fill loaded LXV data.</param>
		public void LoadLXV(Stream s, Encoding encoding, RangePosition targetRange)
		{
			LoadLXV(s, encoding, targetRange, targetRange.IsEntire, 256);
		}

		/// <summary>
		/// Load LXV data from stream into worksheet.
		/// </summary>
		/// <param name="s">Input stream to read LXV data.</param>
		/// <param name="encoding">Text encoding used to read and decode plain-text from stream.</param>
		/// <param name="targetRange">The range used to fill loaded LXV data.</param>
		/// <param name="autoSpread">decide whether or not to append rows or columns automatically to fill csv data</param>
		/// <param name="bufferLines">decide how many lines int the buffer to read and fill csv data</param>
		public void LoadLXV(Stream s, Encoding encoding, RangePosition targetRange, bool autoSpread, int bufferLines)
		{
			this.controlAdapter?.ChangeCursor(CursorStyle.Busy);

			try
			{
				LXVFileFormatProvider csvProvider = new LXVFileFormatProvider();

				var arg = new LXVFormatArgument
				{
					AutoSpread = autoSpread,
					BufferLines = bufferLines,
					TargetRange = targetRange,
				};

				Clear();

				csvProvider.Load(this.workbook, s, encoding, arg);
			}
			finally
			{
				this.controlAdapter?.ChangeCursor(CursorStyle.PlatformDefault);
			}
		}

		#endregion // Load

		#region Export

		/// <summary>
		/// Export spreadsheet as LXV format from specified number of rows.
		/// </summary>
		/// <param name="path">File path to write LXV format as stream.</param>
		/// <param name="startRow">Number of rows start to export data, 
		/// this property is useful to skip the headers on top of worksheet.</param>
		/// <param name="encoding">Text encoding during output text in LXV format.</param>
		public void ExportAsLXV(string path, int startRow = 0, Encoding encoding = null)
		{
			ExportAsLXV(path, new RangePosition(startRow, 0, -1, -1), encoding);
		}

		/// <summary>
		/// Export spreadsheet as LXV format from specified range.
		/// </summary>
		/// <param name="path">File path to write LXV format as stream.</param>
		/// <param name="addressOrName">Range to be output from this worksheet, specified by address or name.</param>
		/// <param name="encoding">Text encoding during output text in LXV format.</param>
		public void ExportAsLXV(string path, string addressOrName, Encoding encoding = null)
		{
			if (RangePosition.IsValidAddress(addressOrName))
			{
				this.ExportAsLXV(path, new RangePosition(addressOrName), encoding);
			}
			else if (this.TryGetNamedRange(addressOrName, out var namedRange))
			{
				this.ExportAsLXV(path, namedRange, encoding);
			}
			else
			{
				throw new InvalidAddressException(addressOrName);
			}
		}

		/// <summary>
		/// Export spreadsheet as LXV format from specified range.
		/// </summary>
		/// <param name="path">File path to write LXV format as stream.</param>
		/// <param name="range">Range to be output from this worksheet.</param>
		/// <param name="encoding">Text encoding during output text in LXV format.</param>
		public void ExportAsLXV(string path, RangePosition range, Encoding encoding = null)
		{
			using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
			{
				this.ExportAsLXV(fs, range, encoding);
			}
		}

		/// <summary>
		/// Export spreadsheet as LXV format from specified number of rows.
		/// </summary>
		/// <param name="s">Stream to write LXV format as stream.</param>
		/// <param name="startRow">Number of rows start to export data, 
		/// this property is useful to skip the headers on top of worksheet.</param>
		/// <param name="encoding">Text encoding during output text in LXV format</param>
		public void ExportAsLXV(Stream s, int startRow = 0, Encoding encoding = null)
		{
			this.ExportAsLXV(s, new RangePosition(startRow, 0, -1, -1), encoding);
		}

		/// <summary>
		/// Export spreadsheet as LXV format from specified range.
		/// </summary>
		/// <param name="s">Stream to write LXV format as stream.</param>
		/// <param name="addressOrName">Range to be output from this worksheet, specified by address or name.</param>
		/// <param name="encoding">Text encoding during output text in LXV format.</param>
		public void ExportAsLXV(Stream s, string addressOrName, Encoding encoding = null)
		{
			if (RangePosition.IsValidAddress(addressOrName))
			{
				ExportAsLXV(s, new RangePosition(addressOrName), encoding);
			}
			else if (this.TryGetNamedRange(addressOrName, out var namedRange))
			{
				ExportAsLXV(s, namedRange, encoding);
			}
			else
			{
				throw new InvalidAddressException(addressOrName);
			}
		}

		/// <summary>
		/// Export spreadsheet as LXV format from specified range.
		/// </summary>
		/// <param name="s">Stream to write LXV format as stream.</param>
		/// <param name="range">Range to be output from this worksheet.</param>
		/// <param name="encoding">Text encoding during output text in LXV format.</param>
		public void ExportAsLXV(Stream s, RangePosition range, Encoding encoding = null)
		{
			range = FixRange(range);

			int maxRow = Math.Min(range.EndRow, this.MaxContentRow);
			int maxCol = 0;// Math.Min(range.EndCol, this.MaxContentCol);

			if (encoding == null) encoding = Encoding.Default;

			// 检查最大列
			for (int r = range.Row; r <= maxRow; r++) {
				for (int c = range.Col; c <= maxCol;) {
					var cell = this.GetCell(r, c);
					if (cell == null || !cell.IsValidCell) {
						c++;
					} else 
					{
                        c += cell.Colspan;
                        var data = cell.Data;

                        if (data is string str) {
                        } else {
                            str = Convert.ToString(data);
                        }

						if (!string.IsNullOrWhiteSpace(str)) {
							maxCol = Math.Max(maxCol, c+1);
						}
                    }
                    
                }
			}

			using (var sw = new StreamWriter(s, encoding))
			{
				StringBuilder sb = new StringBuilder();

				sb.Append("---sep:" + LXV_SEP);


                for (int r = range.Row; r <= maxRow; r++)
				{
					if (sb.Length > 0)
					{
						sw.WriteLine(sb.ToString());
						sb.Length = 0;
					}

					for (int c = range.Col; c <= maxCol;)
					{
						if (sb.Length > 0)
						{
							sb.Append(LXV_SEP);
						}

						var cell = this.GetCell(r, c);
						if (cell == null || !cell.IsValidCell)
						{
							c++;
						}
						else
						{
							var data = cell.Data;

							bool quota = false;
							//if (!quota)
							//{
							//	if (cell.DataFormat == CellDataFormatFlag.Text)
							//	{
							//		quota = true;
							//	}
							//}

							if (data is string str)
							{
							}
							else
							{
								str = Convert.ToString(data);
							}

							if (quota)
							{
								sb.Append('"');
								sb.Append(str.Replace("\"", "\"\""));
								sb.Append('"');
							}
							else
							{
								sb.Append(str);
							}

							c += cell.Colspan;
						}
					}
				}

				if (sb.Length > 0)
				{
					sw.WriteLine(sb.ToString());
					sb.Length = 0;
				}
			}
		}

		#endregion // Export
	}
}
