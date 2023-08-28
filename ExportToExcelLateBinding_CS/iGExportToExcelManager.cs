using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TenTec.Windows.iGridLib;

namespace ExportToExcelLateBinding_CS
{
	class iGExportToExcelManager
	{
		private const int xlContinuous = 1;

		public static void Export(iGrid grid)
		{
			// Create an instance of Excel using late binding.
			var excelType = Type.GetTypeFromProgID("Excel.Application");
			if (excelType == null)
			{
				MessageBox.Show("Excel Application not found.\n\rIt is required to export iGrid.", "iGExportToExcelManager", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			dynamic xlApp = Activator.CreateInstance(excelType);

			// Create a new workbook.
			// IMPORTANT: Never use two dots with COM objects like xlApp.Workbooks.Add() - otherwise
			// you will get internal references to COM objects that will hold Excel.exe in memory.
			// See also the cleanup code with Marshal.ReleaseComObject() calls.
			dynamic xlWorkbooks = xlApp.Workbooks;
			dynamic xlWorkbook = xlWorkbooks.Add();

			#region Export the contents of iGrid

			// Determine column export order.
			// colExportOrder contains Excel column indices for corresponding grid columns.
			int[] colExportOrder = new int[grid.Cols.Count];
			for (int groupIdx = 0; groupIdx < grid.GroupObject.Count; groupIdx++)
			{
				colExportOrder[grid.GroupObject[groupIdx].ColIndex] = groupIdx + 1;
			}
			int skipColCount = 0;
			for (int colPos = 0; colPos < grid.Cols.Count; colPos++)
			{
				int colIndex = grid.Cols.FromOrder(colPos).Index;
				if (!IsGridColHdrVisible(grid, colIndex))
					skipColCount++;
				else
					colExportOrder[colIndex] = grid.GroupObject.Count + colPos + 1 - skipColCount;
			}

			// Export column headers.
			foreach (iGCol col in grid.Cols)
			{
				if (col.Visible)
				{
					dynamic xlCellColHdr = xlApp.Cells(1, colExportOrder[col.Index]);
					xlCellColHdr.Value = col.Text;
					xlCellColHdr.Borders.LineStyle = xlContinuous;
					xlCellColHdr.Font.Bold = true;
					xlCellColHdr.Interior.Color = Color.LightGray;
				}
			}

			// The row start index for cells in Excel.
			// Set it to 1 if column headers are not exported.
			int exportRowIndex = 2;

			// Export rows with cells.
			foreach (iGRow row in grid.Rows)
			{
				if (row.Visible)
				{
					switch (row.Type)
					{
						case iGRowType.AutoGroupRow:
						case iGRowType.ManualGroupRow:

							dynamic xlCellGroup = xlApp.Cells(exportRowIndex, row.Level + 1);
							dynamic xlRangeGroup = xlApp.Range(xlCellGroup, xlApp.Cells(exportRowIndex, grid.Cols.Count - skipColCount + grid.GroupObject.Count));
							xlRangeGroup.Merge();

							xlRangeGroup.Borders.LineStyle = xlContinuous;

							iGCell rowTextCell = row.RowTextCell;
							xlCellGroup.Value = rowTextCell.Value;
							SetExcelCellColors(xlCellGroup, rowTextCell.EffectiveForeColor, rowTextCell.EffectiveBackColor);

							break;

						case iGRowType.Normal:

							foreach (iGCell cell in row.Cells)
							{
								if (IsGridColHdrVisible(grid, cell.ColIndex))
								{
									dynamic xlCellCell = xlApp.Cells(exportRowIndex, colExportOrder[cell.ColIndex]);
									xlCellCell.Value = cell.Value;
									xlCellCell.Borders.LineStyle = xlContinuous;
									SetExcelCellColors(xlCellCell, cell.EffectiveForeColor, cell.EffectiveBackColor);
								}
							}

							break;
					}
					exportRowIndex++;
				}
			}

			#endregion

			// Freeze the top row with column headers.
			dynamic xlActiveWin = xlApp.ActiveWindow;
			xlActiveWin.SplitRow = 1;
			xlActiveWin.FreezePanes = true;

			// Make the Excel app visible.
			xlApp.Visible = true;

			// Clean up the references to the COM objects to release Excel.exe from memory.
			System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbooks);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
		}

		private static int GetExcelColor(Color color)
		{
			return color.R + color.G * 256 + color.B * 65536;
		}

		private static void SetExcelCellColors(dynamic xlCell, Color foreColor, Color backColor)
		{
			if (!foreColor.IsEmpty)
				xlCell.Font.Color = GetExcelColor(foreColor);
			if (!backColor.IsEmpty)
				xlCell.Interior.Color = GetExcelColor(backColor);
		}

		private static bool IsGridColHdrVisible(iGrid grid, int colIndex)
		{
			return grid.Cols[colIndex].Visible && !grid.GroupObject.Contains(colIndex);
		}
	}
}
