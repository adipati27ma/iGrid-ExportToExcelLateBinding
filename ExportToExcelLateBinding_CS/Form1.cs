using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TenTec.Windows.iGridLib;

namespace ExportToExcelLateBinding_CS
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			// Disable redrawing for higher speed.
			iGrid1.BeginUpdate();

			// Initialize the grid.
			iGrid1.GroupBox.Visible = true;
			iGrid1.GroupRowLevelStyles[0].ForeColor = Color.White;
			iGrid1.GroupRowLevelStyles[0].BackColor = Color.SandyBrown;

			// Add columns.
			iGrid1.Cols.AddRange(8);
			foreach (iGCol myCol in iGrid1.Cols)
				myCol.Text = $"Col{myCol.Index + 1}";

			// Test column visibility.
			iGrid1.Cols[5].Visible = false;

			// Test cell colors.
			iGrid1.Cols[7].CellStyle.ForeColor = Color.Yellow;
			iGrid1.Cols[7].CellStyle.BackColor = Color.DeepSkyBlue;

			// Add rows.
			iGrid1.Rows.AddRange(40);

			// Test row visibility.
			iGrid1.Rows[2].Visible = false;

			// Populate cells with test data.
			Random myRandom = new Random();
			foreach (iGCell myCell in iGrid1.Cells)
			{
				if (myCell.ColIndex <= 2)
					myCell.Value = 100 * (myCell.ColIndex + 1) + myRandom.Next(1, 6);
				else
					myCell.Value = $"C{myCell.ColIndex+1}R{myCell.RowIndex+1}";
			}

			// Group the grid by two columns.
			iGrid1.GroupObject.Add(0);
			iGrid1.GroupObject.Add(1);
			iGrid1.Group();

			// Enable redrawing.
			iGrid1.EndUpdate();
		}

		private void buttonExport_Click(object sender, EventArgs e)
		{
			iGExportToExcelManager.Export(iGrid1);
		}
	}
}
