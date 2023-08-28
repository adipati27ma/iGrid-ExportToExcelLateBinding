namespace ExportToExcelLateBinding_CS
{
	partial class Form1
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.iGrid1 = new TenTec.Windows.iGridLib.iGrid();
			this.iGrid1DefaultCellStyle = new TenTec.Windows.iGridLib.iGCellStyle(true);
			this.iGrid1DefaultColHdrStyle = new TenTec.Windows.iGridLib.iGColHdrStyle(true);
			this.iGrid1RowTextColCellStyle = new TenTec.Windows.iGridLib.iGCellStyle(true);
			this.buttonExport = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.iGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// iGrid1
			// 
			this.iGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.iGrid1.DefaultCol.CellStyle = this.iGrid1DefaultCellStyle;
			this.iGrid1.DefaultCol.ColHdrStyle = this.iGrid1DefaultColHdrStyle;
			this.iGrid1.Header.Height = 20;
			this.iGrid1.Location = new System.Drawing.Point(12, 12);
			this.iGrid1.Name = "iGrid1";
			this.iGrid1.Size = new System.Drawing.Size(436, 278);
			this.iGrid1.TabIndex = 0;
			// 
			// buttonExport
			// 
			this.buttonExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.buttonExport.Location = new System.Drawing.Point(465, 12);
			this.buttonExport.Name = "buttonExport";
			this.buttonExport.Size = new System.Drawing.Size(126, 35);
			this.buttonExport.TabIndex = 1;
			this.buttonExport.Text = "Export";
			this.buttonExport.UseVisualStyleBackColor = true;
			this.buttonExport.Click += new System.EventHandler(this.buttonExport_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(603, 302);
			this.Controls.Add(this.buttonExport);
			this.Controls.Add(this.iGrid1);
			this.Name = "Form1";
			this.Text = "Export to Excel Sample";
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.iGrid1)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion

		private TenTec.Windows.iGridLib.iGrid iGrid1;
		private TenTec.Windows.iGridLib.iGCellStyle iGrid1DefaultCellStyle;
		private TenTec.Windows.iGridLib.iGColHdrStyle iGrid1DefaultColHdrStyle;
		private TenTec.Windows.iGridLib.iGCellStyle iGrid1RowTextColCellStyle;
		private System.Windows.Forms.Button buttonExport;
	}
}

