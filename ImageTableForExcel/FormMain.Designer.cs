namespace ImageTableForExcel
{
    partial class FormMain
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.labelRowHeight = new System.Windows.Forms.Label();
            this.numericUpDownRowHeight = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownRowHeight)).BeginInit();
            this.SuspendLayout();
            // 
            // labelRowHeight
            // 
            this.labelRowHeight.AutoSize = true;
            this.labelRowHeight.Location = new System.Drawing.Point(12, 14);
            this.labelRowHeight.Name = "labelRowHeight";
            this.labelRowHeight.Size = new System.Drawing.Size(90, 12);
            this.labelRowHeight.TabIndex = 0;
            this.labelRowHeight.Text = "Row &Height (px):";
            // 
            // numericUpDownRowHeight
            // 
            this.numericUpDownRowHeight.Location = new System.Drawing.Point(108, 12);
            this.numericUpDownRowHeight.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numericUpDownRowHeight.Name = "numericUpDownRowHeight";
            this.numericUpDownRowHeight.Size = new System.Drawing.Size(120, 19);
            this.numericUpDownRowHeight.TabIndex = 1;
            this.numericUpDownRowHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numericUpDownRowHeight.Value = new decimal(new int[] {
            18,
            0,
            0,
            0});
            // 
            // FormMain
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(240, 43);
            this.Controls.Add(this.numericUpDownRowHeight);
            this.Controls.Add(this.labelRowHeight);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.Text = "ImageTableForExcel";
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.dragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.dragEnter);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownRowHeight)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelRowHeight;
        private System.Windows.Forms.NumericUpDown numericUpDownRowHeight;
    }
}

