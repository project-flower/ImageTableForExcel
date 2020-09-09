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
            this.labelVerticalSpacing = new System.Windows.Forms.Label();
            this.numericUpDownVerticalSpacing = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownVerticalSpacing)).BeginInit();
            this.SuspendLayout();
            // 
            // labelVerticalSpacing
            // 
            this.labelVerticalSpacing.AutoSize = true;
            this.labelVerticalSpacing.Location = new System.Drawing.Point(12, 14);
            this.labelVerticalSpacing.Name = "labelVerticalSpacing";
            this.labelVerticalSpacing.Size = new System.Drawing.Size(127, 12);
            this.labelVerticalSpacing.TabIndex = 0;
            this.labelVerticalSpacing.Text = "Vertical &Spacing (rows):";
            // 
            // numericUpDownVerticalSpacing
            // 
            this.numericUpDownVerticalSpacing.Location = new System.Drawing.Point(168, 12);
            this.numericUpDownVerticalSpacing.Name = "numericUpDownVerticalSpacing";
            this.numericUpDownVerticalSpacing.Size = new System.Drawing.Size(60, 19);
            this.numericUpDownVerticalSpacing.TabIndex = 1;
            this.numericUpDownVerticalSpacing.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // FormMain
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(240, 43);
            this.Controls.Add(this.numericUpDownVerticalSpacing);
            this.Controls.Add(this.labelVerticalSpacing);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.Text = "ImageTableForExcel";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.dragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.dragEnter);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownVerticalSpacing)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelVerticalSpacing;
        private System.Windows.Forms.NumericUpDown numericUpDownVerticalSpacing;
    }
}

