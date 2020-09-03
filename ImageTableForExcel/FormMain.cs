using System;
using System.Windows.Forms;

namespace ImageTableForExcel
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void generate(string directoryName)
        {
            string[] faultFiles;

            try
            {
                string fileName = MainEngine.Generate(directoryName, (int)numericUpDownVerticalSpacing.Value, out faultFiles);
                string message = (fileName + " を作成しました。");

                if ((faultFiles != null) && (faultFiles.Length > 0))
                {
                    message += "\r\n以下のファイルは読み取れませんでした。\r\n\r\n";
                    message += string.Join("\r\n", faultFiles);
                }

                showMessage(message, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                showErrorMessage(exception.Message);
            }
        }

        private void showErrorMessage(string text)
        {
            showMessage(text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private DialogResult showMessage(string text, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return MessageBox.Show(this, text, Text, buttons, icon);
        }

        private void dragDrop(object sender, DragEventArgs e)
        {
            var dropData = (e.Data.GetData(DataFormats.FileDrop)) as string[];

            if ((dropData == null) || (dropData.Length < 1))
            {
                return;
            }

            if (dropData.Length > 1)
            {
                showErrorMessage("ディレクトリは1つだけ指定してください。");
            }

            generate(dropData[0]);
        }

        private void dragEnter(object sender, DragEventArgs e)
        {
            e.Effect
                = (e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.All
                : DragDropEffects.None);
        }
    }
}
