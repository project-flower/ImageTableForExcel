using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace ImageTableForExcel
{
    public static class MainEngine
    {
        public static void Generate(string directoryName, int rowHeight, out string[] faultFiles)
        {
            DirectoryInfo directory;
            faultFiles = null;

            try
            {
                directory = new DirectoryInfo(directoryName);
            }
            catch
            {
                throw;
            }

            var builder = new StringBuilder();
            var faultFilesList = new List<string>();
            builder.AppendLine("<!DOCTYPE html>");
            builder.AppendLine("<html>");
            builder.AppendLine(" <head>");
            builder.AppendLine(" </head>");
            builder.AppendLine(" <body>");
            builder.AppendLine("  <table><tbody>");

            foreach (FileInfo fileInfo in directory.GetFiles())
            {
                try
                {
                    using (Image image = Image.FromFile(fileInfo.FullName))
                    {
                        builder.AppendFormat("  <tr><td><img src=\"{0}\"></td></tr>", fileInfo.Name);
                        builder.AppendLine();

                        for (int i = 0; i <= (image.Height / rowHeight); ++i)
                        {
                            builder.AppendFormat("  <tr><td></td></tr>", fileInfo.Name);
                        }
                    }
                }
                catch
                {
                    faultFilesList.Add(fileInfo.Name);
                    continue;
                }
            }

            builder.AppendLine("  </tbody></table>");
            builder.AppendLine(" </body>");
            builder.AppendLine("</html>");

            try
            {
                File.WriteAllText(Path.Combine(directory.FullName, "index.html"), builder.ToString(), Encoding.UTF8);
                faultFiles = faultFilesList.ToArray();
            }
            catch
            {
                throw;
            }
        }
    }
}
