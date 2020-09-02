using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ImageTableForExcel
{
    public static class MainEngine
    {
        public static string Generate(string directoryName, int rowHeight, out string[] faultFiles)
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

            var faultFilesList = new List<string>();

            using (var workBook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workBook.Worksheets.Add(directory.Name);
                int rowIndex = 1;

                foreach (FileInfo fileInfo in directory.GetFiles())
                {
                    try
                    {
                        using (Bitmap image = Image.FromFile(fileInfo.FullName) as Bitmap)
                        {
                            IXLCell cell = worksheet.Cell(rowIndex, 1);
                            // TODO : 行高さを取得する
                            int height = rowHeight;
                            IXLPicture picture = worksheet.AddPicture(image);
                            picture.MoveTo(cell);
                            rowIndex += (image.Height / height + 2);
                        }
                    }
                    catch
                    {
                        faultFilesList.Add(fileInfo.Name);
                        continue;
                    }
                }

                try
                {
                    string fileName = (directory.Name + ".xlsx");
                    workBook.SaveAs(Path.Combine(directory.FullName, fileName));
                    faultFiles = faultFilesList.ToArray();
                    return fileName;
                }
                catch
                {
                    throw;
                }
            }
        }
    }
}
