using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImageTableForExcel
{
    public static class MainEngine
    {
        public static string Generate(string directoryName, int verticalSpacing, out string[] faultFiles)
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

            Excel.Application application = null;
            Workbooks workbooks = null;
            Workbook workbook = null;
            Sheets worksheets = null;
            Worksheet worksheet = null;
            Range startCell = null;
            Shapes shapes = null;

            try
            {
                application = new Excel.Application();
                application.DisplayAlerts = false;
                workbooks = application.Workbooks;
                workbook = workbooks.Add(Type.Missing);
                worksheets = workbook.Sheets;
                worksheet = (Worksheet)worksheets[1];

                try
                {
                    worksheet.Name = directory.Name;
                }
                catch
                {
                }

                float y = 0;
                startCell = (Range)worksheet.Cells[1, 1];
                float rowHeight = float.Parse(startCell.RowHeight.ToString());
                var faultFilesList = new List<string>();
                shapes = worksheet.Shapes;

                foreach (FileInfo fileInfo in directory.GetFiles())
                {
                    Shape shape = null;

                    try
                    {
                        using (Image image = Image.FromFile(fileInfo.FullName))
                        {
                            shape = shapes.AddPicture(
                                // FileName
                                fileInfo.FullName,
                                // LinkToFile
                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                // SaveWithDocument
                                Microsoft.Office.Core.MsoTriState.msoTrue,
                                // Left, Top, Width, Height
                                0, y, -1, -1);

                            y += (((int)(shape.Height / rowHeight + 1 + verticalSpacing)) * rowHeight);
                        }
                    }
                    catch
                    {
                        faultFilesList.Add(fileInfo.Name);
                        continue;
                    }
                    finally
                    {
                        releaseComObject(shape);
                    }
                }

                try
                {
                    string fileName = (directory.Name + ".xlsx");
                    string fileFullName = Path.Combine(directory.FullName, fileName);
                    workbook.SaveAs(fileFullName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    faultFiles = faultFilesList.ToArray();
                    return fileName;
                }
                catch
                {
                    throw;
                }
            }
            finally
            {
                releaseComObject(shapes);
                releaseComObject(startCell);
                releaseComObject(worksheet);
                releaseComObject(worksheets);

                if (workbook != null)
                {
                    workbook.Close();
                    releaseComObject(workbook);
                }

                if (workbooks != null)
                {
                    workbooks.Close();
                    releaseComObject(workbooks);
                }

                GC.Collect();

                if (application != null)
                {
                    application.Quit();
                    releaseComObject(application);
                }

                GC.Collect();
            }
        }

        private static void releaseComObject(object o)
        {
            if (o != null)
            {
                while (Marshal.ReleaseComObject(o) >= 0)
                {
                    ;
                }

                o = null;
            }
        }
    }
}
