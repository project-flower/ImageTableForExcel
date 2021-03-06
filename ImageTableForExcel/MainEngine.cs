﻿using Microsoft.Office.Interop.Excel;
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
                        releaseComObject(ref shape);
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
                releaseComObject(ref shapes);
                releaseComObject(ref startCell);
                releaseComObject(ref worksheet);
                releaseComObject(ref worksheets);

                if (workbook != null)
                {
                    workbook.Close();
                    releaseComObject(ref workbook);
                }

                if (workbooks != null)
                {
                    try
                    {
                        workbooks.Close();
                    }
                    catch
                    {
                        // "通常使うプログラムとして設定されていません。"が表示されている場合、例外がスローされる。
                    }

                    releaseComObject(ref workbooks);
                }

                GC.Collect();

                if (application != null)
                {
                    try
                    {
                        application.Quit();
                    }
                    catch
                    {
                        // "サブスクリプションの有効期限が切れています"が表示されている場合、例外がスローされる。
                    }

                    releaseComObject(ref application);
                }

                GC.Collect();
            }
        }

        private static void releaseComObject<T>(ref T o)
        {
            if (o == null)
            {
                return;
            }

            Marshal.FinalReleaseComObject(o);
            o = default;
        }
    }
}
