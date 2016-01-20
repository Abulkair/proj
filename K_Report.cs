#region PDFsharp - A .NET library for processing PDF
//
// Copyright (c) 2005-2009 empira Software GmbH, Cologne (Germany)
//
// http://www.pdfsharp.com
//
// http://sourceforge.net/projects/pdfsharp
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT OF THIRD PARTY RIGHTS.
// IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
// DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
// USE OR OTHER DEALINGS IN THE SOFTWARE.
#endregion

using System;
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Data;
using System.Collections.Generic;
using System.Windows.Media.Imaging;
using RDotNet;
using System.Runtime.InteropServices;

namespace K_Treasury
{
    /// <summary>
    /// The base interface for all report's classes.
    /// </summary>
    /// 
    interface IReport
    {
        void SaveReport(string path);
        void OpenReport(string path);
        //void Close();
    }

    public class Word_Report : IReport
    {
        private Word.Application app;
        private Word.Document doc;

        public Word_Report(string templatePath = "")
        {
            app = new Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            
            try
            {
                doc = app.Documents.Add(Template: templatePath);
            }
            catch (Exception exc)
            {
                Close();
                throw;
            }
        }

        public void FillBookmark(string BookMark, string Value)
        {
            if (doc.Bookmarks.Exists(BookMark))
                doc.Bookmarks[BookMark].Range.Text = Value; 
        }

        public void FillBookmark(string BookMark, int ipos, string chartPath)
        {
            //if (doc.Bookmarks.Exists(BookMark))
            //    doc.Bookmarks[BookMark].Range.Paste();
            var lastPar = doc.Paragraphs.Last.Range;
            doc.Paragraphs.Add();
            var inlineShape = lastPar.InlineShapes.AddPicture(chartPath);
            doc.Paragraphs.Add();
            //app.Selection.TypeParagraph();
            //app.Selection.TypeParagraph();
            //app.Selection.Paste();

            //object rStart = doc.Content.End-1;
            //object rEnd = doc.Content.End;
            //var frame = doc.Frames.Add(doc.Range(ref rStart, ref rEnd));
            
            //frame.VerticalPosition = 30 + (ipos%2)*600;
            //frame.HorizontalPosition = 50;
            
            //frame.Height = 500;
            //frame.Width = 1000;
            //frame.Select();
            //frame.Range.Paste();


        }

        public void ShowReport(string path)
        {
            try
            {
                app.Visible = true;

                app.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't show word report.", messageBoxText: ex.ToString());
            }
        }

        public void OpenReport(string path)
        {
            try
            {
                Process.Start(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't open report.", messageBoxText: ex.Message);
            }
        }

        public void SaveReport(string path)
        {
            try
            {
                doc.SaveAs2(FileName: path, FileFormat: Word.WdSaveFormat.wdFormatPDF);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't save word report.", messageBoxText: ex.ToString());
            }
        }

        private void Close()
        {
            if(doc != null)
                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            
            if(app != null)
                app.Quit();

            doc = null;
            app = null;
        }

    }

    /// <summary>
    /// The Pdf report class
    /// </summary>
    public class Pdf_Report : IReport
    {
        public XColor backColor;
        public XColor backColor2;
        public XColor shadowColor;
        public double borderWidth;
        public XPen borderPen;
        public PdfDocument document;
        XGraphicsState state;


        public Pdf_Report()
        {
            backColor = XColors.Ivory;
            backColor2 = XColors.WhiteSmoke;

            backColor = XColor.FromArgb(212, 224, 240);
            backColor2 = XColor.FromArgb(253, 254, 254);

            shadowColor = XColors.Gainsboro;
            borderWidth = 4.5;
            borderPen = new XPen(XColor.FromArgb(94, 118, 151), this.borderWidth);//new XPen(XColors.SteelBlue, this.borderWidth);
            document = new PdfDocument();
            document.Info.Title = "Created with PDFsharp";
        }

        /// <summary>
        /// Save created document
        /// </summary>
        public void SaveReport(string path)
        {
            try
            {
                document.Save(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't create report.", messageBoxText: ex.Message);
            }
        }
        /// <summary>
        /// Open created document
        /// </summary>
        public void OpenReport(string path)
        {
            try
            {
                Process.Start(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't open report.", messageBoxText: ex.Message);
            }
        }

        /// <summary>
        /// Draws the page title and footer.
        /// </summary>
        public void DrawTitle(PdfPage page, XGraphics gfx, string title)
        {
            XRect rect = new XRect(new XPoint(), gfx.PageSize);
            rect.Inflate(-10, -15);
            XFont font = new XFont("Verdana", 14, XFontStyle.Bold);
            gfx.DrawString(title, font, XBrushes.MidnightBlue, rect, XStringFormats.TopCenter);

            rect.Offset(0, 5);
            font = new XFont("Verdana", 8, XFontStyle.Italic);
            XStringFormat format = new XStringFormat();
            format.Alignment = XStringAlignment.Near;
            format.LineAlignment = XLineAlignment.Far;
            gfx.DrawString("Created with " + PdfSharp.ProductVersionInfo.Producer, font, XBrushes.DarkOrchid, rect, format);

            font = new XFont("Verdana", 8);
            format.Alignment = XStringAlignment.Center;
            gfx.DrawString(document.PageCount.ToString(), font, XBrushes.DarkOrchid, rect, format);

            document.Outlines.Add(title, page, true);
        }

        /// <summary>
        /// Draws a sample box.
        /// </summary>
        public void BeginBox(XGraphics gfx, int number, string title)
        {
            const int dEllipse = 15;
            XRect rect = new XRect(0, 20, 300, 200);
            if (number % 2 == 0)
                rect.X = 300 - 5;
            rect.Y = 40 + ((number - 1) / 2) * (200 - 5);
            rect.Inflate(-10, -10);
            XRect rect2 = rect;
            rect2.Offset(this.borderWidth, this.borderWidth);
            gfx.DrawRoundedRectangle(new XSolidBrush(this.shadowColor), rect2, new XSize(dEllipse + 8, dEllipse + 8));
            XLinearGradientBrush brush = new XLinearGradientBrush(rect, this.backColor, this.backColor2, XLinearGradientMode.Vertical);
            gfx.DrawRoundedRectangle(this.borderPen, brush, rect, new XSize(dEllipse, dEllipse));
            rect.Inflate(-5, -5);

            XFont font = new XFont("Verdana", 12, XFontStyle.Regular);
            gfx.DrawString(title, font, XBrushes.Navy, rect, XStringFormats.TopCenter);

            rect.Inflate(-10, -5);
            rect.Y += 20;
            rect.Height -= 20;
            //gfx.DrawRectangle(XPens.Red, rect);

            this.state = gfx.Save();
            gfx.TranslateTransform(rect.X, rect.Y);
        }

        public void EndBox(XGraphics gfx)
        {
            gfx.Restore(this.state);
        }

        public void addTableRowPdf(XGraphics gfx, string header, string value, int leftpos, int heighpos)
        {
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
            XFont font = new XFont("Verdana", 14, XFontStyle.Italic, options);

            XRect rect = new XRect(leftpos, heighpos, 130, 20);
            //gfx.DrawRectangle(XBrushes.SeaShell, rect);
            gfx.DrawString(header, font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(leftpos + 140, heighpos, 200, 20);
            //gfx.DrawRectangle(XBrushes.SeaShell, rect);
            gfx.DrawString(value, font, XBrushes.Black, rect, XStringFormats.TopLeft);
        }

        public void addTableRow(XGraphics gfx, string value, double leftpos, double heighpos, double size, XFont font)
        {
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
            XRect rect = new XRect(leftpos, heighpos, size, 20);
            gfx.DrawRectangle(XPens.Transparent, XBrushes.Transparent, rect);
            gfx.DrawString(value, font, XBrushes.Black, rect, XStringFormats.Center);
        }
    }


    /// <summary>
    /// The Excel report class
    /// </summary>
    /// 
    public class Excel_Report : IReport
    {
        //Екземпляр приложения Excel
        public Excel.Application xlApp;
        //Книга
        private Excel.Workbook xlBook;
        //Лист
        Excel.Worksheet xlSheet;
        //Выделеная область
        Excel.Range xlSheetRange;

        public Excel_Report()
        {
            try
            {
                xlApp = new Excel.Application();
                
                //добавляем книгу
                xlBook = xlApp.Workbooks.Add(Type.Missing);

                //делаем временно неактивным документ
                xlApp.Interactive = false;
                xlApp.EnableEvents = false;
                xlApp.DisplayAlerts = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't create excel report.", messageBoxText: ex.ToString());
                Close();
            }
        }
        public void SaveData(string sheetName, DataTable sheetData)
        {
            try
            {
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];

                //Название листа
                xlSheet.Name = sheetName.Substring(startIndex: 0, length: Math.Min(31, sheetName.Length) );

                string data = "";
                int lastUsedRowNumber = 1;
                //Вывод шапки отчёта
                //выделяем 1 столбец шапки

                //называем колонки
                for (int i = 0; i < sheetData.Columns.Count; i++)
                {
                    data = sheetData.Columns[i].ColumnName.ToString();
                    xlSheet.Cells[lastUsedRowNumber, i + 1] = data;

                    //выделяем первую строку
                    xlSheetRange = xlSheet.get_Range(String.Format("A{0}:Z{1}", lastUsedRowNumber, lastUsedRowNumber), Type.Missing);

                    //делаем полужирный текст и перенос слов
                    //xlSheetRange.WrapText = true;
                    xlSheetRange.Font.Bold = true;
                }

                //заполняем строки
                for (int rowInd = 0; rowInd < sheetData.Rows.Count; rowInd++)
                {
                    for (int collInd = 0; collInd < sheetData.Columns.Count; collInd++)
                    {
                        data = sheetData.Rows[rowInd].ItemArray[collInd].ToString();
                        xlSheet.Cells[rowInd + lastUsedRowNumber + 1, collInd + 1] = data;
                    }
                }

                //выбираем всю область данных
                xlSheetRange = xlSheet.UsedRange;

                //выравниваем строки и колонки по их содержимому
                xlSheetRange.Columns.AutoFit();
                xlSheetRange.Rows.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't add Sheet to excel report.", messageBoxText: ex.ToString());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="sheetData"></param>
        public void SaveRData(string sheetName, DataFrame sheetData)
        {
            try
            {
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];

                //Название листа
                xlSheet.Name = sheetName.Substring(startIndex: 0, length: Math.Min(31, sheetName.Length));

                string data = "";
                int lastUsedRowNumber = 1;
                //Вывод шапки отчёта
                //выделяем 1 столбец шапки

                //называем колонки
                for (int i = 0; i < sheetData.ColumnCount; i++)
                {
                    data = sheetData.ColumnNames[i].ToString();
                    xlSheet.Cells[lastUsedRowNumber, i + 1] = data;

                    //выделяем первую строку
                    xlSheetRange = xlSheet.get_Range(String.Format("A{0}:Z{1}", lastUsedRowNumber, lastUsedRowNumber), Type.Missing);

                    //делаем полужирный текст и перенос слов
                    //xlSheetRange.WrapText = true;
                    xlSheetRange.Font.Bold = true;
                }

                //заполняем строки
                for (int rowInd = 0; rowInd < sheetData.RowCount; rowInd++)
                {
                    for (int collInd = 0; collInd < sheetData.ColumnCount; collInd++)
                    {
                        data = sheetData[rowInd, collInd] == null ? "" : sheetData[rowInd, collInd].ToString();
                        xlSheet.Cells[rowInd + lastUsedRowNumber + 1, collInd + 1] = data;
                    }
                }

                //выбираем всю область данных
                xlSheetRange = xlSheet.UsedRange;

                //выравниваем строки и колонки по их содержимому
                xlSheetRange.Columns.AutoFit();
                xlSheetRange.Rows.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't add Sheet to excel report.", messageBoxText: ex.ToString());
                Close();
            }
        }

        public void AddSheet(string sheetName, DataTable sheetData, List<Tuple<string, string>> header)
        {
            try
            {
                //выбираем лист на котором будем работать (Лист 1)

                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];

                //Название листа
                xlSheet.Name = sheetName.Substring(startIndex: 0, length: Math.Min(31, sheetName.Length) );

                string data = "";
                int lastUsedRowNumber = 1;
                //Вывод шапки отчёта
                foreach (var row in header)
                {
                    xlSheet.Cells[lastUsedRowNumber, 1] = row.Item1;
                    xlSheet.Cells[lastUsedRowNumber, 2] = row.Item2;
                    lastUsedRowNumber++;
                }
                //выделяем 1 столбец шапки
                xlSheetRange = xlSheet.get_Range(String.Format("A1:A{0}", lastUsedRowNumber), Type.Missing);
                //делаем полужирный текст и перенос слов
                   // xlSheetRange.WrapText = true;
                xlSheetRange.Font.Bold = true;
                lastUsedRowNumber++;

                //называем колонки
                for (int i = 0; i < sheetData.Columns.Count; i++)
                {
                    data = sheetData.Columns[i].ColumnName.ToString();
                    xlSheet.Cells[lastUsedRowNumber, i + 1] = data;

                    //выделяем первую строку
                    xlSheetRange = xlSheet.get_Range(String.Format("A{0}:Z{1}", lastUsedRowNumber, lastUsedRowNumber), Type.Missing);

                    //делаем полужирный текст и перенос слов
                    //xlSheetRange.WrapText = true;
                    xlSheetRange.Font.Bold = true;
                }

                //заполняем строки
                for (int rowInd = 0; rowInd < sheetData.Rows.Count; rowInd++)
                {
                    for (int collInd = 0; collInd < sheetData.Columns.Count; collInd++)
                    {
                        data = sheetData.Rows[rowInd].ItemArray[collInd].ToString();
                        xlSheet.Cells[rowInd + lastUsedRowNumber + 1, collInd + 1] = data;
                    }
                }

                //выбираем всю область данных
                xlSheetRange = xlSheet.UsedRange;

                //выравниваем строки и колонки по их содержимому
                xlSheetRange.Columns.AutoFit();
                xlSheetRange.Rows.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't add Sheet to excel report.", messageBoxText: ex.ToString());
            }
        }

        public void SaveReport(string path)
        {
            try
            {
                xlApp.Workbooks[1].SaveAs(Filename: path);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't save excel report.", messageBoxText: ex.ToString());
                Close();
            }
        }

        public void OpenReport(string path)
        {
            try
            {
                Process.Start(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(caption: "Can't show excel report.", messageBoxText: ex.ToString());
                Close();
            }

        }

        private void Close()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (xlSheetRange != null)
                    Marshal.FinalReleaseComObject(xlSheetRange);

                if(xlSheet != null)
                    Marshal.FinalReleaseComObject(xlSheet);

                xlBook.Close(false);
                Marshal.FinalReleaseComObject(xlBook);

                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
            }
            catch (Exception)
            {
                
            }
        }
    }
}
