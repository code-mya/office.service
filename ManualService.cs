using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using iTextSharp.text.pdf;
using System.IO;
using System.Reflection;
using iTextSharp.text.pdf.parser;
using Paragraph = Microsoft.Office.Interop.Word.Paragraph;
using System.Data;

namespace Office.Service
{
    /// <summary>
    /// 手动服务
    /// </summary>
    public class ManualService : BaseService
    {
        private readonly OfficeSettings settings;
        public ManualService(OfficeSettings options) => this.settings = options;

        #region word操作
        public override bool WordConvertExcel(string sourceFile, string targetFile)
        {
            var result = true;
            object missing = Type.Missing;
            Word.Application wordApp = null;
            Word.Document wordDoc = null;
            Excel.Application excelApp = null;
            Workbook workBook = null;
            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Open(sourceFile);
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                workBook = excelApp.Workbooks.Add(Missing.Value);
                Worksheet sheet = (Worksheet)workBook.ActiveSheet;
                int num = 1;
                foreach (Paragraph item in wordDoc.Paragraphs)
                {
                    sheet.Cells[num, 1] = item.Range.Text;
                    num++;
                }
                workBook.SaveAs(targetFile);
            }
            catch (Exception ex)
            {
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    excelApp = null;
                }
                if (wordDoc != null)
                {
                    wordDoc.Close(ref missing, ref missing, ref missing);
                    wordDoc = null;
                }
                if (wordApp != null)
                {
                    wordApp.Quit(ref missing, ref missing, ref missing);
                    wordApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        public override bool WordConvertPDF(string sourceFile, string targetFile)
        {
            bool result;
            WdExportFormat exportFormat = WdExportFormat.wdExportFormatPDF;
            object paramMissing = Type.Missing;
            Word.Application wordApplication = new Word.Application();
            Word.Document wordDocument = null;
            try
            {
                object paramSourceDocPath = sourceFile;
                string paramExportFilePath = targetFile;
                WdExportFormat paramExportFormat = exportFormat;
                bool paramOpenAfterExport = false;
                WdExportOptimizeFor paramExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;
                WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                WdExportCreateBookmarks paramCreateBookmarks = WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;
                wordDocument = wordApplication.Documents.Open(
                        ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing);

                if (wordDocument != null)
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                            paramExportFormat, paramOpenAfterExport,
                            paramExportOptimizeFor, paramExportRange, paramStartPage,
                            paramEndPage, paramExportItem, paramIncludeDocProps,
                            paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                            paramBitmapMissingFonts, paramUseISO19005_1,
                            ref paramMissing);
                result = true;
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        public override bool WordMerge(List<string> addLists, string target)
        {
            bool result = true;
            object missing = Type.Missing;
            Word.Application wordApp = null;
            Word.Document wordDoc = null;
            object type = WdBreakType.wdSectionBreakContinuous;
            object readOnly = false;
            object isVisible = false;
            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();
                foreach (string item in addLists)
                {
                    Word.Document openWord;
                    openWord = wordApp.Documents.Open(item, missing, ref readOnly, missing,
                                    missing, missing, missing, missing, missing, missing,
                                    missing, ref isVisible, missing, missing, missing, missing);
                    openWord.Select();
                    int num = openWord.Sections.Count;
                    for (int i = num; i >= 1; i--)
                    {
                        openWord.Sections[i].Range.Copy();
                        object start = 0;
                        Word.Range newRang = wordDoc.Range(ref start, ref start);
                        wordDoc.Sections[1].Range.InsertBreak(ref type);//插入换行符    
                        WdRecoveryType rType = new WdRecoveryType();
                        wordDoc.Sections[1].Range.PasteAndFormat(rType);
                    }
                }
                object format = WdSaveFormat.wdFormatDocument;
                wordDoc.SaveAs(target, ref format);
                return result;
            }
            catch (Exception ex)
            {
                result = false;
            }
            finally
            {
                if (wordDoc != null)
                {
                    wordDoc.Close(ref missing, ref missing, ref missing);
                    wordDoc = null;
                }
                if (wordApp != null)
                {
                    wordApp.Quit(ref missing, ref missing, ref missing);
                    wordApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        #endregion word操作

        #region excel操作
        public override bool ExcelConvertPDF(string sourceFile, string targetFile)
        {
            bool result;
            object missing = Type.Missing;
            Excel.Application application = null;
            Workbook workBook = null;
            try
            {
                application = new Excel.Application();
                object target = targetFile;
                object type = XlFixedFormatType.xlTypePDF;
                workBook = application.Workbooks.Open(sourceFile, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);
                workBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, target, XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        public override bool ExcelConvertWord(string sourceFile, string targetFile)
        {
            var result = true;
            object missing = Type.Missing;
            Word.Document wordDoc = null;
            Excel.Application excelApp = null;
            Word.Application wordApp = null;
            Excel.Workbook workBook = null;
            try
            {
                string[] strAry = new string[] { "A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z" };
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.UserControl = true;
                Workbook wb = excelApp.Workbooks.Open(sourceFile, missing, true, missing, missing, missing, missing, missing, missing, true, missing, missing, missing, missing, missing);
                Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
                int colsint = ws.UsedRange.Cells.Columns.Count; //得到行数
                Excel.Range range = ws.Cells.get_Range("A1", $"{strAry[colsint-1]}{rowsint}");
                string[,] strs = new string[rowsint,colsint];
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();
                object line = WdUnits.wdLine;
                wordApp.Selection.MoveDown(ref line, missing, missing);
                wordApp.Selection.TypeParagraph();//换行
                Word.Range wordRange = wordApp.Selection.Range;
                Word.Table wordTable = wordApp.Selection.Tables.Add(wordRange, rowsint, colsint, ref missing, ref missing);
                //设置表格的字体大小粗细
                wordTable.Range.Font.Size = 10;
                wordTable.Range.Font.Bold = 0;
                wordTable.Borders.Enable = 1;
                for ( int i = 1; i <= rowsint; i++ )
                {
                    for ( int j = 1; j <= colsint; j++ )
                    {
                        wordTable.Cell(i, j).Range.Text = ((object[,])range.Value2)[i,j]?.ToString();
                    }
                }
                wordDoc.SaveAs(targetFile);
            }
            catch (Exception ex)
            {
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    excelApp = null;
                }
                if (wordDoc != null)
                {
                    wordDoc.Close(ref missing, ref missing, ref missing);
                    wordDoc = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        public override bool ExcelMerge(List<string> addLists,string target)
        {
            var result = true;
            object missing = Type.Missing;
            Excel.Application excelApp = null;
            Excel.Workbook workBook = null;
            Excel.Range sourceRange;
            Excel.Range destRange;
            try
            {
                excelApp = new Excel.Application();
                workBook = excelApp.Workbooks.Open(addLists[0]);
                Worksheet sheet = (Worksheet)workBook.Sheets[1];
                for (int i = 1; i < addLists.Count; i++)
                {
                    Workbook addBook = excelApp.Workbooks.Open(addLists[i]);
                    Worksheet addSheet = (Worksheet)addBook.Sheets[1];
                    addSheet.Copy(missing, sheet);
                    addBook.Close();
                }
                workBook.SaveAs(target);
            }
            catch (Exception ex)
            {
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    excelApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        #endregion excel操作

        #region pdf操作
        public override bool PDFConvertExcel(string sourceFile, string targetFile)
        {
            var result = true;
            object missing = Type.Missing;
            Excel.Application excelApp = null;
            Workbook workBook = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                workBook = excelApp.Workbooks.Add(Missing.Value);
                Worksheet sheet = (Worksheet)workBook.ActiveSheet;
                using (PdfReader reader = new PdfReader(sourceFile))
                {
                    int num = 1;
                    // 遍历PDF页面
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        string text = PdfTextExtractor.GetTextFromPage(reader, i);
                        sheet.Cells[num, 1] = text;
                        num++;
                    }
                    workBook.SaveAs(targetFile);
                }
            }
            catch (Exception ex)
            {
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    excelApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        public override bool PDFConvertWord(string sourceFile, string targetFile)
        {
            bool result = true;
            object missing = Type.Missing;
            Word.Application wordApp = null;
            Word.Document wordDoc = null;
            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();
                using (PdfReader reader = new PdfReader(sourceFile))
                {
                    // 遍历PDF页面
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        string text = PdfTextExtractor.GetTextFromPage(reader, i);
                        wordDoc.Paragraphs.Last.Range.Text = text;
                    }
                    wordDoc.SaveAs(targetFile);
                }
                return result;
            }
            catch (Exception ex)
            {
                result = false;
            }
            finally
            {
                if (wordDoc != null)
                {
                    wordDoc.Close(ref missing, ref missing, ref missing);
                    wordDoc = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        public override bool PDFMerge(List<string> addLists, string target)
        {
            using (var stream = new FileStream(target, FileMode.Create))
            {
                using (var doc = new iTextSharp.text.Document())
                {
                    using (var pdf = new PdfCopy(doc, stream))
                    {
                        doc.Open();
                        addLists.ForEach(file =>
                        {
                            var reader = new PdfReader(file);
                            for (int i = 0; i < reader.NumberOfPages; i++)
                            {
                                var page = pdf.GetImportedPage(reader, i + 1);
                                pdf.AddPage(page);
                            }
                            pdf.FreeReader(reader);
                            reader.Close();
                        });
                    }
                }
            }
            return true;
        }
        #endregion pdf操作

        #region powerpoint操作
        public override bool PowerPointConvertExcel(string sourceFile, string targetFile)
        {
            return false;
        }

        public override bool PowerPointConvertPDF(string sourceFile, string targetFile)
        {
            bool result;
            object missing = Type.Missing;
            PowerPoint.Application application = null;
            Presentation persentation = null;
            try
            {
                application = new PowerPoint.Application();
                persentation = application.Presentations.Open(sourceFile, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetFile, PpSaveAsFileType.ppSaveAsPDF, MsoTriState.msoTrue);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        public override bool PowerPointConvertWord(string sourceFile, string targetFile)
        {
            return false;
        }

        public override bool PowerPointMerge(List<string> addLists, string target)
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}