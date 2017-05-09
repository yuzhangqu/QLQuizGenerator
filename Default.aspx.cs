/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * NPOI HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Collections.Generic;

using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExportXlsToDownload
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            //Response.ContentType = "application/vnd.ms-excel";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData2();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        protected void Button3_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData3();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        protected void Button4_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData4();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        protected void Button5_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData5();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        protected void Button6_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData6();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        protected void Button7_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData7();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        HSSFWorkbook hssfworkbook;
        CellStyle indexStyle;
        CellStyle quizStyle;
        CellStyle ansStyle;
        CellStyle headerStyle;
        CellStyle titleStyle;
        CellStyle headtextStyle;
        int indexCount = 10;
        List<int> mixIndex = new List<int> { 3, 4, 8, 9 };
        Random rnd = new Random();

        MemoryStream WriteToStream()
        {
            using (var compressedFileStream = new MemoryStream())
            {
                //Create an archive and store the stream in memory.
                using (var zipArchive = new ZipArchive(compressedFileStream, ZipArchiveMode.Update, false))
                {
                    var zipEntry = zipArchive.CreateEntry("教师版.xls");

                    //Get the stream of the attachment
                    using (var file = new MemoryStream())
                    {
                        //Write the stream data of workbook to the root directory
                        hssfworkbook.Write(file);
                        using (var zipEntryStream = zipEntry.Open())
                        {
                            //Copy the attachment stream to the zip entry stream
                            file.Position = 0;
                            file.CopyTo(zipEntryStream);
                        }
                    }

                    zipEntry = zipArchive.CreateEntry("学生版.xls");

                    //Get the stream of the attachment
                    using (var file = new MemoryStream())
                    {
                        //Write the stream data of workbook to the root directory
                        ansStyle.GetFont(hssfworkbook).Color = HSSFColor.WHITE.index;
                        hssfworkbook.Write(file);
                        using (var zipEntryStream = zipEntry.Open())
                        {
                            //Copy the attachment stream to the zip entry stream
                            file.Position = 0;
                            file.CopyTo(zipEntryStream);
                        }
                    }
                }

                return compressedFileStream;
            }
        }

        void GenerateData()
        {
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                int index = 1;
                int rownum = 0;
                Sheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 11;
                sheet.DefaultRowHeightInPoints = 15;
                sheet.PrintSetup.PaperSize = (short)PaperSize.B4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                Cell c = sheet.CreateRow(rownum++).CreateCell(0);
                c.CellStyle = headtextStyle;
                c.SetCellValue("庆龄幼儿园珠心算训练题——第三套");
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 9));

                Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv7>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv7>(ref sheet, ref rownum, ref rand, ref index);
            }
        }

        void GenerateData2()
        {
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                int index = 1;
                int rownum = 0;
                Sheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 11;
                sheet.DefaultRowHeightInPoints = 15;
                sheet.PrintSetup.PaperSize = (short)PaperSize.B4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
            }
        }

        void GenerateData3()
        {
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                int index = 1;
                int rownum = 0;
                Sheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 11;
                sheet.DefaultRowHeightInPoints = 20;
                sheet.PrintSetup.PaperSize = (short)PaperSize.B4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                PrintHeader(ref sheet, ref rownum);
                Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
            }
        }

        void PrintHeader(ref Sheet sheet, ref int rownum)
        {
            Cell c = sheet.CreateRow(rownum++).CreateCell(0);
            c.CellStyle = titleStyle;
            c.SetCellValue("武汉市珠心算选拔赛模拟试题");
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 9));
            rownum++;

            Row row = sheet.CreateRow(rownum++);
            SetCellString(ref row, 0, headerStyle, "单位");
            SetCellString(ref row, 1, headerStyle, " ");
            SetCellString(ref row, 2, headerStyle, "姓名");
            SetCellString(ref row, 3, headerStyle, " ");
            SetCellString(ref row, 4, headerStyle, "考号");
            SetCellString(ref row, 5, headerStyle, " ");
            SetCellString(ref row, 9, headerStyle, "限时5分钟", false);
            rownum++;

            row = sheet.CreateRow(rownum++);
            SetCellString(ref row, 3, headerStyle, " ");
            SetCellString(ref row, 4, headerStyle, "计算题");
            SetCellString(ref row, 5, headerStyle, "错题");
            SetCellString(ref row, 6, headerStyle, "对题");
            SetCellString(ref row, 7, headerStyle, "初审");
            SetCellString(ref row, 8, headerStyle, "复核");
            SetCellString(ref row, 9, headerStyle, "成绩");

            row = sheet.CreateRow(rownum++);
            SetCellString(ref row, 3, headerStyle, "本页");
            SetCellString(ref row, 4, headerStyle, " ");
            SetCellString(ref row, 5, headerStyle, " ");
            SetCellString(ref row, 6, headerStyle, " ");
            SetCellString(ref row, 7, headerStyle, " ");
            SetCellString(ref row, 8, headerStyle, " ");
            SetCellString(ref row, 9, headerStyle, " ");

            row = sheet.CreateRow(rownum++);
            SetCellString(ref row, 3, headerStyle, "合计");
            SetCellString(ref row, 4, headerStyle, " ");
            SetCellString(ref row, 5, headerStyle, " ");
            SetCellString(ref row, 6, headerStyle, " ");
            SetCellString(ref row, 7, headerStyle, " ");
            SetCellString(ref row, 8, headerStyle, " ");
            SetCellString(ref row, 9, headerStyle, " ");
            rownum++;
        }

        void SetCellString(ref Row row, int i, CellStyle style, string s, bool b = true)
        {
            Cell c = row.CreateCell(i);
            if (b)
            {
                c.CellStyle = style;
            }
            c.SetCellValue(s);
        }

        void GenerateData4()
        {
            indexCount = 6;
            mixIndex.Clear();
            mixIndex.Add(2);
            mixIndex.Add(5);
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                int index = 1;
                int rownum = 0;
                Sheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 18;
                sheet.DefaultRowHeightInPoints = 20;
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                Cell c = sheet.CreateRow(rownum++).CreateCell(0);
                c.CellStyle = headtextStyle;
                c.SetCellValue("3--5位数 10笔");
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 5);
                sheet.AddMergedRegion(region);
                ((HSSFSheet)sheet).SetBorderBottomOfRegion(region, CellBorderType.THIN, HSSFColor.BLACK.index);

                Print<C_3_5_10>(ref sheet, ref rownum, ref rand, ref index, false);
                Print<C_3_5_10>(ref sheet, ref rownum, ref rand, ref index, false);
                Print<C_3_5_10>(ref sheet, ref rownum, ref rand, ref index, false);
                Print<C_3_5_10>(ref sheet, ref rownum, ref rand, ref index, false);
            }

            indexCount = 10;
            mixIndex.Clear();
            mixIndex.Add(3);
            mixIndex.Add(4);
            mixIndex.Add(8);
            mixIndex.Add(9);
        }

        void GenerateData5()
        {
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                int index = 1;
                int rownum = 0;
                Sheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 10;
                sheet.DefaultRowHeightInPoints = 20;
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
            }
        }

        void GenerateData6()
        {
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                int index = 1;
                int rownum = 0;
                Sheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 10;
                sheet.DefaultRowHeightInPoints = 20;
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
            }
        }

        void GenerateData7()
        {
            Random rand = new Random();
            int index = 1;
            int rownum = 0;
            Sheet sheet = hssfworkbook.CreateSheet("书面赛");
            sheet.DefaultColumnWidth = 11;
            sheet.DefaultRowHeightInPoints = 15;
            sheet.PrintSetup.PaperSize = (short)PaperSize.B4;
            sheet.PrintSetup.Landscape = false;

            sheet.CreateRow(rownum++).CreateCell(9);
            rownum++;

            Cell c = sheet.CreateRow(rownum++).CreateCell(0);
            c.CellStyle = headtextStyle;
            c.SetCellValue("2017年武汉市珠心算选拔赛 - 书面赛");
            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, 9));

            Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
            Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
            Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
            Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
            Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
        }

        void Print<T>(ref Sheet sheet, ref int rownum, ref Random rand, ref int index, bool hasIndex = true) where T : Lv, new()
        {
            if (hasIndex)
            {
                PrintIndex(ref sheet, ref rownum, ref index);
            }

            Lv[] lvs = new Lv[indexCount];
            for (int i = 0; i < indexCount; ++i)
            {
                lvs[i] = new T();
                lvs[i].Init(ref rand, mixIndex.Contains(i));
            }

            int count = lvs[0].Count;

            for (int i = 0; i <= count; ++i)
            {
                sheet.CreateRow(rownum + i);
            }

            for (int j = 0; j < indexCount; ++j)
            {
                for (int i = 0; i < count; ++i)
                {
                    Cell c = sheet.GetRow(rownum + i).CreateCell(j, CellType.NUMERIC);
                    c.CellStyle = quizStyle;
                    c.SetCellValue(lvs[j].Nums[i]);
                }

                string colName = CellReference.ConvertNumToColString(j);
                Cell ans = sheet.GetRow(rownum + count).CreateCell(j, CellType.FORMULA);
                ans.CellFormula = "SUM(" + colName + (rownum + 1) + ":" + colName + (rownum + count) + ")";
                ans.CellStyle = ansStyle;

            }

            rownum += count + 1;
        }

        void PrintIndex(ref Sheet sheet, ref int rownum, ref int index)
        {
            Row row = sheet.CreateRow(rownum++);
            for (int i = 0; i < indexCount; ++i)
            {
                Cell c = row.CreateCell(i);
                c.CellStyle = indexStyle;
                c.SetCellValue(index++);
            }
        }

        void InitializeWorkbook()
        {
            hssfworkbook = new HSSFWorkbook();

            indexStyle = hssfworkbook.CreateCellStyle();
            Font indexFont = hssfworkbook.CreateFont();
            indexFont.FontName = "宋体";
            indexFont.FontHeightInPoints = 12;
            indexFont.IsItalic = true;
            indexFont.Boldweight = (short)FontBoldWeight.BOLD;
            indexStyle.SetFont(indexFont);
            indexStyle.Alignment = HorizontalAlignment.CENTER;
            indexStyle.BorderTop = CellBorderType.THIN;
            indexStyle.BorderBottom = CellBorderType.THIN;
            indexStyle.BorderLeft = CellBorderType.THIN;
            indexStyle.BorderRight = CellBorderType.THIN;

            quizStyle = hssfworkbook.CreateCellStyle();
            Font quizFont = hssfworkbook.CreateFont();
            quizFont.FontName = "Lucida Console";
            quizFont.FontHeightInPoints = 13;
            quizStyle.SetFont(quizFont);
            quizStyle.BorderLeft = CellBorderType.THIN;
            quizStyle.BorderRight = CellBorderType.THIN;
            quizStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");

            ansStyle = hssfworkbook.CreateCellStyle();
            Font ansFont = hssfworkbook.CreateFont();
            ansFont.FontName = "Lucida Console";
            ansFont.FontHeightInPoints = 14;
            ansStyle.SetFont(ansFont);
            ansStyle.BorderTop = CellBorderType.THIN;
            ansStyle.BorderBottom = CellBorderType.THIN;
            ansStyle.BorderLeft = CellBorderType.THIN;
            ansStyle.BorderRight = CellBorderType.THIN;
            ansStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");

            headerStyle = hssfworkbook.CreateCellStyle();
            Font headerFont = hssfworkbook.CreateFont();
            headerFont.FontName = "宋体";
            headerFont.FontHeightInPoints = 16;
            headerStyle.SetFont(headerFont);
            headerStyle.BorderTop = CellBorderType.THIN;
            headerStyle.BorderBottom = CellBorderType.THIN;
            headerStyle.BorderLeft = CellBorderType.THIN;
            headerStyle.BorderRight = CellBorderType.THIN;

            titleStyle = hssfworkbook.CreateCellStyle();
            Font titleFont = hssfworkbook.CreateFont();
            titleFont.FontName = "宋体";
            titleFont.FontHeightInPoints = 36;
            titleFont.Boldweight = (short)FontBoldWeight.BOLD;
            titleStyle.SetFont(titleFont);
            titleStyle.Alignment = HorizontalAlignment.CENTER;

            headtextStyle = hssfworkbook.CreateCellStyle();
            Font headtextFont = hssfworkbook.CreateFont();
            headtextFont.FontName = "宋体";
            headtextFont.FontHeightInPoints = 28;
            headtextFont.Boldweight = (short)FontBoldWeight.BOLD;
            headtextStyle.SetFont(headtextFont);
            headerStyle.BorderBottom = CellBorderType.THIN;
            headtextStyle.Alignment = HorizontalAlignment.CENTER;

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "武汉市庆龄幼儿园";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Author = "余章曲";
            si.Subject = "珠心算试题";
            hssfworkbook.SummaryInformation = si;
        }
    }
}
