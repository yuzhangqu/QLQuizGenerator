﻿/* ====================================================================
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
            string filename="quiz.zip";
            //Response.ContentType = "application/vnd.ms-excel";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}",filename));
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

        HSSFWorkbook hssfworkbook;
        CellStyle indexStyle;
        CellStyle quizStyle;
        CellStyle ansStyle;
        int indexCount = 10;
        int[] mixIndex = { 3, 4, 8, 9 };
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

        void Print<T>(ref Sheet sheet, ref int rownum, ref Random rand, ref int index) where T : Lv, new()
        {
            PrintIndex(ref sheet, ref rownum, ref index);
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
            quizFont.FontName = "黑体";
            quizFont.FontHeightInPoints = 14;
            quizStyle.SetFont(quizFont);
            quizStyle.BorderLeft = CellBorderType.THIN;
            quizStyle.BorderRight = CellBorderType.THIN;

            ansStyle = hssfworkbook.CreateCellStyle();
            Font ansFont = hssfworkbook.CreateFont();
            ansFont.FontName = "宋体";
            ansFont.FontHeightInPoints = 14;
            ansStyle.SetFont(ansFont);
            ansStyle.BorderTop = CellBorderType.THIN;
            ansStyle.BorderBottom = CellBorderType.THIN;
            ansStyle.BorderLeft = CellBorderType.THIN;
            ansStyle.BorderRight = CellBorderType.THIN;

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
