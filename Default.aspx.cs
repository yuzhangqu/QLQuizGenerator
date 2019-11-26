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

        protected void Button8_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData8();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        protected void Button9_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData9();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        protected void Button10_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData10();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        protected void Button11_Click(object sender, EventArgs e)
        {
            string filename = "quiz.zip";
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();

            InitializeWorkbook();
            GenerateData11();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();
        }

        HSSFWorkbook hssfworkbook;
        int indexCount = 10;
        List<int> mixIndex = new List<int> { 3, 4, 8, 9 };

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
                        ansStyle.GetFont(hssfworkbook).Color = HSSFColor.White.Index;
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
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 11;
                sheet.DefaultRowHeightInPoints = 15;
                sheet.PrintSetup.PaperSize = (short)PaperSize.B4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                ICell c = sheet.CreateRow(rownum++).CreateCell(0);
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
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
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
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
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

        void PrintHeader(ref ISheet sheet, ref int rownum)
        {
            ICell c = sheet.CreateRow(rownum++).CreateCell(0);
            c.CellStyle = titleStyle;
            c.SetCellValue("武汉市珠心算选拔赛模拟试题");
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 9));
            rownum++;

            IRow row = sheet.CreateRow(rownum++);
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

        void SetCellString(ref IRow row, int i, ICellStyle style, string s, bool b = true)
        {
            ICell c = row.CreateCell(i);
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
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 18;
                sheet.DefaultRowHeightInPoints = 20;
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                ICell c = sheet.CreateRow(rownum++).CreateCell(0);
                c.CellStyle = headtextStyle;
                c.SetCellValue("3--5位数 10笔");
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 5);
                sheet.AddMergedRegion(region);
                ((HSSFSheet)sheet).SetBorderBottomOfRegion(region, BorderStyle.Thin, HSSFColor.Black.Index);

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
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
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
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 10;
                sheet.DefaultRowHeightInPoints = 20;
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                ICell c = sheet.CreateRow(rownum++).CreateCell(0);
                c.CellStyle = headtextStyle;
                c.SetCellValue("武汉市庆龄幼儿园珠心算十级训练题");
                sheet.AddMergedRegion(new CellRangeAddress(rownum - 1, rownum - 1, 0, 9));

                Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
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
            ISheet sheet = hssfworkbook.CreateSheet("看心算");
            sheet.DefaultColumnWidth = 11;
            sheet.DefaultRowHeightInPoints = 15;
            sheet.PrintSetup.PaperSize = (short)PaperSize.B4;
            sheet.PrintSetup.Landscape = false;

            sheet.CreateRow(rownum++).CreateCell(9);
            rownum++;

            ICell c = sheet.CreateRow(rownum++).CreateCell(0);
            c.CellStyle = headtextStyle;
            c.SetCellValue("2018年武汉市珠心算比赛 - 看心算");
            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, 9));

            Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
            Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
            Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
            Print<Lv7>(ref sheet, ref rownum, ref rand, ref index);
        }

        void GenerateData8()
        {
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 14;
                sheet.DefaultRowHeightInPoints = 20;
                sheet.PrintSetup.PaperSize = 9;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;
                sheet.SetMargin(MarginType.LeftMargin, 0);
                sheet.SetMargin(MarginType.RightMargin, 0);

                int rownum = 0;
                ICell c = sheet.CreateRow(rownum++).CreateCell(0);
                c.Row.HeightInPoints = 35;
                c.CellStyle = headtextStyle;
                c.SetCellValue("小学组乘法训练题");
                _ = sheet.AddMergedRegion(new CellRangeAddress(rownum - 1, rownum - 1, 0, 17));

                for (int j = 0; j < 30; ++j)
                {
                    sheet.CreateRow(rownum + j).HeightInPoints = 27;
                }
                    
                int index = 1;
                // 第一列
                for (int j = 0; j < 15; ++j)
                {
                    PrintMulti(ref sheet, ref rand, rownum + j, 0, index++, 1, 2);
                }
                for (int j = 15; j < 30; ++j)
                {
                    PrintMulti(ref sheet, ref rand, rownum + j, 0, index++, 1, 3);
                }

                // 第二列
                for (int j = 0; j < 30; ++j)
                {
                    PrintMulti(ref sheet, ref rand, rownum + j, 6, index++, 2, 2);
                }

                // 第三列
                for (int j = 0; j < 20; ++j)
                {
                    PrintMulti(ref sheet, ref rand, rownum + j, 6 * 2, index++, 2, 3, true);
                }
                for (int j = 20; j < 30; ++j)
                {
                    PrintMulti(ref sheet, ref rand, rownum + j, 6 * 2, index++, 3, 3, true);
                }

                sheet.SetColumnWidth(1, (int)((7.29 + 0.72) * 256));
                sheet.SetColumnWidth(3, (int)((7.29 + 0.72) * 256));
                sheet.SetColumnWidth(5, (int)((8.43 + 0.72) * 256));
                sheet.SetColumnWidth(7, (int)((5.14 + 0.72) * 256));
                sheet.SetColumnWidth(9, (int)((5.14 + 0.72) * 256));
                sheet.SetColumnWidth(11, (int)((8.43 + 0.72) * 256));
                sheet.SetColumnWidth(13, (int)((9.43 + 0.72) * 256));
                sheet.SetColumnWidth(15, (int)((9.43 + 0.72) * 256));
                sheet.SetColumnWidth(17, (int)((11.57 + 0.72) * 256));
                for (int i = 0; i < 18; ++i)
                {
                    if (i % 6 == 0)
                    {
                        sheet.SetColumnWidth(i, (int)((7.14 + 0.72) * 256));
                    }
                    else if (i % 6 == 2)
                    {
                        sheet.SetColumnWidth(i, (int)((2.86 + 0.72) * 256));
                    }
                    else if (i % 6 == 4)
                    {
                        sheet.SetColumnWidth(i, (int)((2.29 + 0.72) * 256));
                    }
                    //else
                    //{
                    //    sheet.AutoSizeColumn(i);
                    //}
                }
            }
        }

        void GenerateData9()
        {
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                int index = 1;
                int rownum = 0;
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 11;
                sheet.DefaultRowHeightInPoints = 20;
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                PrintHeader(ref sheet, ref rownum);
                Print<Lv10>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv9>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv8>(ref sheet, ref rownum, ref rand, ref index);
                Print<Lv7>(ref sheet, ref rownum, ref rand, ref index);
            }
        }

        void GenerateData10()
        {
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                int index = 1;
                int rownum = 0;
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 13;
                sheet.DefaultRowHeightInPoints = 15;
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                indexCount = 6;

                ICell c = sheet.CreateRow(rownum++).CreateCell(0);
                c.CellStyle = headtextStyle;
                c.SetCellValue("庆龄幼儿园满5加破5减训练题");
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, indexCount - 1));

                indexStyle.GetFont(hssfworkbook).FontHeightInPoints = 14;
                quizStyle.GetFont(hssfworkbook).FontHeightInPoints = 16;
                ansStyle.GetFont(hssfworkbook).FontHeightInPoints = 16;

                Print<Break5>(ref sheet, ref rownum, ref rand, ref index);
                Print<Break5>(ref sheet, ref rownum, ref rand, ref index);
                Print<Break5>(ref sheet, ref rownum, ref rand, ref index);
                Print<Break5>(ref sheet, ref rownum, ref rand, ref index);
                Print<Break5>(ref sheet, ref rownum, ref rand, ref index);
            }
        }

        void GenerateData11()
        {
            indexCount = 6;
            mixIndex.Clear();
            mixIndex.Add(4);
            mixIndex.Add(5);
            Random rand = new Random();
            for (int sheetNum = 1; sheetNum <= 20; ++sheetNum)
            {
                int index = 1;
                int rownum = 0;
                ISheet sheet = hssfworkbook.CreateSheet(sheetNum.ToString());
                sheet.DefaultColumnWidth = 18;
                sheet.DefaultRowHeightInPoints = 20;
                sheet.PrintSetup.PaperSize = 9;
                sheet.PrintSetup.Landscape = false;
                sheet.Footer.Center = sheet.SheetName;

                ICell c = sheet.CreateRow(rownum++).CreateCell(0);
                c.Row.HeightInPoints = 40;
                c.CellStyle = headtextStyle;
                c.SetCellValue("小学组加减算训练");
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, indexCount - 1);
                sheet.AddMergedRegion(region);
                ((HSSFSheet)sheet).SetBorderBottomOfRegion(region, BorderStyle.Thin, HSSFColor.Black.Index);

                indexStyle.GetFont(hssfworkbook).FontHeightInPoints = 14;
                quizStyle.GetFont(hssfworkbook).FontHeightInPoints = 16;
                ansStyle.GetFont(hssfworkbook).FontHeightInPoints = 16;

                Print<C_3_5_10>(ref sheet, ref rownum, ref rand, ref index);
                Print<C_3_5_10>(ref sheet, ref rownum, ref rand, ref index);
                Print<XX7>(ref sheet, ref rownum, ref rand, ref index);
                Print<XX7>(ref sheet, ref rownum, ref rand, ref index);
            }

            indexCount = 10;
            mixIndex.Clear();
            mixIndex.Add(3);
            mixIndex.Add(4);
            mixIndex.Add(8);
            mixIndex.Add(9);
        }

        void Print<T>(ref ISheet sheet, ref int rownum, ref Random rand, ref int index, bool hasIndex = true) where T : Lv, new()
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
                    ICell c = sheet.GetRow(rownum + i).CreateCell(j, CellType.Numeric);
                    if (lvs[j].Scale < 1.0)
                    {
                        c.CellStyle = quizPointStyle;
                        c.SetCellValue(lvs[j].Nums[i] * lvs[j].Scale);
                    }
                    else
                    {
                        c.CellStyle = quizStyle;
                        c.SetCellValue(lvs[j].Nums[i] * lvs[j].Scale);
                    }
                }

                string colName = CellReference.ConvertNumToColString(j);
                ICell ans = sheet.GetRow(rownum + count).CreateCell(j, CellType.Formula);
                ans.CellFormula = "SUM(" + colName + (rownum + 1) + ":" + colName + (rownum + count) + ")";
                if (lvs[j].Scale < 1.0)
                {
                    ans.CellStyle = ansPointStyle;
                }
                else
                {
                    ans.CellStyle = ansStyle;
                }

            }

            rownum += count + 1;
        }

        void PrintIndex(ref ISheet sheet, ref int rownum, ref int index)
        {
            IRow row = sheet.CreateRow(rownum++);
            for (int i = 0; i < indexCount; ++i)
            {
                ICell c = row.CreateCell(i);
                c.CellStyle = indexStyle;
                //c.SetCellValue(index++);
                c.SetCellValue(NumToChinese(index++));
            }
        }

        string NumToChinese(int num)  // num < 100
        {
            string[] hanzi = { "", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            if (num >= 100)
            {
                return "N/A";
            }
            else if (num < 10)
            {
                return hanzi[num];
            }
            else if (num < 20)
            {
                return "十" + NumToChinese(num % 10);
            }
            return NumToChinese(num / 10) + "十" + NumToChinese(num % 10);
        }

        int RandDigit(ref Random rand, int digit)
        {
            int result = 0;
            switch (digit)
            {
                case 1:
                    result = rand.Next(2, 10);
                    break;
                case 2:
                    do
                    {
                        result = rand.Next(11, 100);
                    } while (result % 10 == 0);
                    break;
                case 3:
                    do
                    {
                        result = rand.Next(101, 1000);
                    } while (result % 10 == 0);
                    break;
                case 4:
                    do
                    {
                        result = rand.Next(1001, 10000);
                    } while (result % 10 == 0);
                    break;
                default:
                    break;
            }

            return result;
        }

        void PrintMulti(ref ISheet sheet, ref Random rand, int row, int col, int idx, int left, int right, bool point = false)
        {
            int baseCol = col;
            ICell c = sheet.GetRow(row).CreateCell(col++, CellType.String);
            c.CellStyle = indexStyle;
            c.SetCellValue(NumToChinese(idx));
            
            if (rand.Next(0, 10) < 5)
            {
                int temp = left;
                left = right;
                right = temp;
            }

            int leftPoint = 0;
            int rightPoint = 0;
            if (point)
            {
                leftPoint = rand.Next(1, left);
                rightPoint = rand.Next(1, right);
            }

            c = sheet.GetRow(row).CreateCell(col++, CellType.String);
            if (point)
            {
                c.CellStyle = multPointStyle;
                c.SetCellValue(RandDigit(ref rand, left) / Math.Pow(10, leftPoint));
            }
            else
            {
                c.CellStyle = multStyle;
                c.SetCellValue(RandDigit(ref rand, left));
            }
            

            c = sheet.GetRow(row).CreateCell(col++, CellType.String);
            c.CellStyle = multStyle;
            c.SetCellValue("*");

            c = sheet.GetRow(row).CreateCell(col++, CellType.String);
            if (point)
            {
                c.CellStyle = multPointStyle;
                c.SetCellValue(RandDigit(ref rand, right) / Math.Pow(10, rightPoint));
            }
            else
            {
                c.CellStyle = multStyle;
                c.SetCellValue(RandDigit(ref rand, right));
            }

            c = sheet.GetRow(row).CreateCell(col++, CellType.String);
            c.CellStyle = multStyle;
            c.SetCellValue("=");

            c = sheet.GetRow(row).CreateCell(col++, CellType.Formula);
            string colNameL = CellReference.ConvertNumToColString(baseCol + 1);
            string colNameR = CellReference.ConvertNumToColString(baseCol + 3);
            c.CellFormula = colNameL + (row + 1) + "*" + colNameR + (row + 1);
            if (point)
            {
                c.CellStyle = multAnsPointStyle;
            }
            else
            {
                c.CellStyle = multAnsStyle;
            }

            if (idx % 10 == 0)
            {
                sheet.GetRow(row).GetCell(baseCol + 0).CellStyle = indexStyleThick;
                if (point)
                {
                    sheet.GetRow(row).GetCell(baseCol + 1).CellStyle = multPointStyleThick;
                    sheet.GetRow(row).GetCell(baseCol + 3).CellStyle = multPointStyleThick;
                    sheet.GetRow(row).GetCell(baseCol + 5).CellStyle = multAnsPointStyleThick;
                }
                else
                {
                    sheet.GetRow(row).GetCell(baseCol + 1).CellStyle = multStyleThick;
                    sheet.GetRow(row).GetCell(baseCol + 3).CellStyle = multStyleThick;
                    sheet.GetRow(row).GetCell(baseCol + 5).CellStyle = multAnsStyleThick;
                }
                sheet.GetRow(row).GetCell(baseCol + 2).CellStyle = multStyleThick;
                sheet.GetRow(row).GetCell(baseCol + 4).CellStyle = multStyleThick;
            }
        }

        ICellStyle indexStyle;
        ICellStyle indexStyleThick;
        ICellStyle quizStyle;
        ICellStyle quizPointStyle;
        ICellStyle multStyle;
        ICellStyle multStyleThick;
        ICellStyle multPointStyle;
        ICellStyle multPointStyleThick;
        ICellStyle ansStyle;
        ICellStyle ansPointStyle;
        ICellStyle multAnsStyle;
        ICellStyle multAnsStyleThick;
        ICellStyle multAnsPointStyle;
        ICellStyle multAnsPointStyleThick;
        ICellStyle headerStyle;
        ICellStyle titleStyle;
        ICellStyle headtextStyle;

        void InitializeWorkbook()
        {
            hssfworkbook = new HSSFWorkbook();

            IFont indexFont = hssfworkbook.CreateFont();
            indexFont.FontName = "宋体";
            indexFont.FontHeightInPoints = 12;
            indexFont.IsItalic = false;
            //indexFont.Boldweight = (short)FontBoldWeight.BOLD;

            indexStyle = hssfworkbook.CreateCellStyle();
            indexStyle.SetFont(indexFont);
            indexStyle.Alignment = HorizontalAlignment.Center;
            indexStyle.VerticalAlignment = VerticalAlignment.Center;
            indexStyle.BorderTop = BorderStyle.Thin;
            indexStyle.BorderBottom = BorderStyle.Thin;
            indexStyle.BorderLeft = BorderStyle.Thin;
            indexStyle.BorderRight = BorderStyle.Thin;

            indexStyleThick = hssfworkbook.CreateCellStyle();
            indexStyleThick.CloneStyleFrom(indexStyle);
            indexStyleThick.BorderBottom = BorderStyle.Thick;

            IFont quizFont = hssfworkbook.CreateFont();
            quizFont.FontName = "AFont";
            quizFont.FontHeightInPoints = 16;

            quizStyle = hssfworkbook.CreateCellStyle();
            quizStyle.SetFont(quizFont);
            quizStyle.BorderLeft = BorderStyle.Thin;
            quizStyle.BorderRight = BorderStyle.Thin;
            quizStyle.Alignment = HorizontalAlignment.Right;
            quizStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");

            quizPointStyle = hssfworkbook.CreateCellStyle();
            quizPointStyle.CloneStyleFrom(quizStyle);
            quizPointStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.00");

            multStyle = hssfworkbook.CreateCellStyle();
            multStyle.SetFont(quizFont);
            multStyle.BorderTop = BorderStyle.Thin;
            multStyle.BorderBottom = BorderStyle.Thin;
            multStyle.Alignment = HorizontalAlignment.Center;
            multStyle.VerticalAlignment = VerticalAlignment.Center;
            multStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");

            multStyleThick = hssfworkbook.CreateCellStyle();
            multStyleThick.CloneStyleFrom(multStyle);
            multStyleThick.BorderBottom = BorderStyle.Thick;

            multPointStyle = hssfworkbook.CreateCellStyle();
            multPointStyle.CloneStyleFrom(multStyle);
            multPointStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.0##");

            multPointStyleThick = hssfworkbook.CreateCellStyle();
            multPointStyleThick.CloneStyleFrom(multPointStyle);
            multPointStyleThick.BorderBottom = BorderStyle.Thick;

            IFont ansFont = hssfworkbook.CreateFont();
            ansFont.FontName = "Lucida Console";
            ansFont.FontHeightInPoints = 14;

            ansStyle = hssfworkbook.CreateCellStyle();
            ansStyle.SetFont(ansFont);
            ansStyle.BorderTop = BorderStyle.Thin;
            ansStyle.BorderBottom = BorderStyle.Thin;
            ansStyle.BorderLeft = BorderStyle.Thin;
            ansStyle.BorderRight = BorderStyle.Thin;
            ansStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");

            ansPointStyle = hssfworkbook.CreateCellStyle();
            ansPointStyle.CloneStyleFrom(ansStyle);
            ansPointStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.00");

            multAnsStyle = hssfworkbook.CreateCellStyle();
            multAnsStyle.SetFont(ansFont);
            multAnsStyle.BorderTop = BorderStyle.Thin;
            multAnsStyle.BorderBottom = BorderStyle.Thin;
            multAnsStyle.BorderRight = BorderStyle.Thin;
            multAnsStyle.Alignment = HorizontalAlignment.Right;
            multAnsStyle.VerticalAlignment = VerticalAlignment.Center;
            multAnsStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");

            multAnsStyleThick = hssfworkbook.CreateCellStyle();
            multAnsStyleThick.CloneStyleFrom(multAnsStyle);
            multAnsStyleThick.BorderBottom = BorderStyle.Thick;

            multAnsPointStyle = hssfworkbook.CreateCellStyle();
            multAnsPointStyle.CloneStyleFrom(multAnsStyle);
            multAnsPointStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.0###");

            multAnsPointStyleThick = hssfworkbook.CreateCellStyle();
            multAnsPointStyleThick.CloneStyleFrom(multAnsPointStyle);
            multAnsPointStyleThick.BorderBottom = BorderStyle.Thick;

            IFont headerFont = hssfworkbook.CreateFont();
            headerFont.FontName = "宋体";
            headerFont.FontHeightInPoints = 16;

            headerStyle = hssfworkbook.CreateCellStyle();
            headerStyle.SetFont(headerFont);
            headerStyle.BorderTop = BorderStyle.Thin;
            headerStyle.BorderBottom = BorderStyle.Thin;
            headerStyle.BorderLeft = BorderStyle.Thin;
            headerStyle.BorderRight = BorderStyle.Thin;

            IFont titleFont = hssfworkbook.CreateFont();
            titleFont.FontName = "宋体";
            titleFont.FontHeightInPoints = 36;

            titleStyle = hssfworkbook.CreateCellStyle();
            titleFont.Boldweight = (short)FontBoldWeight.Bold;
            titleStyle.SetFont(titleFont);
            titleStyle.Alignment = HorizontalAlignment.Center;

            IFont headtextFont = hssfworkbook.CreateFont();
            headtextFont.FontName = "宋体";
            headtextFont.FontHeightInPoints = 28;
            headtextFont.Boldweight = (short)FontBoldWeight.Bold;

            headtextStyle = hssfworkbook.CreateCellStyle();
            headtextStyle.SetFont(headtextFont);
            headtextStyle.BorderBottom = BorderStyle.Thin;
            headtextStyle.Alignment = HorizontalAlignment.Center;
            headtextStyle.VerticalAlignment = VerticalAlignment.Center;

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
