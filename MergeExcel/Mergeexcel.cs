﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Data;
using System.Data.OleDb; 
using CellType = NPOI.SS.UserModel.CellType;
using JGNet.Common;
using CJBasic;
using JieXi.Common;
using CJBasic.Helpers;
using CJBasic.CJBasic.Helpers;
using CCWin.SkinClass;
using NPOI.SS.Formula.Functions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace MergeExcel
{
    public partial class Mergeexcel : BaseForm
    {

        //本地打开文件夹路径
        public string localOpenPath;
        public string localOpenPath2;

        //本地存储路径
        public string localSavePath;


        public Mergeexcel()
        {
            InitializeComponent();
        }


        //打开Bom文件按钮
        private void btnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open = new OpenFileDialog();
                if (open.ShowDialog() == DialogResult.OK)
                {
                    openPath.Text = System.IO.Path.GetFullPath(open.FileName);
                }
                localOpenPath = System.IO.Path.GetFullPath(open.FileName);
            }
            catch (Exception)
            {

                openPath.Text = string.Empty;
            }
        }
        //打开库房文件按钮
        private void btnOpen2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog open = new OpenFileDialog();
                if (open.ShowDialog() == DialogResult.OK)
                {
                    openPath2.Text = System.IO.Path.GetFullPath(open.FileName);
                }

                localOpenPath2 = System.IO.Path.GetFullPath(open.FileName);
            }
            catch (Exception)
            {

                openPath2.Text = string.Empty;
            }
        }

        //保存本地按钮
        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel 工作簿（*.xlsx）|*.xlsx|Excel 启动宏的工作簿（*.xlsm）|*.xlsm|Excel 97-2003工作簿（*.xls）|*.xls";
            saveFileDialog1.FileName = "新表单";//设置默认文件名
            saveFileDialog1.RestoreDirectory = true;//保存对话框是否记忆上次打开的目录

            saveFileDialog1.CheckPathExists = true;//检查目录

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                savePath.Text = saveFileDialog1.FileName;
            }
            localSavePath = saveFileDialog1.FileName;//文件路径
        }


        //开始操作按钮
        private void btnOpreat_Click(object sender, EventArgs e)
        {
            try
            {

                CbGeneric cb = new CbGeneric(this.DoShowUnCheck);
                cb.BeginInvoke(null, null);
                ShowProgressForm();
            }
            catch (Exception ex)
            {

                ShowError(ex);
            }
            finally
            {
                CompleteProgress();
                UnLockPage();
                ShowMessage("保存成功");
            }
        }

        private void DoShowUnCheck()
        {
            IWorkbook writeBook = null;
            try
            {

                label1.Text = string.Empty;

                FileInfo newFile = new FileInfo(localSavePath);
                if (newFile.Exists)
                {
                    newFile.Delete();
                }

                if (Path.GetExtension(localOpenPath).Equals(".xlsm"))
                {
                    localOpenPath = ChangeExtension1(localOpenPath);
                    //FileInfo newFile = new FileInfo(localSavePath);
                    //if (newFile.Exists)
                    //{
                    //    newFile.Delete();
                    //}
                    // Path.ChangeExtension("xlsx");
                }

                // FileInfo modelFile = new FileInfo(localOpenPath);
                //modelFile.CopyTo(localSavePath);
                FileStream fsRead = new FileStream(localOpenPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                fsRead.Position = 0;
                //  StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default);
                ////读取Bom表
                //  FileStream fsRead = System.IO.File.OpenRead(localOpenPath);
                writeBook = getPath(localOpenPath, fsRead);
                //var sheet = writeBook.GetSheetAt(0);



                //XSSFWorkbook writeBook = null;
                //writeBook = new XSSFWorkbook();
                //ISheet newsheet = writeBook.CreateSheet("Sheet1");
                //XSSFSheet writeSheet = (XSSFSheet)writeBook.GetSheetAt(0);

                ////样式
                //ICellStyle style1 = writeBook.CreateCellStyle();
                //style1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文字水平对齐方式
                //style1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式

                ////单元格边框样式
                //style1.BorderBottom = BorderStyle.Thin;
                //style1.BorderLeft = BorderStyle.Thin;
                //style1.BorderRight = BorderStyle.Thin;
                //style1.BorderTop = BorderStyle.Thin;
                //style1.WrapText = true;//自动换行ho
                //style1.BottomBorderColor = IndexedColors.Black.Index;//设置边框颜色

                //try
                //{
                //    //根据Bom表内容创建新表
                //    for (int k = 0; k < sheet.LastRowNum + 1; k++)
                //    {
                //        //获取第i行，得到对象
                //        IRow row = sheet.GetRow(k);
                //        //新建第i行，并返回得到的对象
                //        IRow writeRow = writeSheet.CreateRow(k);
                //        for (int j = 0; j < row.LastCellNum + 1; j++)
                //        {
                //            ICell cell = row.GetCell(j);
                //            if (cell == null)
                //            {
                //                break;
                //            }

                //            row.GetCell(j).SetCellType(CellType.String);
                //            string readValue = sheet.GetRow(k).GetCell(j).StringCellValue;

                //            if (string.IsNullOrEmpty(readValue))
                //            {

                //                readValue = string.Empty;
                //            }
                //            //新建第i行，第j列
                //            writeRow.CreateCell(j);
                //            writeRow.GetCell(j).CellStyle = style1;
                //            writeRow.GetCell(j).SetCellType(CellType.String);
                //            writeSheet.GetRow(k).GetCell(j).SetCellValue(readValue);
                //        }


                //    }

                //}
                //catch (Exception)
                //{
                //    label1.Width = 300;
                //    label1.Text = "请选择正确文件";
                //}
                // writeBook.Write(writeFile);
                //var workbook2 = new XSSFWorkbook(fy);

                //FileStream fsRead = System.IO.File.OpenRead(localSavePath);
                //IWorkbook wkcopy = getPath(localSavePath, fsRead);

                // 读取新表
                var sheetcopy2 = writeBook.GetSheetAt(0);
                IRow headrow = sheetcopy2.GetRow(2);
                IRow headrow2 = sheetcopy2.GetRow(3);
                IRow h1 = sheetcopy2.GetRow(1);
                IRow[] headrows = new IRow[] { headrow, headrow2, h1 };
                ////在最后添加新列
                //headrow.CreateCell(headrow.LastCellNum).SetCellType(CellType.String);
                //headrow.GetCell(headrow.LastCellNum - 1).CellStyle = style1;
                //headrow.CreateCell(headrow.LastCellNum - 1).SetCellValue("商品编码");

                //headrow.CreateCell(headrow.LastCellNum).SetCellType(CellType.String);
                //headrow.GetCell(headrow.LastCellNum - 1).CellStyle = style1;
                //headrow.CreateCell(headrow.LastCellNum - 1).SetCellValue("Count");



                //Reference列
                //  int refcell = 0;
                //Part 列
                int PartCell = 1;
                int remarkCell = 65;
                int changeColumn = 38;
                int leaveC = 8;
                int startIndexC = 8;
                //商品编码列
                //  int code = 0;

                //   int count = 0;
                //获取列号

                foreach (var item in headrows)
                {
                    for (int i = 0; i < item.LastCellNum - 1; i++)
                    {

                        if (item.Cells[i].CellType != CellType.String)
                            continue;
                        //item.Cells[i].SetCellType(CellType.String);
                        if (item.Cells[i].StringCellValue.Trim().Equals("姓名"))
                        {
                            PartCell = item.Cells[i].ColumnIndex;
                        }
                        if (item.Cells[i].StringCellValue.Trim().Equals("调整"))
                        {
                            changeColumn = item.Cells[i].ColumnIndex;
                        }
                        if (item.Cells[i].StringCellValue.Trim().Equals("离职日期"))
                        {
                            leaveC = item.Cells[i].ColumnIndex;
                            startIndexC = leaveC + 1;
                        }

                        // item.Cells[i].SetCellType(CellType.String);
                        if (item.Cells[i].StringCellValue.Trim().Equals("备注"))
                        {
                            remarkCell = item.Cells[i].ColumnIndex;
                        }
                        //  }
                    }

                }
                //设置列宽【SetColumnWidth(列索引，N*256) 第二个参数是列宽 单位是1/256个字符宽度】
                //    newsheet.SetColumnWidth(refcell, 40 * 256);

                //if (headrow.Cells[refcell].StringCellValue == "Reference")
                //{
                //    List<string> pad = new List<string>();
                //    for (int r = 1; r < sheetcopy2.LastRowNum + 1; r++)
                //    {
                //        IRow row2 = sheetcopy2.GetRow(r);
                //        if (row2 != null)
                //        {
                //            //Reference
                //            row2.GetCell(refcell).SetCellType(CellType.String);
                //            string a = row2.GetCell(refcell).StringCellValue;
                //            pad.Add(a);
                //        }
                //    }
                //    for (int r = 1; r < sheetcopy2.LastRowNum + 1; r++)
                //    {
                //        IRow row2 = sheetcopy2.GetRow(r);
                //        row2.CreateCell(count).SetCellType(CellType.Numeric);
                //        row2.GetCell(count).CellStyle = style1;
                //        try
                //        {
                //            if (pad[r - 1] == null)
                //            {
                //                row2.GetCell(count).SetCellValue(string.Empty);
                //            }
                //            else
                //            {

                //                string refe = pad[r - 1].Replace("，", ",").Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "");
                //                string[] ss = refe.ToString().Split(',');
                //                row2.GetCell(count).SetCellValue(ss.Length);

                //            }
                //        }
                //        catch (Exception)
                //        {

                //            row2.GetCell(count).SetCellValue(string.Empty);
                //        }
                //    }
                //}

                //workbook2.Write(fy);
                //fy.Close();


                using (FileStream fsRead2 = System.IO.File.OpenRead(localOpenPath2))
                {
                    IWorkbook wk = getPath(localOpenPath2, fsRead2);

                    ISheet sheet2 = wk.GetSheetAt(0);
                    IRow headrow3 = sheet2.GetRow(2);
                    //商品编码 列
                    int ComcodeCell = 5;
                    int ComnameCell = 35;
                    for (int i = 0; i < headrow3.LastCellNum - 1; i++)
                    {
                        if (headrow3.Cells[i].CellType != CellType.String)
                            continue;
                        if (headrow3.Cells[i].StringCellValue == "姓名")
                        {
                            ComcodeCell = headrow3.Cells[i].ColumnIndex;
                        }
                        if (headrow3.Cells[i].StringCellValue == "工号")
                        {
                            ComnameCell = headrow3.Cells[i].ColumnIndex;
                        }
                        //}
                    }
                    Dictionary<String, IRow> myDictionary = new Dictionary<String, IRow>();
                    string a = null;
                    string b = null;
                    string md = null;
                    try
                    {
                        if (headrow3.Cells[ComcodeCell].StringCellValue == "姓名" && headrow3.Cells[ComnameCell].StringCellValue == "工号")
                        {

                            ISheet sheet3 = wk.GetSheetAt(0);
                            for (int r = 3; r < sheet3.LastRowNum; r++)
                            {
                                IRow row2 = sheet3.GetRow(r);
                                if (row2 != null)
                                {
                                    //商品编码
                                    row2.GetCell(ComcodeCell).SetCellType(CellType.String);
                                    a = row2.GetCell(ComcodeCell).StringCellValue;

                                    //商品名称
                                    row2.GetCell(ComnameCell).SetCellType(CellType.String);
                                    b = row2.GetCell(ComnameCell).StringCellValue;
                                    U<String, String> keyValue = new U<string, string>(b, a);
                                    if (!String.IsNullOrEmpty(a))
                                    {
                                        try
                                        {
                                            myDictionary.Add(a, row2);
                                        }
                                        catch
                                        {
                                            throw new Exception("重复打卡人姓名:" + a);
                                        }
                                    }
                                }
                            }

                            //d
                            List<int> holiday = new List<int>();
                            List<int> ban = new List<int>();
                            //foreach (ICell cell in rowTitle.Cells)
                            //{
                            //    cell.SetCellType(CellType.String);
                            //    if (cell.StringCellValue.Equals("假"))
                            //    {
                            //        holiday.Add(cell.ColumnIndex);
                            //    }
                            //}

                            //计算是不是周六周日，还有法定节假日

                            IRow rowTitle = sheetcopy2.GetRow(0);
                            IRow rowTitle1 = sheetcopy2.GetRow(1);
                            IRow rowTitle2 = sheetcopy2.GetRow(2);
                            String workDays = rowTitle.Cells[1].StringCellValue;
                            Int32 datetime = Int32.Parse(rowTitle1.Cells[0].NumericCellValue.ToString());

                            DateTime dateMonth = DateTime.Parse("1900/01/01").AddDays(datetime - 1);
                            List<HolidayList> holidayLists = HolidayHelper.GetHolidayMonth(dateMonth.Year, dateMonth.Month);
                            foreach (var item in holidayLists)
                            {
                                String[] days = item.date.Split('-');
                                int day = Int32.Parse(days[days.Length - 1]);
                                if (item.status == 1)
                                {
                                    if (!holiday.Contains(day))
                                    {

                                        holiday.Add(day);
                                    }
                                }
                                else
                                {
                                    if (!ban.Contains(day))
                                    {

                                        ban.Add(day);
                                    }
                                }
                            }

                            for (int i = 0; i < 31; i++)
                            {

                                DateTime date = DateTime.Parse(dateMonth.Year + "-" + dateMonth.Month + "-" + (i + 1));

                                String value = String.Empty;
                                switch (date.DayOfWeek)
                                {
                                    case DayOfWeek.Sunday:
                                        value = "日";
                                        break;
                                    case DayOfWeek.Monday:
                                        value = "一";
                                        break;
                                    case DayOfWeek.Tuesday:
                                        value = "二";
                                        break;
                                    case DayOfWeek.Wednesday:
                                        value = "三";
                                        break;
                                    case DayOfWeek.Thursday:
                                        value = "四";
                                        break;
                                    case DayOfWeek.Friday:
                                        value = "五";
                                        break;
                                    case DayOfWeek.Saturday:
                                        value = "六";
                                        break;
                                    default:
                                        break;
                                }
                                if (changeColumn > startIndexC + i)
                                {
                                    rowTitle1.GetCell(startIndexC + i).SetCellValue(value);
                                }


                                if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                                {
                                    if (date.Month == dateMonth.Month)
                                    {
                                        if (!holiday.Contains(date.Day) && !ban.Contains(date.Day))
                                        {
                                            holiday.Add(date.Day);
                                        }
                                    }
                                }
                            }

                            //  style1.FillBackgroundColor = IndexedColors.Blue.Index; 
                            // cell.CellStyle = style1;


                            ICellStyle style = writeBook.CreateCellStyle();
                            //关键点 IndexedColors.AQUA.getIndex() 对应颜色
                            // style.FillBackgroundColor=IndexedColors.Black.Index;
                            // style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                            for (int i = startIndexC; i < changeColumn; i++)
                            {
                                var cell = rowTitle.GetCell(i);
                                var cell1 = rowTitle1.GetCell(i);
                                var cell2 = rowTitle2.GetCell(i);

                                if (holiday.Contains(i - startIndexC + 1))
                                {
                                    cell.SetCellValue("假");
                                    //cell.CellStyle = style;
                                    //cell2.CellStyle = style;
                                    //cell1.CellStyle = style;
                                    cell.CellStyle.FillForegroundColor = IndexedColors.SkyBlue.Index;
                                    //cell.CellStyle.FillBackgroundColor = IndexedColors.SkyBlue.Index;
                                    cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                                    cell2.CellStyle.FillForegroundColor = IndexedColors.SkyBlue.Index;
                                    //cell2.CellStyle.FillBackgroundColor = IndexedColors.SkyBlue.Index;
                                    cell2.CellStyle.FillPattern = FillPattern.SolidForeground;
                                    cell1.CellStyle.FillForegroundColor = IndexedColors.SkyBlue.Index;
                                    //cell1.CellStyle.FillBackgroundColor = IndexedColors.SkyBlue.Index;
                                    cell1.CellStyle.FillPattern = FillPattern.SolidForeground;
                                    cell2.SetCellType(CellType.String);
                                    cell2.SetCellValue(i - startIndexC + 1);

                                }
                                else
                                {
                                    cell.CellStyle.FillForegroundColor = IndexedColors.White.Index;
                                    //cell.CellStyle.FillBackgroundColor = IndexedColors.White.Index;
                                    cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                                    //cell2.CellStyle.FillBackgroundColor = IndexedColors.White.Index;
                                    cell2.CellStyle.FillForegroundColor = IndexedColors.White.Index;
                                    cell2.CellStyle.FillPattern = FillPattern.SolidForeground;
                                    //  cell1.CellStyle.FillBackgroundColor = IndexedColors.White.Index;
                                    cell1.CellStyle.FillForegroundColor = IndexedColors.White.Index;
                                    cell1.CellStyle.FillPattern = FillPattern.SolidForeground;
                                    cell.SetCellValue(String.Empty);
                                    cell2.SetCellType(CellType.String);
                                    cell2.SetCellValue(i - startIndexC + 1);
                                    // cell.CellStyle.FillBackgroundColorColor = 
                                }
                            }

                            InitProgress(sheetcopy2.LastRowNum + 1 - 4);
                            for (int r = 4; r < sheetcopy2.LastRowNum + 1; r++)
                            {
                                IRow row2 = sheetcopy2.GetRow(r);
                                //row2.CreateCell(code).SetCellType(CellType.String);
                                // row2.GetCell(code).CellStyle = style1;
                                if (row2 == null) { continue; }
                                row2.GetCell(PartCell).SetCellType(CellType.String);
                                md = row2.GetCell(PartCell).StringCellValue;
                                IRow row = null;
                                try
                                {
                                    row = myDictionary[md];
                                    UpdateProgress(md);
                                }
                                catch
                                { }
                                getPath(startIndexC, md, row, row2, holiday, changeColumn, remarkCell);
                            }
                        }

                        //判断选择保存路径是否存在该文件，存在则替换
                        newFile = new FileInfo(localSavePath);
                        if (newFile.Exists)
                        {
                            newFile.Delete();
                        }
                        //创建新表

                        FileStream writeFile = new FileStream(localSavePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
                        // FileStream writeFile = new FileStream(localSavePath, FileMode.OpenOrCreate);
                        writeFile.Position = 0;
                        // writeBook.w
                        //writeBook.
                        writeBook.Write(writeFile);

                        writeFile.Close();

                        // UpdateProgress( "保存成功");
                    }
                    finally
                    {
                        if (fsRead != null) { fsRead.Close(); }
                        if (fsRead2 != null) { fsRead2.Close(); }
                        //fsRead.Close();
                        //fsRead2.Close();
                    }


                    //try
                    //{

                    //    FileStream fs = new FileStream(localSavePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    //    fs.Position = 0;
                    //    StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default);

                    //    StringBuilder sb = new StringBuilder();
                    //    while (!sr.EndOfStream)
                    //    {
                    //        sb.AppendLine(sr.ReadLine() + "<br>");
                    //    }
                    //    FileStream saveRead = System.IO.File.OpenRead(localSavePath);
                    //    IWorkbook wb = getPath(localSavePath, fs);
                    //    int n = wb.NumberOfSheets;
                    //    if (n > 0)
                    //    {
                    //        label1.Text = "保存成功";
                    //        writeFile.Close();
                    //        fsRead.Close();
                    //        fsRead2.Close();
                    //        fs.Close();
                    //    }
                    //}
                    //catch (Exception ex)
                    //{
                    //    label1.ForeColor = Color.Red;
                    //    label1.Width = 200;
                    //    writeFile.Close();
                    //    fsRead.Close();
                    //    fsRead2.Close();
                    //    label1.Text = "请选择正确文件";
                    //}
                }
                UpdateProgress("保存成功");
            }
            catch (Exception ex)
            {
                ShowError(ex);
                //   label1.Text = ex.Message;     

            }
            finally
            {
                Complete(writeBook);
            }

        }

        private String ChangeExtension1(String filename)
        {
            String newName = filename;
            byte[] byteArray = File.ReadAllBytes(filename);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(stream, true))
                {
                    // Change from template type to workbook type
                    spreadsheetDoc.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                }
                File.WriteAllBytes(newName =Path.ChangeExtension(filename, "xlsx"), stream.ToArray());
            }
            return newName;
        }

        private String ChangeExtension(String filename)
        {
            String newName = filename;
            IWorkbook workbook;

            //read original xlsm file into workbook
            //using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read))
            // { file.Position = 0;
            //   workbook = new XSSFWorkbook(file);

            FileStream fsRead = new FileStream(localOpenPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            fsRead.Position = 0;
            //  StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default);
            ////读取Bom表
            //  FileStream fsRead = System.IO.File.OpenRead(localOpenPath);
            workbook = getPath(localOpenPath, fsRead);
            //} 

            //change file extension to xlsx and save in a new location
            newName = Path.ChangeExtension(filename, "xlsx");
            //if (!Directory.Exists("NewFile"))
            //    Directory.CreateDirectory("NewFile");
            FileStream stream = new FileStream((newName), FileMode.Create, System.IO.FileAccess.Write);
            workbook.Write(stream);
            stream.Close();
            fsRead.Close();

            ////read the newly created file from the new location
            //using (FileStream file = new FileStream(@"NewFile\\New" + filename, FileMode.Open, FileAccess.Read))
            //{ workbook = new XSSFWorkbook(file); }
            return newName;
        }



        private void InitProgress(int num)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new CJBasic.CbGeneric<int>(this.InitProgress), num);
            }
            else
            {
                this.dataMergedView1.DataSource = null;
                base.InitProgress(num);
            }
        }

        private void Complete(IWorkbook writeBook)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new CJBasic.CbGeneric<IWorkbook>(this.Complete), writeBook);
            }
            else
            {
                CompleteProgress();
                FileStream fsRead = null;
                try
                {
                    fsRead = new FileStream(localSavePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    fsRead.Position = 0;
                    //  StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default);
                    ////读取Bom表
                    //  FileStream fsRead = System.IO.File.OpenRead(localOpenPath);
                    writeBook = getPath(localSavePath, fsRead);
                    //  dataMergedView1.DataSource = NPOIHelper.FormatToDatatable(writeBook, 0);
                }
                catch { }
                finally
                {
                    if (fsRead != null) { fsRead.Close(); }

                }
            }
        }

        private void getPath(int leaveC, String md, IRow rows, IRow row2, List<int> holiday, int dayEnd, int remark)
        {
            int start = leaveC;
            int end = dayEnd;
            int workStart = 6;
            int workEnd = 36;
            //
            //int remark = 41;


            //
            row2.GetCell(remark).SetCellValue(String.Empty);

            foreach (var item in row2.Cells)
            {

                //.Name = "Wingdings";

                //   item.SetCellValue(" ");
                if (item.ColumnIndex >= start && item.ColumnIndex <= end)
                {

                    item.CellStyle.FillForegroundColor = IndexedColors.White.Index;
                    //item.CellStyle.FillForegroundColor = IndexedColors.White.Index;
                    item.CellStyle.FillPattern = FillPattern.SolidForeground;
                    if (holiday.Contains(item.ColumnIndex - start + 1))
                    {
                        item.CellStyle.FillForegroundColor = IndexedColors.SkyBlue.Index;
                        item.CellStyle.FillPattern = FillPattern.SolidForeground;
                        //item.CellStyle.FillForegroundColor = IndexedColors.SkyBlue.Index;
                        item.SetCellValue(String.Empty);
                    }
                    else
                    {
                        //item.CellStyle.FillBackgroundColor = IndexedColors.White.Index;
                        item.SetCellValue(g);
                        if (md.Equals(a) || md.Equals(""))
                        {
                            item.SetCellValue(g);
                            return;
                        }
                    }
                    //获取时间
                    String times = null;
                    //try
                    //{
                    //    times = row2.Cells[workStart + item.ColumnIndex - start].StringCellValue;
                    //}
                    //catch { item.SetCellValue(String.Empty); continue; }

                }
            }

            if (rows != null)
            {
                for (int i = 0; i < row2.Cells.Count; i++)
                {

                    ICell item = row2.Cells[i];
                    //item.SetCellType(CellType.String); 


                    if (item.ColumnIndex >= start && item.ColumnIndex <= end && !holiday.Contains(item.ColumnIndex - start + 1))
                    {

                        //判断迟到早退
                        //row2.
                        item.SetCellValue(String.Empty);

                        //获取时间
                        String times = null;
                        try
                        {
                            times = rows.Cells[workStart + item.ColumnIndex - start].StringCellValue;
                        }
                        catch { item.SetCellValue(String.Empty); continue; }

                        if (String.IsNullOrEmpty(times))
                        {
                            //item.
                            //  item.SetCellType(CellType.String);
                            item.SetCellValue(g);
                            //   row2.Cells[remark].SetCellValue(row2.Cells[remark].StringCellValue + "\n" + (item.ColumnIndex - start + 1) + "号旷工");

                        }
                        else
                        {
                            String[] timeArrays = times.Split('\n');

                            String beginTime = g;
                            String endTime = g;
                            String timeFirst = null;
                            String timeEnd = null;
                            if (timeArrays.Length < 2)
                            {
                                timeFirst = timeArrays[0].Trim().Replace(":", "");
                                // timeEnd = timeFirst;
                            }
                            else
                            {
                                timeFirst = timeArrays[0].Trim().Replace(":", "");
                                timeEnd = timeArrays[timeArrays.Length - 1].Trim().Replace(":", "");
                            }


                            if (Int32.Parse(timeFirst) <= 1200 && Int32.Parse(timeFirst) > 901)
                            {
                                beginTime = s;
                                row2.GetCell(remark).SetCellValue(row2.GetCell(remark).StringCellValue + "\n" + (item.ColumnIndex - start + 1) + "号迟到" + (Int32.Parse(timeFirst) - 900) + "分钟");

                            }
                            else if (Int32.Parse(timeFirst) >= 1200 && Int32.Parse(timeFirst) < 1300 || Int32.Parse(timeFirst) > 1800 || Int32.Parse(timeFirst) < 1800 && Int32.Parse(timeFirst) >= 1300)
                            {
                                beginTime = g;
                                //row2.Cells[remark].SetCellValue(row2.Cells[remark].StringCellValue + "\n" + (item.ColumnIndex - start + 1) + "号缺卡");
                            }
                            else
                            {
                                beginTime = y;
                            }

                            if (String.IsNullOrEmpty(timeEnd))
                            {
                                if (Int32.Parse(timeFirst) > 1800)
                                {
                                    endTime = y;
                                }
                            }
                            else
                            {//Int32.Parse(timeEnd) < 1800 && Int32.Parse(timeEnd) > 1300 

                                if (Int32.Parse(timeEnd) > 1700 && Int32.Parse(timeEnd) < 1800)
                                {
                                    //if (Int32.Parse(timeEnd) > 1530 && Int32.Parse(timeEnd) < 1800)
                                    //{
                                    endTime = s;
                                    row2.GetCell(remark).SetCellValue(row2.GetCell(remark).StringCellValue + "\n" + (item.ColumnIndex - start + 1) + "号早退" + (1800 - Int32.Parse(timeEnd)) + "分钟");
                                    //}
                                    //else {
                                    //    //
                                    //    endTime = "";
                                    //}
                                }
                                else if (Int32.Parse(timeEnd) < 1800 && Int32.Parse(timeEnd) > 1300 || (Int32.Parse(timeEnd) > 600 && Int32.Parse(timeEnd) < 900))
                                {
                                    //有记录就不算旷工
                                    endTime = g;
                                    //  row2.Cells[remark].SetCellValue(row2.Cells[remark].StringCellValue + "\n" + (item.ColumnIndex - start + 1) + "号缺卡");
                                }
                                else

                                {
                                    endTime = y;
                                }
                            }






                            //加班
                            //if (holiday.Contains(item.ColumnIndex))
                            //{
                            //    item.SetCellValue(" ");
                            //    //if (!beginTime.Equals("口") && !endTime.Equals("口")) { 
                            //    //row2.Cells[remark].SetCellValue(row2.Cells[remark].StringCellValue +"\n" + (item.ColumnIndex-start+1)+ "日加班");
                            //    //}
                            //}
                            // item.SetCellType(CellType.String);
                            item.SetCellValue(beginTime + "\n" + endTime);
                            if (beginTime.Equals(y) && endTime.Equals(y))
                            {
                                item.SetCellValue(y);
                            }
                        }

                        //   // item.SetCellValue("♀");

                        //   //if () { 
                        //   //}
                    }
                }

            }

        }


        //判断Excel文件类型
        public IWorkbook getPath(string strPath, FileStream fsRead)
        {
            //FileStream fsRead = System.IO.File.OpenRead(strPath);
            IWorkbook wk = null;
            //获取后缀名
            string extension = strPath.Substring(strPath.LastIndexOf(".")).ToString().ToLower();
            //判断是否是excel文件
            if (extension == ".xlsx" || extension == ".xlsm" || extension == ".xls")
            {
                //判断excel的版本
                if (extension == ".xlsm" || extension == ".xlsx")
                {
                    wk = new XSSFWorkbook(fsRead);
                    //
                }
                else
                {
                    wk = new HSSFWorkbook(fsRead);
                }
            }
            return wk;
        }


        private void mergeexcel_Load(object sender, EventArgs e)
        {

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            FileHelper.OpenFileInExplorer(savePath.Text);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }
    }
}
