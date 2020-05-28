using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Data;

namespace BaseLayer
{
    public partial class 图形 : Form
    {

        List<string> txDataPie = new List<string>() { "不合格", "合格" };
        List<int> tyDataPie = new List<int>() { 6, 4 };


        List<string> txDataColumn = new List<string>() { "2020-03-23", "2020-03-23", "2020-03-23", "2020-03-24", "2020-03-24", "2020-03-24", "2020-03-25", "2020-03-25", "2020-03-25", "2020-03-25" };
        List<string> tyDataOk = new List<string>() { "D10-23", "D10-S2", "D10-31", "D10-01", "D10-02", "D10-31", "D10-01", "D10-31", "D10-01", "D10-02" };
        List<int> tyDataNo = new List<int>() { 7, 40, 185, 7, 40, 18, 40, 40, 18, 40 };

        public 图形()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //申明保存对话框   
            SaveFileDialog dlg = new SaveFileDialog();
            //默然文件后缀   
            dlg.DefaultExt = "xlsx ";
            //文件后缀列表   
            dlg.Filter = "EXCEL文件(*.XLSX)|*.xlsx ";
            //默然路径是系统当前路径   
            dlg.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            //打开保存对话框   
            if (dlg.ShowDialog() == DialogResult.Cancel) return;
            //返回文件路径   
            string fileNameString = dlg.FileName;
            //验证strFileName是否为空或值无效   
            if (fileNameString.Trim() == " ")
            { return; }
            Excel.Application objExcel = new Excel.Application();//创建一个excel的实例

            Excel.Workbook objWorkbook = null;

            Excel.Worksheet objsheet = null;


            //Excel.Sheets objsheets = objExcel.Workbooks.Item[1].Worksheets;
            //objsheets.Item[1] = "";
            //Excel.Worksheet ws = (Excel.Worksheet)objExcel.Worksheets.get_Item(1);
            //ws.Name = "狐狸!";




            // Excel.Sheets objsheets = objExcel.Sheets;
            //var ss =  objsheets.Item[1];
            // var aa = (Excel.Worksheet)ss;
            // aa.Name = "dasdasdasd";
            //objsheets["Sheet1"].Nasdasda = "11";


            //objsheet = (Excel.Worksheet)objsheets.get_Item(objsheets.Count);

            //worksheet.Name = "sadas ";
            ////objsheets.get_Item(1


            //objWorkbooks = objExcel.Workbooks;
            ////oBooks.Open(sTemplate, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //objWorkbook = objWorkbooks.get_Item(1);
            //objsheets = objWorkbook.Worksheets;

            //objsheet = (Excel.Worksheet)objsheets.get_Item(1);
            ////命名该sheet
            //objsheet.Name = "Sheet1";

            try
            {
                //申明对象   
                objExcel = new Excel.Application();
                objWorkbook = objExcel.Workbooks.Add(Missing.Value);
                objsheet = (Excel.Worksheet)objWorkbook.ActiveSheet;
                //合格率
                objExcel.Cells[1, 1] = "客户";
                objExcel.Cells[1, 2] = "数量";
                objExcel.Cells[1, 3] = "占比";

                objExcel.Cells[2, 1] = "不合格";
                objExcel.Cells[3, 1] = "合格";
                objExcel.Cells[2, 2] = tyDataPie[0];
                objExcel.Cells[3, 2] = tyDataPie[1];
                objExcel.Cells[2, 3] = @"=B2 / B4";
                objExcel.Cells[3, 3] = @"=B3 / B4";

                //求和
                Excel.Range rangesummary73 = objsheet.Range["B4"];//--ActiveCell = rangesummary110
                rangesummary73.Formula = "=SUM(R[-2]C:R[-1]C)";
                rangesummary73.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //ActiveCell.FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)";
                //            Range("A2:A3,C2:C3").Select
                //Range("C2").Activate
                //ActiveSheet.Shapes.AddChart2(251, xlPie).Select
                //ActiveChart.SetSourceData Source:= Range("Sheet1!$A$2:$A$3,Sheet1!$C$2:$C$3")
                //ActiveChart.SetElement(msoElementDataLabelBestFit)

                //            Excel.Range rangesummary5 = objsheet.Range["A2:A3", "C2:C3"];


                //设置百分比格式
                Excel.Range rangesummary4 = objsheet.Range["C2", "C3"];
                //Range("C2:C3").Select;
                rangesummary4.NumberFormatLocal = "0.00%";
                //饼图

                //新建一个饼图
                Excel.Chart xlChart = (Excel.Chart)objWorkbook.Charts.Add(Type.Missing, objsheet, Type.Missing, Type.Missing);
                xlChart.ChartType = Excel.XlChartType.xlPie;//设置图形
                xlChart.SetSourceData(objsheet.get_Range("A1:A3, C1:C3"), Excel.XlRowCol.xlColumns);//两种方法都可以
                                                                                                    // xlChart.SetSourceData(objsheet.Range["A1:A3", "C1:C3"], Excel.XlRowCol.xlColumns);


                //加border和居中设置
                Excel.Range rangesummary110 = objsheet.Range["A1", "C3"];
                rangesummary110.Borders.Color = 0;
                rangesummary110.Borders.Weight = 2;
                rangesummary110.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //设置属性标签
                xlChart.SetElement(MsoChartElementType.msoElementDataLabelOutSideEnd);  //数据标签
                xlChart.SetElement(MsoChartElementType.msoElementLegendBottom);  //设为底部显示
                xlChart.SetElement(MsoChartElementType.msoElementChartTitleAboveChart);  //设置标题
                objsheet.Range["F:F"].ColumnWidth = 20.5;   //设置宽度
                xlChart.ChartTitle.Text = "客户占比(2020.3.23-3.28)";





                //objWorkbook.ActiveChart.Location(Excel.XlChartLocation.xlLocationAutomatic, "合格率");//xlLocationAutomatic :Excel 控制图表位置。
                objWorkbook.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, objsheet.Name); // xlLocationAsObject,将图表嵌入到现有工作表中。

                // oResizeRange = (Excel.Range)objsheet.Rows.get_Item(7, Missing.Value);

                objsheet.Shapes.Item("Chart 1").Top = 150;  //调图表的位置上边距
                objsheet.Shapes.Item("Chart 1").Left = 10;
                objsheet.Shapes.Item("Chart 1").Width = 200;   //调图表的宽度
                objsheet.Shapes.Item("Chart 1").Height = 250;  //调图表的高度




                ///////////////////////////////////////////////////////////////




                #region 管理人员
                int col = 6;
                objExcel.Cells[2, col] = "日期";
                objExcel.Cells[2, col + 1] = "产品名称";
                objExcel.Cells[2, col + 2] = "数量";
                int row = 3;
                for (int i = 0; i < txDataColumn.Count; i++)
                {
                    objExcel.Cells[row, col] = txDataColumn[i];
                    row++;
                }
                row = 3;
                for (int i = 0; i < tyDataOk.Count; i++)
                {
                    objExcel.Cells[row, col + 1] = tyDataOk[i];
                    row++;
                }
                row = 3;
                for (int i = 0; i < tyDataNo.Count; i++)
                {
                    objExcel.Cells[row, col + 2] = tyDataNo[i];
                    row++;
                }
                #endregion
                //假定要合并excel文件中第2行的1~3列，并且显示黑色边框7a686964616fe4b893e5b19e31333264656666
                //objExcel.ActiveSheet.Columns[7] = 5;
                Microsoft.Office.Interop.Excel.Range rangesummary1 = objsheet.Range["F3", "F5"];
                Microsoft.Office.Interop.Excel.Range rangesummary2 = objsheet.Range["F6", "F8"];
                Microsoft.Office.Interop.Excel.Range rangesummary3 = objsheet.Range["F9", "F12"];
                Excel.Range rangesummary120 = objsheet.Range["F2", "H12"];
                rangesummary120.Borders.Color = 0;
                rangesummary120.Borders.Weight = 2;
                rangesummary120.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangesummary120.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;


                rangesummary1.Value2 = Type.Missing;
                rangesummary1.Merge(Type.Missing);
                rangesummary1.Value2 = "2020-03-23";


                rangesummary2.Value2 = Type.Missing;
                rangesummary2.Merge(Type.Missing);
                rangesummary2.Value2 = "2020-03-24";


                rangesummary3.Value2 = Type.Missing;
                rangesummary3.Merge(Type.Missing);
                rangesummary3.Value2 = "2020-03-25";


                //柱状图
                Excel.Chart xlChart2 = (Excel.Chart)objWorkbook.Charts.Add(Type.Missing, objsheet, Type.Missing, Type.Missing);
                Excel.Range cellRange = objsheet.get_Range((Excel.Range)objsheet.Cells[2, 6], (Excel.Range)objsheet.Cells[3 + txDataColumn.Count - 1, 8]);
                //1-cellRange:数据源的范围，2-图表类型，3-Type.Missing，4-在图表上将列或行用作数据系列的方式，5、6-第五个第六个参数设置图表的x轴和y轴分别是数据源的哪些列/行，7-图表是否有图例，8、9、10-设置标题
                xlChart2.ChartWizard(cellRange,
                                Excel.XlChartType.xlColumnClustered, //2-图表类型
                                Type.Missing,//内置自动套用格式的选项编号。 可为从 1 到 10 的数字，其取值依赖于库的类型。 如果省略此参数, 则 Excel 根据库的类型和数据源选择默认值。
                                Excel.XlRowCol.xlColumns, //在图表上将列或行用作数据系列的方式
                                2, //第五个第六个参数设置图表的x轴和y轴分别是数据源的哪些列/行--这个2代表数据源的x轴由两个参数确认；可以不写，默认的就很难看
                                1,//--这个2代表数据源的x轴由1个参数确认；可以不写，默认的就很难看
                                true, //图表是否有图例
                                "每日总量统计", //以下都是标题
                                Type.Missing,
                                Type.Missing,
                                "");
                xlChart2.SetElement(MsoChartElementType.msoElementDataLabelOutSideEnd);
                xlChart2.Location(Excel.XlChartLocation.xlLocationAsObject, objsheet.Name);

                Excel.Range oResizeRange1 = (Excel.Range)objsheet.Rows.get_Item(1);
                Excel.Range oResizeRange2 = (Excel.Range)objsheet.Columns.get_Item(10);
                objsheet.Shapes.Item("Chart 2").Top = (float)oResizeRange1.Top;  //调图表的位置上边距--1行的高度
                objsheet.Shapes.Item("Chart 2").Left = (float)(double)oResizeRange2.Left;//调图表的位置左边距--10列的宽度
                objsheet.Shapes.Item("Chart 2").Width = 300;   //调图表的宽度
                objsheet.Shapes.Item("Chart 2").Height = 200;  //调图表的高度

                //保存文件   
                objWorkbook.SaveAs(fileNameString, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "警告 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            finally
            {
                ////关闭Excel应用   
                //if (objWorkbook != null) objWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
                //if (objExcel.Workbooks != null) objExcel.Workbooks.Close();
                //if (objExcel != null) objExcel.Quit();
                //objsheet = null;
                //objWorkbook = null;
                //objExcel = null;
                objsheet.Name = "牙模盒使用记录";
                objsheet.Tab.Color = Excel.XlThemeColor.xlThemeColorLight1;
                ClosePro(fileNameString, objExcel, objWorkbook);
                System.Diagnostics.Process.Start(fileNameString);
            }

        }
        /// <summary>
        /// 关闭Excel进程
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="excel"></param>
        /// <param name="wb"></param>
        public void ClosePro(string excelPath, Excel.Application excel, Excel.Workbook wb)
        {

            Process[] localByNameApp = Process.GetProcessesByName(excelPath);//获取程序名的所有进程
            if (localByNameApp.Length > 0)
            {
                foreach (var app in localByNameApp)
                {
                    if (!app.HasExited)
                    {
                        #region
                        ////设置禁止弹出保存和覆盖的询问提示框   
                        //excel.DisplayAlerts = false;
                        //excel.AlertBeforeOverwriting = false;

                        ////保存工作簿   
                        //excel.Application.Workbooks.Add(true).Save();
                        ////保存excel文件   
                        //excel.Save("D:" + "\\test.xls");
                        ////确保Excel进程关闭   
                        //excel.Quit();
                        //excel = null; 
                        #endregion
                        app.Kill();//关闭进程  
                    }
                }
            }
            if (wb != null)
                wb.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            // 安全回收进程
            System.GC.GetGeneration(excel);
        }
        ExcelHelper helper = new ExcelHelper();

        /// <summary>
        /// //////////////////////////////////////////////////////////////////////////////
        /// </summary>


























        System.Data.DataTable dt;

        //详细信息
        private void button2_Click(object sender, EventArgs e)
        {
           string sql = " SELECT CONVERT(VARCHAR(10),CAST(m.AcceptDate AS DATE),120) AS [date],md.ProductName ,m.HName,SUM(md.Number)AS [count]";
            sql += " FROM dbo.Mechanic m LEFT JOIN MechanicDetail md ON md.MID = m.ID ";
            sql += " WHERE md.ProductName IS NOT NULL  AND left(convert(char(10),AcceptDate,120),10) BETWEEN '2020-02-13' AND '2020-04-19'";
            sql += " GROUP BY md.ProductName,m.HName,CAST(m.AcceptDate AS DATE)";
            sql += " ORDER BY CAST(m.AcceptDate AS DATE)";
            dt = SqlHelper.ExecuteDataTable(sql);

            helper.ShowDetail(dataGridView1);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<string> txDataColumn = new List<string>();
            List<string> tyProductColumn = new List<string>();
            List<int> tyNumberColumn = new List<int>();

            String sql;
            sql = " SELECT CONVERT(VARCHAR(10),CAST(m.AcceptDate AS DATE),120) AS '日期',md.ProductName '产品名称',SUM(md.Number)AS 数量";
            sql += " FROM dbo.Mechanic m LEFT JOIN MechanicDetail md ON md.MID = m.ID ";
            sql += " WHERE md.ProductName IS NOT NULL  AND left(convert(char(10),AcceptDate,120),10) BETWEEN '2020-02-13' AND '2020-04-19'";
            sql += " GROUP BY md.ProductName,m.HName,CAST(m.AcceptDate AS DATE)";
            sql += " ORDER BY CAST(m.AcceptDate AS DATE)";
            dt = SqlHelper.ExecuteDataTable(sql);
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count - 1; i++)
                {
                    txDataColumn.Add(dt.Rows[i]["日期"].ToString());
                    tyProductColumn.Add(dt.Rows[i]["产品名称"].ToString());
                    tyNumberColumn.Add((int)dt.Rows[i]["数量"]);
                }
            }




            helper.ExportRectangle(txDataColumn, tyProductColumn, tyNumberColumn);

            //helper.ExportRectangle();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string DateBegin = "2020.3.23";
            string DateEnd = "2020.3.28";

            List<string> txClientColumn = new List<string>();
            List<int> tyNumberColumn = new List<int>();

            string sql = " SELECT m.HName AS '客户',SUM(md.Number)AS 数量";
            sql += " FROM dbo.Mechanic m LEFT JOIN MechanicDetail md ON md.MID = m.ID";
            sql += " WHERE md.ProductName IS NOT NULL  AND left(convert(char(10),AcceptDate,120),10) BETWEEN '2020-02-13' AND '2020-04-19'";
            sql += " GROUP BY m.HName";

            dt = SqlHelper.ExecuteDataTable(sql);

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count - 1; i++)
                {
                    txClientColumn.Add(dt.Rows[i]["客户"].ToString());
                    tyNumberColumn.Add((int)dt.Rows[i]["数量"]);
                }
            }

            helper.ExportCircle(txClientColumn, tyNumberColumn, DateBegin, DateEnd);

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void 图形_Load(object sender, EventArgs e)
        {

            string sql = "SELECT CONVERT(VARCHAR(10),CAST(m.AcceptDate AS DATE),120) AS [date],md.ProductName ,m.HName,SUM(md.Number)AS [count]";
            sql += " FROM dbo.Mechanic m LEFT JOIN MechanicDetail md ON md.MID = m.ID";
            sql += " WHERE md.ProductName IS NOT NULL  AND left(convert(char(10),AcceptDate,120),10) BETWEEN '2020-02-13' AND '2020-04-19'";
            sql += " GROUP BY md.ProductName,m.HName,CAST(m.AcceptDate AS DATE)";
            sql += " ORDER BY CAST(m.AcceptDate AS DATE)";
            dt = SqlHelper.ExecuteDataTable(sql);

            int index;
            if (dt.Rows.Count>0)
            {
               
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    index = dataGridView1.Rows.Add();
                    dataGridView1.Rows[index].Cells[0].Value = dt.Rows[i]["date"].ToString();
                    dataGridView1.Rows[index].Cells[1].Value = dt.Rows[i]["HName"].ToString();
                    dataGridView1.Rows[index].Cells[2].Value = dt.Rows[i]["ProductName"].ToString();
                    dataGridView1.Rows[index].Cells[3].Value = dt.Rows[i]["count"].ToString();

                    dataGridView1.Rows[index].HeaderCell.Value = (index + 1).ToString();
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelHelper2 excelHelper2 = new ExcelHelper2();
            excelHelper2.Test();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string DateBegin = "2020.3.23";
            string DateEnd = "2020.3.28";

            List<string> txClientColumn = new List<string>();
            List<int> tyNumberColumn = new List<int>();

            string sql = " SELECT m.HName AS '客户',SUM(md.Number)AS 数量";
            sql += " FROM dbo.Mechanic m LEFT JOIN MechanicDetail md ON md.MID = m.ID";
            sql += " WHERE md.ProductName IS NOT NULL  AND left(convert(char(10),AcceptDate,120),10) BETWEEN '2020-02-13' AND '2020-04-19'";
            sql += " GROUP BY m.HName";

            dt = SqlHelper.ExecuteDataTable(sql);

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count - 1; i++)
                {
                    txClientColumn.Add(dt.Rows[i]["客户"].ToString());
                    tyNumberColumn.Add((int)dt.Rows[i]["数量"]);
                }
            }

            List<string> txDataColumn = new List<string>();
            List<string> tyProductColumn = new List<string>();
            List<int> tyNumberColumn2 = new List<int>();

            sql = " SELECT CONVERT(VARCHAR(10),CAST(m.AcceptDate AS DATE),120) AS '日期',md.ProductName '产品名称',SUM(md.Number)AS 数量";
            sql += " FROM dbo.Mechanic m LEFT JOIN MechanicDetail md ON md.MID = m.ID ";
            sql += " WHERE md.ProductName IS NOT NULL  AND left(convert(char(10),AcceptDate,120),10) BETWEEN '2020-02-13' AND '2020-04-19'";
            sql += " GROUP BY md.ProductName,m.HName,CAST(m.AcceptDate AS DATE)";
            sql += " ORDER BY CAST(m.AcceptDate AS DATE)";
            dt = SqlHelper.ExecuteDataTable(sql);
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count - 1; i++)
                {
                    txDataColumn.Add(dt.Rows[i]["日期"].ToString());
                    tyProductColumn.Add(dt.Rows[i]["产品名称"].ToString());
                    tyNumberColumn2.Add((int)dt.Rows[i]["数量"]);
                }
            }




            //helper.ExportRectangle(txDataColumn, tyProductColumn, tyNumberColumn);
            //helper.ExportCircle();
            ExcelHelper3 helper3 = new ExcelHelper3();
            helper3.MainExport(dataGridView1, txClientColumn, tyNumberColumn, DateBegin, DateEnd, txDataColumn, tyProductColumn, tyNumberColumn2);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            List<string> txDataTime = new List<string>();
            List<string> txClient = new List<string>();
            List<string> txProduct= new List<string>();
            List<int> txCount= new List<int>();

            string sql = " SELECT CONVERT(VARCHAR(10),CAST(m.AcceptDate AS DATE),120) AS [date],md.ProductName ,m.HName,SUM(md.Number)AS [count]";
            sql += " FROM dbo.Mechanic m LEFT JOIN MechanicDetail md ON md.MID = m.ID ";
            sql += " WHERE md.ProductName IS NOT NULL  AND left(convert(char(10),AcceptDate,120),10) BETWEEN '2020-02-13' AND '2020-04-19'";
            sql += " GROUP BY md.ProductName,m.HName,CAST(m.AcceptDate AS DATE)";
            sql += " ORDER BY CAST(m.AcceptDate AS DATE)";
            dt = SqlHelper.ExecuteDataTable(sql);
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count - 1; i++)
                {
                    txDataTime.Add(dt.Rows[i]["date"].ToString());
                    txClient.Add(dt.Rows[i]["HName"].ToString());
                    txProduct.Add(dt.Rows[i]["ProductName"].ToString());
                    txCount.Add((int)dt.Rows[i]["count"]);
                }
            }

            helper.ShowDetail(txDataTime, txClient, txProduct, txCount);
        }
    }
}
