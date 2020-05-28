using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace QueryInterface
{
    public class ExcelHelper4
    {

        public int intExcelTempIndex = 0;
        bool Open = true;
        string fileName;

        /// <summary>
        /// 饼状图
        /// </summary>
        /// <param name="txClientColumn"></param>
        /// <param name="tyNumberColumn"></param>
        /// <param name="DateBegin">开始时间</param>
        /// <param name="DateEnd">结束时间</param>
        public void ExportCircle(List<string> txClientColumn, List<int> tyNumberColumn, string Title, string DateBegin, string DateEnd)
        {
            intExcelTempIndex++;
            Open = true;
            //List<string> txClientColumn = txClientColumn;
            //List<string> tyDataOk = tyNumberColumn;

            Excel.Application objExcel = null;//创建一个excel的实例
            Excel.Workbook objWorkbook = null;
            Excel.Worksheet objsheet = null;

            string fileNameString = "";

            try
            {
                //申明对象   
                objExcel = new Excel.Application();
                objWorkbook = objExcel.Workbooks.Add(Missing.Value);
                objsheet = (Excel.Worksheet)objWorkbook.ActiveSheet;

                objsheet.Range["A:A"].ColumnWidth = 30;   //设置宽度
                objsheet.Range["A:A"].NumberFormatLocal = "@";
                objsheet.Range["B:B"].ColumnWidth = 20;   //设置宽度
                objsheet.Range["C:C"].ColumnWidth = 20;   //设置宽度

                //合格率
                if (Title == "客户占比")
                {
                    objExcel.Cells[1, 1] = "客户";
                }
                else
                {
                    objExcel.Cells[1, 1] = "产品";
                }
                objExcel.Cells[1, 2] = "数量";
                objExcel.Cells[1, 3] = "占比";

                int col = 1;
                int row = 2;    //row和 i得对应关系是row = i+2 ; i = row -2

                for (int i = 0; i < txClientColumn.Count; i++)
                {
                    objExcel.Cells[row, col] = txClientColumn[i];
                    row++;
                }

                row = 2;
                for (int i = 0; i < tyNumberColumn.Count; i++)
                {
                    objExcel.Cells[row, col + 1] = tyNumberColumn[i];
                    row++;
                }

                int temp = tyNumberColumn.Count + 1;

                string sumCell = "B" + (temp + 1);
                string Cell = "B" + temp;


                //求和
                Excel.Range rangeSum = objsheet.Range[sumCell];//--ActiveCell = rangesummary110
                rangeSum.Formula = "=SUM(B2:" + Cell + ")";
                rangeSum.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                row = 2;
                for (int i = 0; i < tyNumberColumn.Count; i++)
                {

                    objExcel.Cells[row, col + 2] = @"=B" + (i + 2) + " / " + sumCell;
                    row++;
                }

                //设置百分比格式
                Excel.Range rangePercent = objsheet.Range["C2", "C" + temp];
                rangePercent.NumberFormatLocal = "0.00%";


                //新建一个饼图
                Excel.Chart xlChart = (Excel.Chart)objWorkbook.Charts.Add(Type.Missing, objsheet, Type.Missing, Type.Missing);
                xlChart.ChartType = Excel.XlChartType.xlPie;//设置图形
                xlChart.SetSourceData(objsheet.get_Range("A1:A" + temp + ", C1:C" + temp), Excel.XlRowCol.xlColumns);//两种方法都可以
                xlChart.ChartStyle = 251;  //设置风格                                                                 // xlChart.SetSourceData(objsheet.Range["A1:A3", "C1:C3"], Excel.XlRowCol.xlColumns);


                //加border和居中设置
                Excel.Range rangesummary110 = objsheet.Range["A1", "C" + temp];
                rangesummary110.Borders.Color = 0;
                rangesummary110.Borders.Weight = 2;
                //居中
                rangesummary110.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangesummary110.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;


                //设置属性标签
                xlChart.SetElement(MsoChartElementType.msoElementDataLabelOutSideEnd);  //数据标签
                xlChart.SetElement(MsoChartElementType.msoElementLegendBottom);  //设为底部显示
                xlChart.SetElement(MsoChartElementType.msoElementChartTitleAboveChart);  //设置标题
                objsheet.Range["F:F"].ColumnWidth = 20.5;   //设置宽度
                xlChart.ChartTitle.Text = Title + "(" + DateBegin + "-" + DateEnd + ")";

                objWorkbook.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, objsheet.Name); // xlLocationAsObject,将图表嵌入到现有工作表中。

                objsheet.Shapes.Item("Chart 1").Top = 100;  //调图表的位置上边距
                objsheet.Shapes.Item("Chart 1").Left = 400;
                objsheet.Shapes.Item("Chart 1").Width = 700;   //调图表的宽度
                objsheet.Shapes.Item("Chart 1").Height = 340;  //调图表的高度

                //保存
                objsheet.Name = Title + intExcelTempIndex.ToString();
                objsheet.Tab.Color = Excel.XlThemeColor.xlThemeColorLight1;
                fileNameString = System.Windows.Forms.Application.StartupPath + "\\" + objsheet.Name + ".xlsx";

                objExcel.DisplayAlerts = false;
                objExcel.AlertBeforeOverwriting = false;
                //保存文件   
                objWorkbook.SaveAs(fileNameString, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value);


            }
            catch (Exception ex)
            {
                MessageBox.Show("该文件已打开，请关闭后重试！", "警告 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Open = false;
                return;
            }
            finally
            {
                //确认进程关闭
                if (objWorkbook != null) objWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
                if (objExcel.Workbooks != null) objExcel.Workbooks.Close();
                if (objExcel != null) objExcel.Quit();
                // 安全回收进程
                var aa = System.GC.GetGeneration(objExcel);

                objsheet = null;
                objWorkbook = null;
                objExcel = null;

                if (Open == true)
                {
                    Process.Start(fileNameString);
                }
            }

        }

        /// <summary>
        /// 柱形图
        /// </summary>
        /// <param name="txDataColumnc"></param>
        /// <param name="tyProductColumnc"></param>
        /// <param name="tyNumberColumnc"></param>
        public void ExportRectangle(List<string> txDataColumnc, List<string> tyProductColumnc, List<int> tyNumberColumnc, string DateBegin, string DateEnd)
        {
            Open = true;
            intExcelTempIndex++;
            List<string> txDataColumn = txDataColumnc;
            List<string> tyDataOk = tyProductColumnc;
            List<int> tyDataNo = tyNumberColumnc;

            Excel.Application objExcel = null;//创建一个excel的实例
            Excel.Workbook objWorkbook = null;
            Excel.Worksheet objsheet = null;

            string fileNameString = "";

            try
            {
                //申明对象   
                objExcel = new Excel.Application();
                objWorkbook = objExcel.Workbooks.Add(Missing.Value);
                objsheet = (Excel.Worksheet)objWorkbook.ActiveSheet;

                //设置属性标签
                objsheet.Range["A:A"].ColumnWidth = 20;   //设置宽度
                objsheet.Range["A:A"].NumberFormatLocal = "@";
                objsheet.Range["B:B"].ColumnWidth = 20;   //设置宽度
                objsheet.Range["C:C"].ColumnWidth = 20;   //设置宽度

                #region 管理人员
                int col = 1;
                objExcel.Cells[1, col] = "日期";
                objExcel.Cells[1, col + 1] = "产品名称";
                objExcel.Cells[1, col + 2] = "数量";
                int row = 2;    //row和 i得对应关系是row = i+2 ; i = row -2
                int temp = row;
                int cell = 0;
                for (int i = 0; i < txDataColumn.Count; i++)
                {
                    objExcel.Cells[row, col] = txDataColumn[i].ToString();
                    row++;
                }

                for (int i = 0; i < txDataColumn.Count; i++)
                {
                    if (i == 0)
                    {
                        //objExcel.Cells[row, col] = txDataColumn[i];
                    }
                    else
                    {
                        if (txDataColumn[i] != txDataColumn[i - 1])
                        {
                            cell = i - 1 + 2;
                            Excel.Range rangeChange = objsheet.Range["A" + temp, "A" + cell];
                            rangeChange.Value2 = Type.Missing;
                            rangeChange.Merge(Type.Missing);
                            rangeChange.Value2 = txDataColumn[i - 1].ToString();

                            temp = i + 2;
                        }
                    }
                    row++;
                }

                row = 2;
                for (int i = 0; i < tyDataOk.Count; i++)
                {
                    objExcel.Cells[row, col + 1] = tyDataOk[i];
                    row++;
                }
                row = 2;
                for (int i = 0; i < tyDataNo.Count; i++)
                {
                    objExcel.Cells[row, col + 2] = tyDataNo[i];
                    row++;
                }
                #endregion

                int num = txDataColumn.Count + 1;
                Excel.Range rangeAll = objsheet.Range["A1", "C" + num];
                rangeAll.Borders.Color = 0;
                rangeAll.Borders.Weight = 2;
                rangeAll.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangeAll.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;



                //柱状图
                Excel.Chart xlChart2 = (Excel.Chart)objWorkbook.Charts.Add(Type.Missing, objsheet, Type.Missing, Type.Missing);
                Excel.Range cellRange = objsheet.get_Range((Excel.Range)objsheet.Cells[1, 1], (Excel.Range)objsheet.Cells[1 + txDataColumn.Count, 3]);
                //1-cellRange:数据源的范围，2-图表类型，3-Type.Missing，4-在图表上将列或行用作数据系列的方式，5、6-第五个第六个参数设置图表的x轴和y轴分别是数据源的哪些列/行，7-图表是否有图例，8、9、10-设置标题
                xlChart2.ChartWizard(cellRange,
                                Excel.XlChartType.xlColumnClustered, //2-图表类型
                                Type.Missing,//内置自动套用格式的选项编号。 可为从 1 到 10 的数字，其取值依赖于库的类型。 如果省略此参数, 则 Excel 根据库的类型和数据源选择默认值。
                                Excel.XlRowCol.xlColumns, //在图表上将列或行用作数据系列的方式
                                2, //第五个第六个参数设置图表的x轴和y轴分别是数据源的哪些列/行--这个2代表数据源的x轴由两个参数确认；可以不写，默认的就很难看
                                1,//--这个2代表数据源的x轴由1个参数确认；可以不写，默认的就很难看
                                true, //图表是否有图例
                                "每日总量统计(" + DateBegin + "-" + DateEnd + ")", //以下都是标题
                                Type.Missing,
                                Type.Missing,
                                "");

                xlChart2.ChartStyle = 201;
                xlChart2.SetElement(MsoChartElementType.msoElementDataLabelOutSideEnd); // 设置图表上图表元素。 为可读/写属性。  
                xlChart2.Location(Excel.XlChartLocation.xlLocationAsObject, objsheet.Name);//将图表移动到新位置。

                objWorkbook.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, objsheet.Name); // xlLocationAsObject,将图表嵌入到现有工作表中。
                objsheet.Shapes.Item("Chart 1").Top = 100;  //调图表的位置上边距
                objsheet.Shapes.Item("Chart 1").Left = 400;
                objsheet.Shapes.Item("Chart 1").Width = objsheet.Shapes.Item("Chart 1").Width = txDataColumn.Count * 100 > 400 ? txDataColumn.Count * 30 : 400;   //调图表的宽度
                objsheet.Shapes.Item("Chart 1").Height = 300;  //调图表的高度

                //保存
                objsheet.Name = "每日总量统计" + intExcelTempIndex.ToString();
                objsheet.Tab.Color = Excel.XlThemeColor.xlThemeColorLight1;
                fileNameString = Application.StartupPath + "\\" + objsheet.Name + ".xlsx";

                objExcel.DisplayAlerts = false;
                objExcel.AlertBeforeOverwriting = false;
                //保存文件   
                objWorkbook.SaveAs(fileNameString, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Missing.Value);



            }
            catch (Exception error)
            {

                MessageBox.Show("该文件已打开，请关闭后重试！", "警告 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Open = false;
                return;
            }
            finally
            {
                //确认进程关闭
                if (objWorkbook != null) objWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
                if (objExcel.Workbooks != null) objExcel.Workbooks.Close();
                if (objExcel != null) objExcel.Quit();
                // 安全回收进程
                var aa = System.GC.GetGeneration(objExcel);

                objsheet = null;
                objWorkbook = null;
                objExcel = null;

                if (Open == true)
                {
                    Process.Start(fileNameString);
                }
            }



        }

        /// <summary>
        /// 详细信息
        /// </summary>
        /// <param name="DataGridView"></param>
        public void ShowDetail(DataGridView DataGridView)
        {
            if (DataGridView.Rows.Count < 1)
            {
                return;
            }
            Open = true;
            intExcelTempIndex++;

            Excel.Application objExcel = null;//创建一个excel的实例
            Excel.Workbook objWorkbook = null;
            Excel.Worksheet objsheet = null;

            string fileNameString = "";

            try
            {
                //申明对象   
                objExcel = new Excel.Application();
                objWorkbook = objExcel.Workbooks.Add(Missing.Value);
                objsheet = (Excel.Worksheet)objWorkbook.ActiveSheet;

                objsheet.Range["A:A"].ColumnWidth = 30;   //设置宽度
                objsheet.Range["A:A"].NumberFormatLocal = "@";
                objsheet.Range["B:B"].ColumnWidth = 20;   //设置宽度
                objsheet.Range["C:C"].ColumnWidth = 20;   //设置宽度
                objsheet.Range["D:D"].ColumnWidth = 20;   //设置宽度


                //出货明细
                objExcel.Cells[1, 1] = "日期";
                objExcel.Cells[1, 2] = "客户";
                objExcel.Cells[1, 3] = "产品";
                objExcel.Cells[1, 4] = "数量";

                ArrayList colList = new ArrayList(); //存放不显示的列
                //表内容
                //col 
                for (int i = 0; i < DataGridView.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        objExcel.Cells[i + 2, j + 1] = DataGridView.Rows[i].Cells[j].Value.ToString();
                    }
                }

                int num = DataGridView.Rows.Count;
                Excel.Range rangeAll = objsheet.Range["A1", "D" + num];
                rangeAll.Borders.Color = 0;
                rangeAll.Borders.Weight = 2;
                rangeAll.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangeAll.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;


                objsheet.Name = "详细信息" + intExcelTempIndex.ToString();
                objsheet.Tab.Color = Excel.XlThemeColor.xlThemeColorLight1;
                fileNameString = Application.StartupPath + "\\" + objsheet.Name + ".xlsx";


                objExcel.DisplayAlerts = false; //默认值为 True 。 将此属性设置为False ，如果您不想被应用程序打扰提示和通知消息时
                objExcel.AlertBeforeOverwriting = false; //如果为true之前改写非空单元格拖放编辑操作过程中，Microsoft Excel 显示一条消息。
                objWorkbook.SaveAs(fileNameString, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                         Missing.Value, Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value,
                         Missing.Value, Missing.Value);


            }
            catch (Exception error)
            {
                MessageBox.Show("该文件已打开，请关闭后重试！", "警告 ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Open = false;
                return;
            }
            finally
            {

                //关闭Excel应用   
                if (objWorkbook != null) objWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
                if (objExcel.Workbooks != null) objExcel.Workbooks.Close();
                if (objExcel != null) objExcel.Quit();

                // 安全回收进程
                System.GC.GetGeneration(objExcel);

                objsheet = null;
                objWorkbook = null;
                objExcel = null;

                if (Open == true)
                {
                    Process.Start(fileNameString);

                }
            }
        }

    }

}
