using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BaseLayer
{
    public class ExcelHelper2
    {




        public void Test()
        {


            Excel.Application app = null;//创建一个excel的实例
            Excel.Workbook wbk = null;

            Excel.Worksheet wks1 = null;
            Excel.Worksheet wks2 = null;
            Excel.Worksheet wks3 = null;
            Excel.Worksheet wks4 = null;




           
            

           

            //申明对象   
            app = new Excel.Application();
            wbk = app.Workbooks.Add(Missing.Value);
            app.Visible = true;

            wks1 = (Excel.Worksheet)wbk.Sheets.get_Item(1);
            app.Cells[1, 1] = "客户";
            app.Cells[1, 2] = "数量";
            app.Cells[1, 3] = "占比";


            wks2 = (Excel.Worksheet)wbk.Sheets.get_Item("Sheet1");
            wks2 = (Excel.Worksheet)wbk.Worksheets.Add(wks2, Type.Missing, Type.Missing, Type.Missing);

           

            wks2.Range["A:A"].ColumnWidth = 30;   //设置宽度
            wks2.Range["B:B"].ColumnWidth = 20;   //设置宽度
            wks2.Range["C:C"].ColumnWidth = 20;   //设置宽度

            //合格率
            app.Cells[1, 1] = "客户";
            app.Cells[1, 2] = "数量";
            app.Cells[1, 3] = "占比";

            wks3 = (Excel.Worksheet)wbk.Sheets.get_Item("Sheet2");
            wks3 = (Excel.Worksheet)wbk.Worksheets.Add(wks3, Type.Missing, Type.Missing, Type.Missing);
            app.Cells[1, 1] = "客户";
            app.Cells[1, 2] = "数量";
            app.Cells[1, 3] = "占比";


            wks4 = (Excel.Worksheet)wbk.Sheets.get_Item("Sheet3");
            wks4 = (Excel.Worksheet)wbk.Worksheets.Add(wks3, Type.Missing, Type.Missing, Type.Missing);

            app.Cells[1, 1] = "客户";
            app.Cells[1, 2] = "数量";
            app.Cells[1, 3] = "占比";
            // 应用程序
            //Application app = new Application();
            //// 工作簿
            ////Workbook wbk = app.Workbooks.Open(tbFilePath.Text);
            ////或
            //Workbooks wbks = app.Workbooks;
            //Workbook wbk = wbks.Add(Missing.Value);
            ////工作表
            //Worksheet wsh = wbk.Sheets["All"];
            //或
            //Sheets shs = wbk.Sheets;
            //Worksheet wsh = (Worksheet)shs.get_Item(1);


            //申明保存对话框   
            SaveFileDialog dlg = new SaveFileDialog();
            //默然文件后缀   
            dlg.DefaultExt = "xlsx ";
            //文件后缀列表   
            dlg.Filter = "EXCEL文件(*.XLSX)|*.xlsx ";
            //默然路径是系统当前路径   
            dlg.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            //打开保存对话框   
            if (dlg.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            //返回文件路径   
            string fileNameString = dlg.FileName;
            //验证strFileName是否为空或值无效   
            if (fileNameString.Trim() == " ")
            {
                return;
            }





            //应用程序
            //Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            ////工作簿
            //Workbook wbk = app.Workbooks.Open(tbFilePath.Text);
            ////工作表
            //Worksheet wsh = wbk.Sheets["All"];

            //dlg.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            //fileNameString = System.Windows.Forms.Application.StartupPath + "\\" + objsheet.Name + ".xlsx";

            //读取
            //string str = wks1.Cells[1, 1].Value.ToString();

            //写入，索引以1开始
            wks1.Cells[2, 1] = "str";
            //wks2.Cells[2, 1] = "str";

            wks1.Name = "每日占用";
            //wks2.Name = "每日占用";
            //保存
            wbk.SaveAs(fileNameString);

            //退出
            app.Quit();
            //释放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);



        }
    }
}
