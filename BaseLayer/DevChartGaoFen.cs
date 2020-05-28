using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraCharts;

namespace BaseLayer
{
    class DevChartGaoFen
    {
        private void ChartBinding(object sender)
        {
            string Str_SQL = "select top 7 UnitPrice,UnitsInStock,ReorderLevel,ProductID from Products order by ProductID ";
            DataSet DS = SqlHelper.ExecuteDataSet(Str_SQL);
            //定义线型，名称
            Series S1 = new Series("线条图测试", ViewType.Line);
            ChartControl chart = new ChartControl();
            //定义X轴的数据的类型。质量，数字，时间
            S1.ArgumentScaleType = ScaleType.Numerical;
            //定义线条上点的标识形状
            ((LineSeriesView)S1.View).LineMarkerOptions.Kind = MarkerKind.Circle;
            //线条的类型，虚线，实线
            ((LineSeriesView)S1.View).LineStyle.DashStyle = DashStyle.Solid;
            //S1绑定数据源
            S1.DataSource = DS.Tables[0].DefaultView;
            //S1的X轴数据源字段
            S1.ArgumentDataMember = "UnitPrice";
            //S2的Y轴数据源字段
            S1.ValueDataMembers[0] = "UnitsInStock";
            //柱状图演示
            Series S2 = new Series("柱状图测试", ViewType.Bar);
            S2.ArgumentScaleType = ScaleType.Numerical;
            S2.DataSource = DS.Tables[0].DefaultView;
            S2.ArgumentDataMember = "UnitPrice";
            S2.ValueDataMembers[0] = "ReorderLevel";
            //光滑线条演示
            Series S3 = new Series("弧度曲线测试", ViewType.Spline);
            S3.ArgumentScaleType = ScaleType.Numerical;
            S3.DataSource = DS.Tables[0].DefaultView;
            S3.ArgumentDataMember = "UnitPrice";
            S3.ValueDataMembers[0] = "UnitsInStock";
            //加入chartcontrol
           
            //定义chart标题
            ChartTitle CT1 = new ChartTitle();
            CT1.Text = "这是第一个DEMO";
            ChartTitle CT2 = new ChartTitle();


            CT2.Text = "CopyRight By BJYD";
            CT2.TextColor = System.Drawing.Color.Black;
            //标题对齐方式
            CT2.Dock = ChartTitleDockStyle.Bottom;
           
            CT2.Indent = 1;
            //坐标标题的定义
            //坐标值说明的字体尺寸，颜色定义
            
            //图例的位置定义
           
        }
    }
}
