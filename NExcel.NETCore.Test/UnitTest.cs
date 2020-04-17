using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using Xunit;

namespace Colipu.Extensions.Excel.NETCore.Test
{
    public class UnitTest
    {
        [Fact]
        public void CreateNoArgumentExcelTable()
        {
            Assert.Throws<ArgumentNullException>(() => { new ExcelTable(); });
        }

        [Fact]
        public void CreateWorkbook()
        {
            var wb = ExcelExtension.CreateWorkbook();
            Assert.IsType<HSSFWorkbook>(wb);
            wb = ExcelExtension.CreateWorkbook(WorkbookStyle.HSSFWorkbook);
            Assert.IsType<HSSFWorkbook>(wb);
            wb = ExcelExtension.CreateWorkbook(WorkbookStyle.XSSFWorkbook);
            Assert.IsType<XSSFWorkbook>(wb);
        }

        [Fact]
        public void DefaultTitleStyleTest()
        {
            //默认设置标题列样式
            ExcelExtension.ConfigurationTitleCellStyle((cellStyle) =>
            {
                cellStyle.FillForegroundColor = HSSFColor.Green.Index;
                cellStyle.FillPattern = FillPattern.SolidForeground;
                cellStyle.Alignment = HorizontalAlignment.Center;
            });
            var excelTable = new ExcelTable("TitleOne", "TitleTwo");
            var wb = ExcelExtension.CreateWorkbook().AddData(excelTable);
            var titleRow = wb.GetSheetAt(0).GetRow(0);
            var titleOneCellStyle = titleRow.GetCell(0).CellStyle;
            Assert.Equal(HSSFColor.Green.Index, titleOneCellStyle.FillForegroundColor);
            Assert.Equal(FillPattern.SolidForeground, titleOneCellStyle.FillPattern);
            Assert.Equal(HorizontalAlignment.Center, titleOneCellStyle.Alignment);
            var titleTwoCellStyle = titleRow.GetCell(1).CellStyle;
            Assert.Equal(HSSFColor.Green.Index, titleTwoCellStyle.FillForegroundColor);
            Assert.Equal(FillPattern.SolidForeground, titleTwoCellStyle.FillPattern);
            Assert.Equal(HorizontalAlignment.Center, titleTwoCellStyle.Alignment);
        }


        [Fact]
        public void DefaultTitleStyleTest1()
        {
            //设置类型与样式的映射关系
            //ExcelExtension.ConfigurationCellStyle(new Dictionary<Type, Action<ICellStyle>>
            //{
            //    { typeof(decimal), (cellStyle) => { cellStyle.Alignment = HorizontalAlignment.Right; /*所有decimal类型的数据样式居右*/ } }
            //});
            ////设置类型如何格式化为字符串
            //ExcelExtension.ConfigurationCellValueFormat(new Dictionary<Type, Func<object, string>>
            //{
            //    { typeof(decimal), (value) => { return ((decimal)value).ToString("0.0000"); /*所有decimal类型的数据格式化为保留4位小数*/ } }
            //});
            //默认设置标题列样式
            //ExcelExtension.ConfigurationTitleCellStyle((cellstyle) =>
            //{
            //    cellstyle.FillForegroundColor = 57;
            //    cellstyle.FillPattern = FillPattern.AltBars;
            //    cellstyle.Alignment = HorizontalAlignment.Center;
            //});
            //var excelTable = new ExcelTable("商品名称", new CellStyle
            //{
            //    //设置如何格式化传入的数据
            //    Format = (value) =>
            //    {
            //        return ((decimal)value).ToString("0.0000");
            //    },
            //    //设置列样式, 此样式会应用于除了标题行之外的整列
            //    Style = (cellStyle) =>
            //    {
            //        cellStyle.Alignment = HorizontalAlignment.Right;
            //    },
            //    //列标题
            //    TitleValue = "单价"
            //});
            //excelTable.TitleStyleAction = (style) =>
            //{
            //    //水平居中
            //    style.Alignment = HorizontalAlignment.Center;
            //    //设置前景色为绿色
            //    style.FillForegroundColor = HSSFColor.Green.Index;
            //    //填充样式
            //    style.FillPattern = FillPattern.SolidForeground;
            //};
            //excelTable.Add("钢笔", 5.3M);

            //var wb = ExcelExtension.CreateWorkbook().AddData(excelTable);
            //var titleRow = wb.GetSheetAt(0).GetRow(0);
            //var titleOneCellStyle = titleRow.GetCell(0).CellStyle;

        }
    }
}
