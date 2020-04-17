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
            //Ĭ�����ñ�������ʽ
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
            //������������ʽ��ӳ���ϵ
            //ExcelExtension.ConfigurationCellStyle(new Dictionary<Type, Action<ICellStyle>>
            //{
            //    { typeof(decimal), (cellStyle) => { cellStyle.Alignment = HorizontalAlignment.Right; /*����decimal���͵�������ʽ����*/ } }
            //});
            ////����������θ�ʽ��Ϊ�ַ���
            //ExcelExtension.ConfigurationCellValueFormat(new Dictionary<Type, Func<object, string>>
            //{
            //    { typeof(decimal), (value) => { return ((decimal)value).ToString("0.0000"); /*����decimal���͵����ݸ�ʽ��Ϊ����4λС��*/ } }
            //});
            //Ĭ�����ñ�������ʽ
            //ExcelExtension.ConfigurationTitleCellStyle((cellstyle) =>
            //{
            //    cellstyle.FillForegroundColor = 57;
            //    cellstyle.FillPattern = FillPattern.AltBars;
            //    cellstyle.Alignment = HorizontalAlignment.Center;
            //});
            //var excelTable = new ExcelTable("��Ʒ����", new CellStyle
            //{
            //    //������θ�ʽ�����������
            //    Format = (value) =>
            //    {
            //        return ((decimal)value).ToString("0.0000");
            //    },
            //    //��������ʽ, ����ʽ��Ӧ���ڳ��˱�����֮�������
            //    Style = (cellStyle) =>
            //    {
            //        cellStyle.Alignment = HorizontalAlignment.Right;
            //    },
            //    //�б���
            //    TitleValue = "����"
            //});
            //excelTable.TitleStyleAction = (style) =>
            //{
            //    //ˮƽ����
            //    style.Alignment = HorizontalAlignment.Center;
            //    //����ǰ��ɫΪ��ɫ
            //    style.FillForegroundColor = HSSFColor.Green.Index;
            //    //�����ʽ
            //    style.FillPattern = FillPattern.SolidForeground;
            //};
            //excelTable.Add("�ֱ�", 5.3M);

            //var wb = ExcelExtension.CreateWorkbook().AddData(excelTable);
            //var titleRow = wb.GetSheetAt(0).GetRow(0);
            //var titleOneCellStyle = titleRow.GetCell(0).CellStyle;

        }
    }
}
