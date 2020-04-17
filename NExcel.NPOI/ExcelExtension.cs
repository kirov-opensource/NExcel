using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace Colipu.Extensions.Excel
{
    public static class ExcelExtension
    {
        private static Action<ICellStyle> _titleCellStyle;
        private static Dictionary<Type, Action<ICellStyle>> _cellStyles;
        private static Dictionary<Type, Func<object, string>> _cellValueFormat;

        /// <summary>
        /// 配置默认标题行样式
        /// </summary>
        /// <param name="titleCellStyles"></param>
        public static void ConfigurationTitleCellStyle(Action<ICellStyle> titleCellStyle)
        {
            _titleCellStyle = titleCellStyle;
        }

        /// <summary>
        /// 配置默认普通行样式
        /// </summary>
        /// <param name="cellStyles"></param>
        public static void ConfigurationCellStyle(Dictionary<Type, Action<ICellStyle>> cellStyles)
        {
            _cellStyles = cellStyles;
        }

        /// <summary>
        /// 配置类型如何填充到单元格
        /// </summary>
        /// <param name="cellValueFormat"></param>
        public static void ConfigurationCellValueFormat(Dictionary<Type, Func<object, string>> cellValueFormat)
        {
            _cellValueFormat = cellValueFormat;
        }

        /// <summary>
        /// 向工作簿填充数据
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="excelTable"></param>
        /// <returns></returns>
        public static IWorkbook AddData(this IWorkbook wb, ExcelTable excelTable)
        {
            return wb.AddData(excelTable, 0, false);
        }

        /// <summary>
        /// 从指定起始行开始填充数据
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="excelTable"></param>
        /// <param name="startRowIndex"></param>
        /// <returns></returns>
        public static IWorkbook AddData(this IWorkbook wb, ExcelTable excelTable, int startRowIndex)
        {
            return wb.AddData(excelTable, startRowIndex, false);
        }

        /// <summary>
        /// 向工作簿填充数据
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="rows"></param>
        /// <param name="startRowIndex">填充起始行</param>
        /// <param name="ignoreDefaultConfiguration">忽略默认配置</param>
        /// <returns></returns>
        public static IWorkbook AddData(this IWorkbook wb, ExcelTable excelTable, int startRowIndex, bool ignoreDefaultConfiguration)
        {
            var sheet = wb.GetSheetAt(0);
            //所有列样式
            var cellStyles = excelTable.GetCellStyles();
            //默认标题样式
            ICellStyle defaultTitleStyle = null;
            if (_titleCellStyle != null)
            {
                defaultTitleStyle = wb.CreateCellStyle();
                _titleCellStyle.Invoke(defaultTitleStyle);
            }
            //默认列样式
            var defaultCellStyles = GenerateCellStyle(wb, _cellStyles);
            //列的数量
            var columnLenght = cellStyles.Count();
            var dataRows = excelTable.GetDataRows();
            //将插入行之下的行全部下移
            if (sheet.LastRowNum >= startRowIndex) { sheet.ShiftRows(startRowIndex, sheet.LastRowNum, dataRows.Count(), true, false); }
            var sheetRowIndex = startRowIndex;
            //展示标题行
            if (!excelTable.HidentTitle)
            {
                //从sheet创建行
                var sheetRow = sheet.CreateRow(sheetRowIndex);
                ICellStyle customTitleStyle = null;
                if (excelTable.TitleStyleAction != null)
                {
                    customTitleStyle = wb.CreateCellStyle();
                    excelTable.TitleStyleAction?.Invoke(customTitleStyle);
                }
                for (var columnIndex = 0; columnIndex < columnLenght; columnIndex++)
                {
                    //从当前行创建列
                    var sheetCell = sheetRow.CreateCell(columnIndex);
                    //自定义标题样式
                    if (customTitleStyle != null)
                    {
                        sheetCell.CellStyle = customTitleStyle;
                    }
                    //默认标题样式
                    else if (!ignoreDefaultConfiguration && defaultTitleStyle != null)
                    {
                        //默认标题样式
                        sheetCell.CellStyle = defaultTitleStyle;
                    }
                    var cellStyle = cellStyles[columnIndex];
                    if (cellStyle?.TitleValue == null) { continue; }
                    //设置标题
                    sheetCell.SetCellValue(cellStyle.TitleValue);
                    //自动宽度
                    SetColumnWidth(sheet, columnIndex, cellStyle.TitleValue);
                }
                sheetRowIndex++;
            }


            var customCellStyles = cellStyles.Select(cellStyle =>
            {
                if (cellStyle?.Style == null) { return null; }
                var style = wb.CreateCellStyle();
                cellStyle.Style.Invoke(style);
                return style;
            }).ToArray();

            foreach (var dataRow in dataRows)
            {
                //从sheet创建行
                var sheetRow = sheet.CreateRow(sheetRowIndex);
                if (dataRow == null) { continue; }



                for (var columnIndex = 0; columnIndex < columnLenght; columnIndex++)
                {
                    //从当前行创建列
                    var sheetCell = sheetRow.CreateCell(columnIndex);
                    var cellStyle = cellStyles[columnIndex];
                    var customCellStyle = customCellStyles[columnIndex];
                    var cellValue = dataRow[columnIndex];
                    //自定义样式
                    if (customCellStyle != null)
                    {
                        sheetCell.CellStyle = customCellStyle;
                    }
                    //自定义格式化
                    if (cellStyle.Format != null)
                    {
                        var value = cellStyle.Format.Invoke(cellValue);
                        sheetCell.SetCellValue(value);
                        //自动宽度
                        SetColumnWidth(sheet, columnIndex, value);
                    }
                    //值为空的话无法匹配默认样式与格式化
                    if (cellValue == null)
                    {
                        continue;
                    }
                    //默认样式
                    if (!ignoreDefaultConfiguration && customCellStyle == null && defaultCellStyles != null && defaultCellStyles.TryGetValue(cellValue.GetType(), out var defaultCellStyle))
                    {
                        //默认样式
                        sheetCell.CellStyle = defaultCellStyle;
                    }
                    if (cellStyle.Format != null) { continue; }
                    //默认格式化
                    if (!ignoreDefaultConfiguration && _cellValueFormat != null && _cellValueFormat.TryGetValue(cellValue.GetType(), out var cellFormatFunc))
                    {
                        var value = cellFormatFunc.Invoke(cellValue);
                        //默认类型格式化
                        sheetCell.SetCellValue(value);
                        //自动宽度
                        SetColumnWidth(sheet, columnIndex, value);
                        continue;
                    }
                    //无自定义格式化也无默认格式化 设置值
                    sheetCell.SetCellValue(cellValue.ToString());
                    //自动宽度
                    SetColumnWidth(sheet, columnIndex, cellValue.ToString());

                }
                sheetRowIndex++;
            }
            return wb;
        }

        /// <summary>
        /// 创建新的97-2003 .xls工作簿
        /// </summary>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook()
        {
            return CreateWorkbook(stream: null, WorkbookStyle.HSSFWorkbook);
        }

        /// <summary>
        /// 创建新的工作簿
        /// </summary>
        /// <param name="workbookStyle"></param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(WorkbookStyle workbookStyle)
        {
            return CreateWorkbook(stream: null, workbookStyle);
        }

        /// <summary>
        /// 从Stream创建97-2003 .xls工作簿,如果流为空,则新建工作簿
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(Stream stream)
        {
            return CreateWorkbook(stream, WorkbookStyle.HSSFWorkbook);
        }

        /// <summary>
        /// 从本地路径创建97-2003 .xls工作簿
        /// </summary>
        /// <param name="path">本地路径</param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                return CreateWorkbook(stream, WorkbookStyle.HSSFWorkbook);
            }
        }

        /// <summary>
        /// 从Stream创建工作簿,如果流为空,则新建工作簿
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="workbookStyle">指定创建工作簿的类型</param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(Stream stream, WorkbookStyle workbookStyle)
        {
            //创建excel
            IWorkbook wb = workbookStyle == WorkbookStyle.HSSFWorkbook ? (IWorkbook)(stream == null ? new HSSFWorkbook() : new HSSFWorkbook(stream)) : (IWorkbook)(stream == null ? new XSSFWorkbook() : new XSSFWorkbook(stream));
            try
            {
                wb.GetSheetAt(0);
            }
            catch (Exception)
            {

                wb.CreateSheet();
            }
            return wb;
        }

        /// <summary>
        /// 从本地路径创建工作簿
        /// </summary>
        /// <param name="path">本地路径</param>
        /// <param name="workbookStyle">工作簿类型</param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(string path, WorkbookStyle workbookStyle)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                return CreateWorkbook(stream, workbookStyle);
            }
        }

        /// <summary>
        /// 获取单元格的合并数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static CellRangeAddress GetCellRangeAddress(this ISheet sheet, int rowIndex, int columnIndex)
        {
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var mergedRegions = sheet.GetMergedRegion(i);
                if (rowIndex >= mergedRegions.FirstRow && rowIndex <= mergedRegions.LastRow && columnIndex >= mergedRegions.FirstColumn && columnIndex <= mergedRegions.LastColumn)
                {
                    return mergedRegions;
                }
            }
            return null;
        }

        /// <summary>
        /// 批量合并单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="excelMergeRegionModel"></param>
        public static IWorkbook AddMergedRegion(this IWorkbook wb, List<ExcelMergeRegion> excelMergeRegionModel)
        {
            var sheet = wb.GetSheetAt(0);
            if (excelMergeRegionModel == null)
            {
                return wb;
            }
            //合并单元格
            excelMergeRegionModel.ForEach(item =>
            {
                sheet.AddMergedRegion(new CellRangeAddress(item.FirstRow, item.LastRow, item.FirstCloumn, item.LastCloumn));
            });
            return wb;
        }

        /// <summary>
        /// Excel转换Byte
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static byte[] ToBytes(this IWorkbook wb)
        {
            var stream = new MemoryStream();
            switch (wb)
            {
                case XSSFWorkbook xssf:
                    xssf.Write(stream, true);
                    break;
                default:
                    wb.Write(stream);
                    break;
            }
            var bytes = new byte[stream.Length];
            // 设置当前流的位置为流的开始
            stream.Seek(0, SeekOrigin.Begin);
            stream.Read(bytes, 0, bytes.Length);
            return bytes;
        }

        /// <summary>
        /// 向Excel指定坐标填充数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="positions"></param>
        public static IWorkbook PositioningPadding(this IWorkbook wb, List<PositioningPadding> positions)
        {
            var sheet = wb.GetSheetAt(0);
            foreach (var position in positions)
            {
                var point = GetPosition(position.Position);
                sheet.GetRow(point.Y).GetCell(point.X).SetCellValue(position.Value);
            }
            return wb;
        }

        /// <summary>
        /// 将字母数字坐标转换为X,Y坐标
        /// </summary>
        /// <param name="position">
        /// Excel Cell Position
        /// A1,B2,C3
        /// </param>
        /// <returns></returns>
        private static Point GetPosition(string position)
        {
            var letters = new StringBuilder();
            var number = new StringBuilder();
            foreach (var c in position)
            {
                var ascii = (int)(char.ToUpper(c));
                //字母
                if (ascii >= 65 && ascii <= 90)
                {
                    letters.Append(char.ToUpper(c));
                }
                //数字
                if (ascii >= 48 && ascii <= 57)
                {
                    number.Append(c);
                }
            }

            if (number.Length == 0 || letters.Length == 0)
            {
                return new Point { X = 0, Y = 0 };
            }

            var point = new Point();
            var currentIndex = 0;

            // 公式 AA即为 1*26¹+1  AAA即为 1*26²+1*26¹+1 以此类推
            foreach (var c in letters.ToString())
            {
                currentIndex++;
                //获取到ASCII码 减去64
                var x = c - 64;
                //算次方
                var pow = Math.Pow(26D, double.Parse((letters.Length - currentIndex).ToString()));
                point.X += (int)(x * (pow == 0 ? 1 : pow));
            }
            point.Y = int.Parse(number.ToString());
            point.X--;
            point.Y--;
            return point;
        }

        /// <summary>
        /// 将excel流转换为二维数组
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static List<List<string>> ReadExcelToArray(this Stream stream)
        {
            //读取一个xls格式的excel
            var wb = new HSSFWorkbook(stream);
            //读取第一个sheet
            var sheet = wb.GetSheetAt(0);
            //总共有多少列
            var columnLength = sheet.GetRow(0).LastCellNum;
            //定义集合
            var result = new List<List<string>>();
            //循环行
            for (var rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                //忽略标题行
                if (rowIndex == 0)
                {
                    continue;
                }
                //获取到当前行数据
                var currentRow = sheet.GetRow(rowIndex);
                //定义当前行的集合
                var currentRowList = new List<string>();
                //循环列
                for (var columnIndex = 0; columnIndex < columnLength; columnIndex++)
                {
                    var cell = currentRow.GetCell(columnIndex);
                    if (cell == null)
                    {
                        currentRowList.Add(null);
                        continue;
                    }
                    //先设置Excel格式为String
                    cell.SetCellType(CellType.String);
                    currentRowList.Add(cell.StringCellValue);
                }
                //讲当前行添加到整个集合中
                result.Add(currentRowList);
            }
            return result;
        }

        private static Dictionary<Type, ICellStyle> GenerateCellStyle(IWorkbook wb, Dictionary<Type, Action<ICellStyle>> cellStyleActions)
        {
            if (cellStyleActions == null) { return new Dictionary<Type, ICellStyle>(); }
            var cellStyles = new Dictionary<Type, ICellStyle>();
            foreach (var cellStyleAction in cellStyleActions)
            {
                var style = wb.CreateCellStyle();
                cellStyleAction.Value.Invoke(style);
                cellStyles.Add(cellStyleAction.Key, style);
            }
            return cellStyles;
        }

        private static void SetColumnWidth(ISheet sheet, int columnIndex, string value)
        {
            var width = (value.Length + 6) * 256;
            if (sheet.GetColumnWidth(columnIndex) < width)
            {
                sheet.SetColumnWidth(columnIndex, width);
            }
        }

    }
}
