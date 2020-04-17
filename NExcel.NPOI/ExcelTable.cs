using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Colipu.Extensions.Excel
{
    public class ExcelTable
    {
        private IList<object[]> _dataRows;
        private IList<CellStyle> _cellStyles;
        public Action<ICellStyle> TitleStyleAction { get; set; }
        public bool HidentTitle { get; }
        public ExcelTable(params object[] titleRows) : this(hidentTitle: false, titleRows)
        {
        }
        public ExcelTable(IEnumerable<object> titleRows) : this(hidentTitle: false, cellStyles: titleRows.ToArray())
        {
        }
        public ExcelTable(bool hidentTitle, IEnumerable<object> titleRows) : this(hidentTitle: hidentTitle, cellStyles: titleRows.ToArray())
        {
        }

        public ExcelTable(bool hidentTitle, params object[] cellStyles)
        {
            if (cellStyles == null || cellStyles.Length == 0)
            {
                throw new ArgumentNullException(nameof(cellStyles));
            }
            _cellStyles = new List<CellStyle>();
            _dataRows = new List<object[]>();
            HidentTitle = hidentTitle;
            foreach (var cellStyle in cellStyles)
            {
                switch (cellStyle)
                {
                    case CellStyle style:
                        _cellStyles.Add(style);
                        break;
                    case string title:
                        _cellStyles.Add(new CellStyle { TitleValue = title });
                        break;
                    case null:
                        _cellStyles.Add(new CellStyle());
                        break;
                    default:
                        throw new ArgumentException($"无效的列样式类型{cellStyle.GetType()}");
                }
            }
        }

        public void Add(params object[] values)
        {
            this._dataRows.Add(values);
        }

        public IList<object[]> GetDataRows()
        {
            return _dataRows;
        }

        public IList<CellStyle> GetCellStyles()
        {
            return _cellStyles;
        }
    }
}
