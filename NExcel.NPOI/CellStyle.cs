using NPOI.SS.UserModel;
using System;

namespace Colipu.Extensions.Excel
{
    /// <summary>
    /// 列样式
    /// </summary>
    public class CellStyle
    {
        /// <summary>
        /// 列标题
        /// </summary>
        public string TitleValue { get; set; }

        /// <summary>
        /// 自定义列样式
        /// </summary>
        public Action<ICellStyle> Style { get; set; }

        public Func<object, string> Format { get; set; }
    }
}
