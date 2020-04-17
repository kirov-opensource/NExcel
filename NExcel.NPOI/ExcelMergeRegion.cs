using System;
using System.Collections.Generic;
using System.Text;

namespace Colipu.Extensions.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelMergeRegion
    {
        /// <summary>
        /// 开始Y坐标
        /// </summary>
        public int FirstRow { get; set; }

        /// <summary>
        /// 开始X坐标
        /// </summary>
        public int FirstCloumn { get; set; }

        /// <summary>
        /// 结束Y坐标
        /// </summary>
        public int LastRow { get; set; }

        /// <summary>
        /// 结束X坐标
        /// </summary>
        public int LastCloumn { get; set; }
    }
}
