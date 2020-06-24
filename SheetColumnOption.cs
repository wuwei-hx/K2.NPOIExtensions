using System;
using NPOI.SS.UserModel;

namespace K2.NPOIExtensions
{
    /// <summary>
    /// Excel表格列的配置项
    /// </summary>
    public class SheetColumnOption<TModel>
    {
        public string ColumnName { get; set; }
        public int ColumnWidth { get; set; }
        public Func<IWorkbook, ICellStyle> CreateHeaderStyle { get; set; }
        public Func<IWorkbook, ICellStyle> CreateLineStyle { get; set; }
        public Action<ICell, TModel> SetCellValue { get; set; }
    }
}
