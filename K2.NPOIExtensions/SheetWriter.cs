using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace K2.NPOIExtensions
{
    /// <summary>
    /// 用于向Excel表格写入数据
    /// </summary>
    public class SheetWriter<TModel>
    {
        readonly ISheet sheet;
        readonly IEnumerable<SheetColumnOption<TModel>> options;
        readonly Dictionary<SheetColumnOption<TModel>, ICellStyle> headerStyles;
        readonly Dictionary<SheetColumnOption<TModel>, ICellStyle> lineStyles;

        public SheetWriter(ISheet sheet, IEnumerable<SheetColumnOption<TModel>> options, Dictionary<SheetColumnOption<TModel>, ICellStyle> headerStyles, Dictionary<SheetColumnOption<TModel>, ICellStyle> lineStyles)
        {
            this.sheet = sheet ?? throw new ArgumentNullException("sheet");
            this.options = options ?? throw new ArgumentNullException("options");
            this.headerStyles = headerStyles ?? throw new ArgumentNullException("headerStyles");
            this.lineStyles = lineStyles ?? throw new ArgumentNullException("lineStyles");
        }

        /// <summary>
        /// 将一组TModel数据写入到Excel表格中
        /// </summary>
        public SheetWriter<TModel> Write(IEnumerable<TModel> source)
        {
            sheet.IsPrintGridlines = true;
            sheet.DisplayGridlines = true;

            var rowIndex = 0;

            IRow row = sheet.CreateRow(rowIndex++);
            WriteHeader(sheet, row);

            foreach (TModel model in source)
            {
                row = sheet.CreateRow(rowIndex++);
                WriteLine(row, model);
            }

            return this;
        }

        void WriteHeader(ISheet sheet, IRow row)
        {
            var columnIndex = 0;

            foreach (SheetColumnOption<TModel> option in options)
            {
                if (option.ColumnWidth > 0)
                    sheet.SetColumnWidth(columnIndex, option.ColumnWidth);

                ICell cell = row.CreateCell(columnIndex);

                cell.SetCellValue(option.ColumnName);

                if (headerStyles[option] != null)
                    cell.CellStyle = headerStyles[option];

                columnIndex++;
            }
        }

        void WriteLine(IRow row, TModel model)
        {
            var columnIndex = 0;
            foreach (var option in options)
            {
                var cell = row.CreateCell(columnIndex);

                option.SetCellValue(cell, model);

                if (lineStyles[option] != null)
                    cell.CellStyle = lineStyles[option];

                columnIndex++;
            }
        }
    }
}
