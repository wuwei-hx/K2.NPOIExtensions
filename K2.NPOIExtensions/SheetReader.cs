using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace K2.NPOIExtensions
{
    /// <summary>
    /// 用于从Excel表格中读取数据
    /// </summary>
    public class SheetReader
    {
        readonly ISheet sheet;

        public SheetReader(ISheet sheet)
        {
            this.sheet = sheet ?? throw new ArgumentNullException("sheet");
        }

        /// <summary>
        /// 读取一行数据，并映射到指定类型
        /// </summary>
        public TModel ReadRow<TModel>(int rowIndex, Func<IRow, TModel> mapRow)
        {
            IRow row = sheet.GetRow(rowIndex);

            return mapRow(row);
        }

        public TModel ReadFirstRow<TModel>(Func<IRow, TModel> mapRow)
        {
            return ReadRow(sheet.FirstRowNum, mapRow);
        }

        /// <summary>
        /// 读取表格指定范围的多行数据，并映射到指定类型
        /// </summary>
        public IEnumerable<TModel> ReadSheet<TModel>(Func<IRow, TModel> mapRow, int startRow = 0, int count = 0)
        {
            var result = new List<TModel>();

            int endRow = Math.Min(sheet.LastRowNum, count == 0 ? int.MaxValue : startRow + count - 1);

            for (int index = startRow; index <= endRow; index++)
            {
                IRow row = sheet.GetRow(index);
                TModel model = mapRow(row);

                // 不允许有空行
                if (model == null)
                    break;

                result.Add(model);
            }

            return result;
        }
    }
}
