using System;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace K2.NPOIExtensions
{
    /// <summary>
    /// 用于向Excel工作簿写入数据
    /// </summary>
    public class WorkbookWriter
    {
        readonly IWorkbook workBook;

        public WorkbookWriter(IWorkbook workBook)
        {
            this.workBook = workBook ?? throw new ArgumentNullException("workBook");
        }

        /// <summary>
        /// 支持Excel 97-2003 工作簿(xls)
        /// </summary>
        public static WorkbookWriter CreateHSSFWorkbookWriter()
        {
            return new WorkbookWriter(new HSSFWorkbook());
        }

        /// <summary>
        /// 支持Excel 工作簿(xlsx)
        /// </summary>
        public static WorkbookWriter CreateXSSFWorkbookWriter()
        {
            return new WorkbookWriter(new XSSFWorkbook());
        }

        /// <summary>
        /// 将数据写入一个指定名称的表格
        /// </summary>
        public WorkbookWriter WriteSheet<TModel>(string sheetName, SheetWriterBuilder<TModel> builder, IEnumerable<TModel> source)
        {
            ISheet sheet = workBook.GetSheet(sheetName) ?? workBook.CreateSheet(sheetName);
            SheetWriter<TModel> writer = builder.Build(sheet);
            writer.Write(source);

            return this;
        }

        /// <summary>
        /// 将数据写入一个新的表格
        /// </summary>
        public WorkbookWriter WriteSheet<TModel>(SheetWriterBuilder<TModel> builder, IEnumerable<TModel> source)
        {
            var sheetName = $"Sheet{workBook.NumberOfSheets + 1}";

            return WriteSheet(sheetName, builder, source);
        }

        /// <summary>
        /// 导出工作簿
        /// </summary>
        public byte[] ToArray()
        {
            using MemoryStream memoryStream = new MemoryStream();

            workBook.Write(memoryStream);
            return memoryStream.ToArray();
        }
    }
}
