using System;
using System.IO;
using NPOI.SS.UserModel;

namespace K2.NPOIExtensions
{
    /// <summary>
    /// 用于从Excel工作簿中读取数据
    /// </summary>
    public class WorkbookReader
    {
        readonly IWorkbook workBook;

        public WorkbookReader(IWorkbook workBook)
        {
            this.workBook = workBook ?? throw new ArgumentNullException("workBook");
        }

        public WorkbookReader(Stream inputStream)
        {
            if (inputStream == null) throw new ArgumentNullException("inputStream");

            this.workBook = WorkbookFactory.Create(inputStream);
        }

        public SheetReader CreateSheetReader(string sheetName)
        {
            ISheet sheet = workBook.GetSheet(sheetName);

            if (sheet == null)
                return null;

            return new SheetReader(sheet);
        }

        public SheetReader CreateSheetReader(int sheetIndex)
        {
            var sheet = workBook.GetSheetAt(sheetIndex);

            if (sheet == null)
                return null;

            return new SheetReader(sheet);
        }
    }
}
