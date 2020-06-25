using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using NPOI.SS.UserModel;

namespace K2.NPOIExtensions
{
    /// <summary>
    /// 用于生成SheetWriter类
    /// </summary>
    public class SheetWriterBuilder<TModel>
    {
        readonly List<SheetColumnOption<TModel>> options;
        Func<IWorkbook, ICellStyle> createDefaultHeaderStyle;
        Func<IWorkbook, ICellStyle> createDefaultLineStyle;

        SheetColumnOption<TModel> currentColumnOption;

        public SheetWriterBuilder()
        {
            options = new List<SheetColumnOption<TModel>>();

            createDefaultHeaderStyle = null;
            createDefaultLineStyle = null;
            currentColumnOption = null;
        }

        public SheetWriterBuilder<TModel> Column(string name)
        {
            currentColumnOption = new SheetColumnOption<TModel>
            {
                ColumnName = name
            };

            options.Add(currentColumnOption);
            return this;
        }

        // 根据成员表达式，获取模型指定Property的DisplayName属性
        string GetDisplayName<TValue>(Expression<Func<TModel, TValue>> expression)
        {
            Type type = typeof(TModel);

            MemberExpression memberExpression = (MemberExpression)expression.Body;
            string propertyName = ((memberExpression.Member is PropertyInfo) ? memberExpression.Member.Name : null);

            // First look into attributes on a type and it's parents
            DisplayNameAttribute attr;
            attr = (DisplayNameAttribute)type.GetProperty(propertyName).GetCustomAttributes(typeof(DisplayNameAttribute), true).SingleOrDefault();

            // Look for [MetadataType] attribute in type hierarchy
            // http://stackoverflow.com/questions/1910532/attribute-isdefined-doesnt-see-attributes-applied-with-metadatatype-class
            if (attr == null)
            {
                MetadataTypeAttribute metadataType = (MetadataTypeAttribute)type.GetCustomAttributes(typeof(MetadataTypeAttribute), true).FirstOrDefault();
                if (metadataType != null)
                {
                    var property = metadataType.MetadataClassType.GetProperty(propertyName);
                    if (property != null)
                    {
                        attr = (DisplayNameAttribute)property.GetCustomAttributes(typeof(DisplayNameAttribute), true).SingleOrDefault();
                    }
                }
            }
            return (attr != null) ? attr.DisplayName : propertyName; // String.Empty;
        }

        public SheetWriterBuilder<TModel> ColumnFor<TProperty>(Expression<Func<TModel, TProperty>> expression)
        {
            var name = GetDisplayName(expression);
            return Column(name);
        }

        public SheetWriterBuilder<TModel> ColumnWidth(int width)
        {
            currentColumnOption.ColumnWidth = width;
            return this;
        }

        public SheetWriterBuilder<TModel> HeaderStyle(Func<IWorkbook, ICellStyle> createStyle)
        {
            currentColumnOption.CreateHeaderStyle = createStyle;
            return this;
        }

        public SheetWriterBuilder<TModel> LineStyle(Func<IWorkbook, ICellStyle> createStyle)
        {
            currentColumnOption.CreateLineStyle = createStyle;
            return this;
        }

        public SheetWriterBuilder<TModel> DefaultHeaderStyle(Func<IWorkbook, ICellStyle> createStyle)
        {
            createDefaultHeaderStyle = createStyle;
            return this;
        }

        public SheetWriterBuilder<TModel> DefaultLineStyle(Func<IWorkbook, ICellStyle> createStyle)
        {
            createDefaultLineStyle = createStyle;
            return this;
        }

        public SheetWriterBuilder<TModel> CellValue(Action<ICell, TModel> setCellValue)
        {
            currentColumnOption.SetCellValue = setCellValue;
            return this;
        }

        public SheetWriter<TModel> Build(ISheet sheet)
        {
            if (sheet == null)
                return null;

            var workBook = sheet.Workbook;
            var headerStyles = new Dictionary<SheetColumnOption<TModel>, ICellStyle>();
            var lineStyles = new Dictionary<SheetColumnOption<TModel>, ICellStyle>();
            Func<IWorkbook, ICellStyle> createEmptyStyle = (workbook) => null;

            foreach (var option in options)
            {
                var createHeaderStyle = option.CreateHeaderStyle ?? createDefaultHeaderStyle ?? createEmptyStyle;
                var createLineStyle = option.CreateLineStyle ?? createDefaultLineStyle ?? createEmptyStyle;

                headerStyles.Add(option, createHeaderStyle(workBook));
                lineStyles.Add(option, createLineStyle(workBook));
            }

            var writer = new SheetWriter<TModel>(sheet, options, headerStyles, lineStyles);

            return writer;
        }
    }
}
