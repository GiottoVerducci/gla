using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace GLA
{
    public static class Const
    {
        public const char NBSP = '\u00a0';
    }
    
    public static class SheetDataLoader<T>
        where T : ArmyPart, new()
    {
        private static readonly List<Tuple<PropertyInfo, ExcelDataAttribute>> _metadata = new List<Tuple<PropertyInfo, ExcelDataAttribute>>();

        static SheetDataLoader()
        {
            PropertyInfo[] propertyInfos = typeof(T).GetProperties();
            foreach (var propertyInfo in propertyInfos)
            {
                var excelDataAttribute = GetAttribute<ExcelDataAttribute>(propertyInfo);
                if (excelDataAttribute == null)
                    continue;
                _metadata.Add(new Tuple<PropertyInfo, ExcelDataAttribute>(propertyInfo, excelDataAttribute));
            }
        }

        public static TAttribute GetAttribute<TAttribute>(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
                return default(TAttribute);

            var attributes = propertyInfo.GetCustomAttributes(typeof(TAttribute), true);
            return (attributes.Length > 0) ? attributes.OfType<TAttribute>().FirstOrDefault() : default(TAttribute);
        }

        /// <summary>
        /// Read the data in the sheet and returns the list of items read.
        /// </summary>
        /// <param name="workbookPart">The workbook part containing the sheet.</param>
        /// <param name="sheetData">The excel sheet.</param>
        /// <param name="army">The army in which the items will be added.</param>
        /// <returns></returns>
        public static List<T> Read(WorkbookPart workbookPart, SheetData sheetData, Army army)
        {
            var result = new List<T>();
            bool firstRow = true;
            List<Action<T, string>> setters = null;

            foreach (Row row in sheetData.Elements<Row>())
            {
                if (firstRow)
                {
                    setters = AssignSetters(workbookPart, row);
                    firstRow = false;
                    continue;
                }

                var item = new T { Army = army };
                bool hasData = false;
                foreach (Cell c in row.Elements<Cell>())
                {
                    if (c.CellValue != null)
                    {
                        hasData = true;
                        string text = ExcelLoader.GetCellValue(c, workbookPart);
                        var cellReference = c.CellReference.ToString();

                        var cellIndex = GetCellColumnIndex(cellReference);
                        if (setters[cellIndex] != null)
                        {
                            text = ExcelLoader.GetCleanText(text);
                            setters[cellIndex](item, text);
                        }
                    }
                }
                if(hasData)
                    result.Add(item);
            }
            return result;
        }

        /// <summary>
        /// Returns the 0-based cell column index from a cell reference.
        /// </summary>
        /// <param name="cellReference">The reference of the cell, such as B12.</param>
        /// <returns>The column index of the cell.</returns>
        /// <example>A25 returns 0. E3 returns 5. AC10 returns 28.</example>
        private static int GetCellColumnIndex(string cellReference)
        {
            int result = 0;
            foreach (var c in cellReference)
            {
                if (c >= 'A' && c <= 'Z')
                    result = (result * 26) + (c - 'A' + 1);
                else
                    break;
            }
            return result - 1;
        }

        private static List<Action<T, string>> AssignSetters(WorkbookPart workbookPart, Row row)
        {
            var result = new List<Action<T, string>>();

            foreach (Cell c in row.Elements<Cell>())
            {
                if (c.CellValue != null)
                {
                    string columnName = ExcelLoader.GetCellValue(c, workbookPart);
                    var setter = GetSetter(columnName);
                    result.Add(setter);
                }
            }
            return result;
        }

        public delegate void Setter(T entity, string value);

        private static Action<T, string> GetSetter(string columnName)
        {
            var metadata = _metadata.FirstOrDefault(p => p.Item2.ColumnName == columnName);
            if (metadata == null)
                return null;

            PropertyInfo propertyInfo = metadata.Item1;
            var setter = propertyInfo.GetSetMethod();
            //var setterDelegate = setter.CreateDelegate(typeof(Setter));
            var setterDelegate = Delegate.CreateDelegate(typeof(Setter), setter);
            return (item, s) => ((Setter)setterDelegate)(item, s);
        }
    }
}