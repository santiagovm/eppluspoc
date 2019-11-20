using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace EpPlusPoC
{
    public static class EpPlusExtensions
    {
        public static IEnumerable<T> ConvertSheetToObjects<T>(this ExcelWorksheet worksheet) where T : new()
        {
            bool IsColumnAttribute(CustomAttributeData y) => y.AttributeType == typeof(Column);

            var columns = typeof(T).GetProperties()
                                   .Where(propertyInfo => propertyInfo.CustomAttributes.Any(IsColumnAttribute))
                                   .Select(propertyInfo => new
                                                           {
                                                               Property = propertyInfo,
                                                               
                                                               propertyInfo.GetCustomAttributes<Column>()
                                                                           .First()
                                                                           .ColumnIndex
                                                           })
                                   .ToArray();

            IOrderedEnumerable<int> rowIndexesList = worksheet.Cells
                                                              .Select(cell => cell.Start.Row)
                                                              .Distinct()
                                                              .OrderBy(x => x);

            return rowIndexesList.Skip(1)
                                 .Select(rowIndex =>
                                         {
                                             var newObject = new T();

                                             foreach (var col in columns)
                                             {
                                                 // this is the real wrinkle to using reflection
                                                 // Excel stores all numbers as double including int
                                                 ExcelRange val = worksheet.Cells[rowIndex, col.ColumnIndex];

                                                 // if it is numeric it is a double since that is how excel stores all numbers
                                                 if (val.Value == null)
                                                 {
                                                     col.Property.SetValue(newObject, null);
                                                     continue;
                                                 }

                                                 if (col.Property.PropertyType == typeof(int))
                                                 {
                                                     col.Property.SetValue(newObject, val.GetValue<int>());
                                                     continue;
                                                 }

                                                 if (col.Property.PropertyType == typeof(double))
                                                 {
                                                     col.Property.SetValue(newObject,
                                                                           val.GetValue<double>());
                                                     continue;
                                                 }

                                                 if (col.Property.PropertyType == typeof(DateTime))
                                                 {
                                                     col.Property.SetValue(newObject,
                                                                           val.GetValue<DateTime>());
                                                     continue;
                                                 }

                                                 // it is a string
                                                 col.Property.SetValue(newObject, val.GetValue<string>());
                                             }

                                             return newObject;
                                         });
        }
    }
}
