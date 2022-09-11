using ClosedXML.Excel;
using System.Reflection;
using EasyUseExcel.Attribute;

namespace EasyUseExcel
{
    public static class ExcelReader
    {
        /// <summary>
        /// 執行轉入
        /// </summary>
        /// <typeparam name="T">使用ExcelImport.attribute.ColumnIndexAttribute定義欄位順序的MODEL</typeparam>
        /// <param name="stream">FileStream</param>
        /// <param name="BeginSheet">開始工作表</param>
        /// <param name="BeginRow">開始行</param>
        /// <returns></returns>
        public static IList<T> Excute<T>(Stream stream, int BeginSheet, int BeginRow)
        {

            var results = new List<T>();

            XLWorkbook workbook = new XLWorkbook(stream);
            var ws = workbook.Worksheet(BeginSheet);

            var rowCount = ws.RowsUsed().ToList().Count;
            for (var i = 1; i <= rowCount; i++)
            {
                if (BeginRow >= i) continue;

                var row = ws.Row(i);

                var o = GetValue<T>(row);
                results.Add(o);
            }

            stream.Flush();
            stream.Close();
            return results;

        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        private static T GetValue<T>(IXLRow row)
        {
            var props = typeof(T).GetProperties();

            var propertyInfos = props
                .Where(o => !o.GetCustomAttributes(true)
                .Any(attr => typeof(IgnoreAttribute).Name.Equals(attr.GetType().Name)))
                .ToArray();

            var result = Activator.CreateInstance(typeof(T));

            var i = 1;
            foreach (var propInfo in propertyInfos)
            {

                var attrbiute = propInfo.GetCustomAttribute<OrderAttribute>();

                IXLCell cell;
                if (attrbiute != null)
                {
                    cell = row.Cell(attrbiute.Index);

                    try
                    {
                        propInfo.SetValue(result, cell.Value);
                    }
                    catch (Exception e)
                    {
                        try
                        {
                            Transformation(result, propInfo, cell);
                        }
                        catch (Exception e2)
                        {
                            throw new Exception(string.Format("儲存格格式不符; Cell Index: {0}, Object Type:{1}", attrbiute.Index, cell.Value.GetType().Name));
                        }
                    }
                }

            }

            return (T)result;
        }
        /// <summary>
        /// 轉型
        /// </summary>
        private static void Transformation(object result, PropertyInfo prop, IXLCell cell)
        {
            var type = prop.PropertyType;


            if (type == typeof(int))
            {
                if (cell.Value.GetType() == typeof(string))
                {
                    prop.SetValue(result, new Nullable<int>(int.Parse(cell.GetString())));
                }
                else
                {
                    prop.SetValue(result, new Nullable<int>(Convert.ToInt32(cell.GetDouble())));
                }
            }
            else if (type == typeof(int?))
            {
                if (cell.Value.GetType() == typeof(string))
                {
                    prop.SetValue(result, int.Parse(cell.GetString()));
                }
                else
                {
                    prop.SetValue(result, Convert.ToInt32(cell.GetDouble()));
                }
            }
            else if (type == typeof(double))
            {
                if (cell.Value.GetType() == typeof(string))
                {
                    prop.SetValue(result, double.Parse(cell.GetString()));
                }
                else
                {
                    prop.SetValue(result, cell.GetDouble());
                }

            }
            else if (type == typeof(double?))
            {
                if (cell.Value.GetType() == typeof(string))
                {
                    prop.SetValue(result, new Nullable<double>(double.Parse(cell.GetString())));
                }
                else
                {
                    prop.SetValue(result, new Nullable<double>(cell.GetDouble()));
                }
            }
            else if (type == typeof(DateTime))
            {
                if (cell.Value.GetType() == typeof(string))
                {
                    prop.SetValue(result, DateTime.Parse(cell.GetString()));
                }
                else
                {
                    prop.SetValue(result, cell.GetDateTime());
                }
            }
            else if (type == typeof(DateTime?))
            {
                if (cell.Value.GetType() == typeof(string))
                {
                    prop.SetValue(result, new Nullable<DateTime>(DateTime.Parse(cell.GetString())));
                }
                else
                {
                    prop.SetValue(result, new Nullable<DateTime>(cell.GetDateTime()));
                }
            }
            else if (type == typeof(string))
            {
                prop.SetValue(result, cell.GetString());
            }
            else
            {
                throw new Exception();
            }
        }

    }
}
