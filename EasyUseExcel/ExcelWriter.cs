using ClosedXML.Excel;
using EasyUseExcel.Attribute;
using System.Data;

namespace EasyUseExcel
{
    public static class ExcelWriter
    {
        private static IDictionary<string, List<string>> KeyValue = new Dictionary<string, List<string>>();

        private static DataTable GetTable<T>(IList<T> datas)
        {
            var sheetName = typeof(T).Name;

            if (KeyValue.ContainsKey(sheetName))
            {
                var sheetNames = new List<string>();
                if (KeyValue.TryGetValue(sheetName, out sheetNames))
                {
                    sheetName = string.Format("{0}{1}", sheetName, sheetNames.Count);
                    sheetNames.Add(sheetName);
                }
                else 
                {
                    throw new Exception();
                }
            }
            else
            {
                List<string> sheetNames = new List<string>();
                sheetNames.Add(sheetName);
                KeyValue.Add(sheetName, sheetNames);
            }

            var table = new DataTable(sheetName);
            var index = 0;
            foreach (var data in datas)
            {
                if (typeof(T).Name == "Object")
                {
                    var keyValue = data as IDictionary<string, object>;

                    // create Column 只需一次
                    if (index++ == 0)
                    {
                        foreach (var key in keyValue.Keys)
                        {
                            table.Columns.Add(key);
                        }
                    }

                    IList<object> objs = new List<object>();
                    foreach (var value in keyValue.Values)
                    {
                        objs.Add(value);
                    }
                    table.Rows.Add(objs.ToArray());
                }
                else
                {
                    var props = typeof(T).GetProperties();
                    props = props.Where(o => !o.GetCustomAttributes(true).Any(i => i.GetType().Name.Equals(typeof(IgnoreAttribute).Name)))
                                .ToArray();

                    try
                    {
                        if (props.Any(o => o.GetCustomAttributes(true).Any(i => i.GetType().Name == typeof(OrderAttribute).Name)))
                            props = props.OrderBy(o => ((OrderAttribute)o.GetCustomAttributes(true).Where(i => i.GetType().Name.Equals(typeof(OrderAttribute).Name)).FirstOrDefault()).Index)
                                        .ToArray();
                    }
                    catch (Exception e)
                    {
                        throw new Exception(" PropertyInfo have to a OrderAttribute");
                    }


                    // create Column 只需一次
                    if (index++ == 0)
                    {
                        foreach (var prop in props)
                        {
                            var propName = prop.Name;

                            if (prop.GetCustomAttributes(true).Any(o => o.GetType().Name.Equals(typeof(DisplayAttribute).Name)))
                            {
                                var attr = prop.GetCustomAttributes(true).SingleOrDefault(o => o.GetType().Name.Equals(typeof(DisplayAttribute).Name)) as dynamic;
                                if (attr != null && attr.Name != null)
                                {
                                    propName = attr.Name;
                                }
                            }

                            if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                            {
                                table.Columns.Add(propName, prop.PropertyType.GetGenericArguments()[0]);
                            }
                            else if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                            {

                            }
                            else
                            {
                                table.Columns.Add(propName, prop.PropertyType);
                            }
                        }
                    }

                    IList<object> objs = new List<object>();
                    foreach (var prop in props)
                    {
                        objs.Add(prop.GetValue(data));
                    }
                    table.Rows.Add(objs.ToArray());
                }
            }

            return table;
        }

        private static DataTable GetTable(IList<object> datas)
        {
            var type = datas.FirstOrDefault().GetType();

            var sheetName = datas.FirstOrDefault().GetType() == null
                        ? "Sheet1" : datas.FirstOrDefault().GetType().Name;

            if (KeyValue.ContainsKey(sheetName))
            {
                var sheetNames = new List<string>();
                KeyValue.TryGetValue(sheetName, out sheetNames);
                sheetName = string.Format("{0}{1}", sheetName, sheetNames.Count);
                sheetNames.Add(sheetName);
            }
            else
            {
                List<string> sheetNames = new List<string>();
                sheetNames.Add(sheetName);
                KeyValue.Add(sheetName, sheetNames);
            }


            var table = new DataTable(sheetName);

            if (datas.FirstOrDefault() is IDictionary<string, object>)
            {
                var index = 0;
                foreach (var data in datas)
                {
                    var map = data as IDictionary<string, object>;
                    var obj = new Object();
                    if (index++ == 0)
                    {
                        foreach (var key in map.Keys)
                        {
                            map.TryGetValue(key, out obj);
                            table.Columns.Add(key, obj == null ? typeof(string) : obj.GetType());
                        }
                    }


                    table.Rows.Add(map.Values.ToArray());
                }
            }
            else
            {
                var props = datas.FirstOrDefault().GetType().GetProperties();

                foreach (var prop in props)
                {
                    var propName = prop.Name;

                    if (prop.GetCustomAttributes(true).Any(o => o.GetType().Name.Equals(typeof(DisplayAttribute).Name)))
                    {
                        var attr = prop.GetCustomAttributes(true).SingleOrDefault(o => o.GetType().Name.Equals(typeof(DisplayAttribute).Name)) as dynamic;
                        if (attr != null && attr.Name != null)
                        {
                            propName = attr.Name;
                        }
                    }

                    if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        table.Columns.Add(propName, prop.PropertyType.GetGenericArguments()[0]);
                    }
                    else if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                    {

                    }
                    else
                    {
                        table.Columns.Add(propName, prop.PropertyType);
                    }
                }

                IList<object> objs = new List<object>();

                foreach (var data in datas)
                {
                    objs.Clear();
                    foreach (var prop in props)
                    {
                        objs.Add(prop.GetValue(data));
                    }
                    table.Rows.Add(objs.ToArray());
                }
            }
            return table;
        }


        public static Stream Excute<T>(IList<T> datas)
        {
            MemoryStream ms = new MemoryStream();
            using (XLWorkbook workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(GetTable(datas));
                workbook.SaveAs(ms);
                ms.Position = 0;
                return ms;

            }
        }

        public static Stream Excute<T, T2>(IList<T> datas, IList<T2> data2s)
        {
            MemoryStream ms = new MemoryStream();
            using (XLWorkbook workbook = new XLWorkbook()) 
            {
                workbook.Worksheets.Add(GetTable(datas));
                workbook.Worksheets.Add(GetTable(data2s));

                workbook.SaveAs(ms);
                ms.Position = 0;
                return ms;
            }
        }

        public static Stream Excute<T, T2, T3>(IList<T> datas, IList<T2> data2s, IList<T3> data3s)
        {
            MemoryStream ms = new MemoryStream();
            using (XLWorkbook workbook = new XLWorkbook())
            { 
                workbook.Worksheets.Add(GetTable(datas));
                workbook.Worksheets.Add(GetTable(data2s));
                workbook.Worksheets.Add(GetTable(data3s));

                workbook.SaveAs(ms);
                ms.Position = 0;
                return ms;
            }
        }

        public static Stream Excute<T, T2, T3, T4>(IList<T> datas, IList<T2> data2s, IList<T3> data3s, IList<T4> data4s)
        {
            MemoryStream ms = new MemoryStream();
            using(XLWorkbook workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(GetTable(datas));
                workbook.Worksheets.Add(GetTable(data2s));
                workbook.Worksheets.Add(GetTable(data3s));
                workbook.Worksheets.Add(GetTable(data4s));

                workbook.SaveAs(ms);
                ms.Position = 0;
                return ms;            
            }
        }

        public static Stream Excute<T, T2, T3, T4, T5>(IList<T> datas, IList<T2> data2s, IList<T3> data3s, IList<T4> data4s, IList<T5> data5s)
        {
            MemoryStream ms = new MemoryStream();
            using (XLWorkbook workbook = new XLWorkbook()) 
            {
                workbook.Worksheets.Add(GetTable(datas));
                workbook.Worksheets.Add(GetTable(data2s));
                workbook.Worksheets.Add(GetTable(data3s));
                workbook.Worksheets.Add(GetTable(data4s));
                workbook.Worksheets.Add(GetTable(data5s));

                workbook.SaveAs(ms);
                ms.Position = 0;
                return ms;
            }
        }

        public static Stream Excute(IList<object> datas)
        {
            MemoryStream ms = new MemoryStream();
            XLWorkbook workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add(GetTable(datas));
            workbook.SaveAs(ms);
            ms.Position = 0;
            return ms;
        }       
    }
}
