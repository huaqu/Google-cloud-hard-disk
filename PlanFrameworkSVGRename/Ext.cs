using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlanFrameworkSVGRename
{
    class Ext
    {
        /// <summary>
        /// 集合转DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="collection"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(IEnumerable<T> collection)
        {

            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.Add("是否已配对", typeof(string));
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray());
            if (collection.Count() > 0)
            {
                for (int i = 0; i < collection.Count(); i++)
                {
                    ArrayList tempList = new ArrayList();
                    tempList.Add("N");
                    foreach (PropertyInfo pi in props)
                    {
                        object obj = pi.GetValue(collection.ElementAt(i), null);
                        tempList.Add(obj);

                    }
                    object[] array = tempList.ToArray();
                    dt.LoadDataRow(array, true);
                }
            }
            return dt;
        }
        /// <summary>
        /// 数组转DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="collection"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(string[] collection)
        {

            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.Add("是否已配对", typeof(string));
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray()); if (collection.Count() > 0)
            {
                for (int i = 1; i < collection.Count(); i++)
                {
                    string[] vs = collection[i].TrimEnd('\t').Split('\t');
                    dt.LoadDataRow(vs, true);
                }
            }
            return dt;
        }
        public static StreamWriter sw;
       public static StringBuilder strbu;
        public static FileStream file;
        /// <summary>
        /// DataTable转Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static Stream ExportExcel(DataTable dt, string fileName)
        {
            string saveFileName = AppDomain.CurrentDomain.BaseDirectory;
            ////创建文件
             file = new FileStream(saveFileName+ fileName, FileMode.Truncate, FileAccess.ReadWrite);
               //以指定的字符编码向指定的流写入字符
               sw = new StreamWriter(file, Encoding.UTF8);

             strbu = new StringBuilder();
            try
            {
                //string saveFileName = AppDomain.CurrentDomain.BaseDirectory;

              

                //写入标题
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    strbu.Append(dt.Columns[i].ColumnName.ToString() + "\t");
                }
                //加入换行字符串
                strbu.Append(Environment.NewLine);

                //写入内容
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        strbu.Append(dt.Rows[i][j].ToString() + "\t");
                    }
                    strbu.Append(Environment.NewLine);
                }

                sw.Write(strbu.ToString());
                sw.Flush();
                file.Flush();
                return file;
            }
            catch (Exception ex)
            {
                close();
                Log.WriteLogs("LOG", "INFO", $"生成xlsx出错");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                return file;
            }
            
        }
        public static void close()
        {
         
            sw.Close();
            sw.Dispose();
            file.Close();
            file.Dispose();
        }
    }
}
