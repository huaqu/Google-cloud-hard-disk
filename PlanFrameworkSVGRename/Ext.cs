using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
        /// 集合转DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="collection"></param>
        /// <returns></returns>
        public static DataTable ToDataTable2<T>(IEnumerable<T> collection)
        {

            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray());
            if (collection.Count() > 0)
            {
                for (int i = 0; i < collection.Count(); i++)
                {
                    ArrayList tempList = new ArrayList();
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
        public static Stream ExportExcel1(DataTable dt, string fileName)
        {
            string saveFileName = AppDomain.CurrentDomain.BaseDirectory + fileName;
            if (!File.Exists(saveFileName))
            {  
                ////创建文件
                file = File.Create(saveFileName);
            }
            else
            {
                file = new FileStream(saveFileName, FileMode.Truncate, FileAccess.ReadWrite);
            }
               //以指定的字符编码向指定的流写入字符
               sw = new StreamWriter(file, Encoding.UTF8);

             strbu = new StringBuilder();
            try
            {
                //string saveFileName = AppDomain.CurrentDomain.BaseDirectory;

              

                //写入标题
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    strbu.Append(dt.Columns[i].ColumnName.ToString().Trim() +Convert.ToChar(9));
                }
                //加入换行字符串
                strbu.Append(Environment.NewLine);

                //写入内容
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (string.IsNullOrWhiteSpace(dt.Rows[i][j].ToString().Trim()))
                        {
                            strbu.Append(" " + Convert.ToChar(9));
                        }
                        strbu.Append(dt.Rows[i][j].ToString().Trim().Replace("\r\n", " ").Replace("\t", " ") + Convert.ToChar(9));
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
        public static Stream ExportExcel(DataTable dt, string fileName, bool isShowExcle=false)
        {
        
            IWorkbook workbook = null;
            string filepath = AppDomain.CurrentDomain.BaseDirectory + fileName;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    //workbook = new HSSFWorkbook();//导出.xls
                    workbook = new XSSFWorkbook();//导出.xlsx
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表  
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数  

                    //设置列头  
                    row = sheet.CreateRow(0);//excel第一行设为列头  
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,  
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据  
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    if (!File.Exists(filepath))
                    {
                        ////创建文件
                        file = File.Create(filepath);
                    }
                    else
                    {
                        file = new FileStream(filepath, FileMode.Truncate, FileAccess.ReadWrite);
                    }
                    workbook.Write(file);//向打开的这个xls文件中写入数据  
                    file = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    file.Flush();
                    return file;
                }
                return file;
            }
            catch (Exception ex)
            {
             
                Log.WriteLogs("LOG", "INFO", $"生成xlsx出错");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                close();
                return file;
            }
        }
        public static void close()
        {
         
           // sw.Close();
            //sw.Dispose();
            file.Close();
            file.Dispose();
        }
    }
}
