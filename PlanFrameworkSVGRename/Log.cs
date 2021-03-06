using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlanFrameworkSVGRename
{
    class Log
    {
        /// <summary>
        /// 日志部分
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="type"></param>
        /// <param name="content"></param>
        public static void WriteLogs(string fileName, string type, string content)
        {
            try
            {

                string path = AppDomain.CurrentDomain.BaseDirectory;
                if (!string.IsNullOrEmpty(path))
                {
                    path = AppDomain.CurrentDomain.BaseDirectory + "LOG\\" + fileName;
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    path = path + "\\" + DateTime.Now.ToString("yyyyMMdd");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    string errlog= path + "\\" + DateTime.Now.ToString("yyyyMMdd") + "_error.txt";
                   
                    path = path + "\\" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
                    if (!File.Exists(path))
                    {
                        FileStream fs = File.Create(path);
                        fs.Close();
                    }
                    if (!File.Exists(errlog))
                    {
                        FileStream fs = File.Create(errlog);
                        fs.Close();
                    }
                    if (File.Exists(path))
                    {
                        StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.UTF8);
                        sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + type + "-->" + content);
                        //  sw.WriteLine("----------------------------------------");
                        sw.Close();
                    }
                    if (type == "ERROR")
                    {
                        if(File.Exists(errlog))
                    {
                            StreamWriter sw = new StreamWriter(errlog, true, System.Text.Encoding.UTF8);
                            sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + type + "-->" + content);
                            //  sw.WriteLine("----------------------------------------");
                            sw.Close();
                        }
                    }
                    if (type != "INFO")
                    {
                        return;
                    }
                    Form1.form.show(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + type + "-->" + content);
                }
            }
            catch (Exception)
            {

              
            }
        }
    }
}
