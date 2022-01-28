using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace PlanFrameworkSVGRename
{
    public partial class Form1 : Form
    {
        public static Form1 form;
        GoogleDriveDriveService service;
        Imgid_unitplan imgid_Unitplan = new Imgid_unitplan();
        string nowid = "";
        /// <summary>
        /// 当前文件夹
        /// </summary>
        TreeNode nownode;
        public Form1()
        {
            InitializeComponent();
            form = this;
        }
        /// <summary>
        /// 日志
        /// </summary>
        /// <param name="content"></param>
        public void show(string content)
        {
            Task.Run(() =>
            {
                this.richTextBox1.BeginInvoke(
               new Action(() =>
               {
                   this.richTextBox1.Text = content + "\r\n" + this.richTextBox1.Text;
               }));
            });
        }
        /// <summary>
        /// 创建treeview根节点和listview
        /// </summary>
        private void CreatTree()
        {
            try
            {
                string[] name = ConfigurationManager.AppSettings["FloderName"].Replace("，", ",").Split(',');
                List<Google.Apis.Drive.v3.Data.File> list = service.GetSharedWithMeFileList().list;
                list.Sort(delegate (Google.Apis.Drive.v3.Data.File p1, Google.Apis.Drive.v3.Data.File p2)
                {
                    return p1.Name.CompareTo(p2.Name);//升序
                });

                this.treeView1.BeginInvoke(
                    new Action(() =>
                    {
                        this.listView1.Items.Clear();
                        this.listView1.BeginUpdate();
                    }));
                foreach (var item in list)
                {
                    if (!name.Contains(item.Name))
                    {
                        continue;
                    }
                    ListViewItem listView = new ListViewItem();
                    listView.SubItems[0].Text = item.Name;
                    FileType fileType = filetype(item.MimeType);
                    listView.SubItems.Add(fileType.type);
                    listView.ImageIndex = fileType.image;
                    listView.SubItems.Add(treeView1.Nodes[0].FullPath);
                    listView.SubItems.Add(item.Id);
                    listView.SubItems.Add(item.Parents == null || item.Parents.Count == 0 ? "" : item.Parents.First());
                    this.treeView1.BeginInvoke(
                 new Action(() =>
                 {
                     listView1.Items.Add(listView);
                 }));

                    if (fileType.type != "文件夾")
                    {
                        continue;
                    }
                    TreeNode node = new TreeNode(); //定义根节点
                    node.Name = item.Id; //将类Model的各个属性赋值给根节点
                    node.Text = item.Name;
                    if (treeView1.Nodes[0].Nodes.ContainsKey(node.Name))
                    {
                        continue;
                    }
                    this.treeView1.BeginInvoke(
                     new Action(() =>
                     {
                         treeView1.Nodes[0].Nodes.Add(node);//将节点node作为treeView1的根节点
                         toolStripButton5.AutoCompleteCustomSource.Add(node.FullPath);
                     }));
                }
                this.treeView1.BeginInvoke(
                  new Action(() =>
                  {
                      this.listView1.Columns[0].Width = -1;
                      this.listView1.Columns[1].Width = -2;
                      this.listView1.Columns[2].Width = -1;
                      this.listView1.EndUpdate();  //结束数据处理，UI界面一次性绘制。
                      treeView1.ExpandAll(); //展开所有节点
                  }));
            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"查询共享文件出錯:{ex.Message}");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
            }
        }
        private void CreateListView()
        {
            string[] name = ConfigurationManager.AppSettings["FloderName"].Replace("，", ",").Split(',');
            List<Google.Apis.Drive.v3.Data.File> list = service.GetSharedWithMeFileList().list;
            list.Sort(delegate (Google.Apis.Drive.v3.Data.File p1, Google.Apis.Drive.v3.Data.File p2)
            {
                return p1.Name.CompareTo(p2.Name);//升序
            });
            this.treeView1.BeginInvoke(
                  new Action(() =>
                  {
                      this.listView1.Items.Clear();
                      this.listView1.BeginUpdate();
                  }));
            foreach (var item in list)
            {
                if (!name.Contains(item.Name))
                {
                    continue;
                }
                ListViewItem listView = new ListViewItem();
                listView.SubItems[0].Text = item.Name;
                FileType fileType = filetype(item.MimeType);
                listView.SubItems.Add(fileType.type);
                listView.ImageIndex = fileType.image;
                listView.SubItems.Add(treeView1.Nodes[0].FullPath);
                listView.SubItems.Add(item.Id);
                listView.SubItems.Add(item.Parents == null || item.Parents.Count == 0 ? "" : item.Parents.First());
                this.treeView1.BeginInvoke(
                    new Action(() =>
                    {
                        listView1.Items.Add(listView);
                    }));
                if (fileType.type != "文件夾")
                {
                    continue;
                }

                TreeNode node = new TreeNode(); //定义根节点
                node.Name = item.Id; //将类Model的各个属性赋值给根节点
                node.Text = item.Name;
                if (treeView1.Nodes.ContainsKey(node.Name))
                {
                    return;
                }
                this.treeView1.BeginInvoke(
                  new Action(() =>
                  {
                      treeView1.Nodes[0].Nodes.Add(node);//将节点node作为treeView1的根节点
                      toolStripButton5.AutoCompleteCustomSource.Add(node.FullPath);
                  }));
                Task.Run(() =>
                {
                    CreatTree2(node);
                });
            }
            this.treeView1.BeginInvoke(
                new Action(() =>
                {
                    this.listView1.Columns[0].Width = -1;
                    this.listView1.Columns[1].Width = -2;
                    this.listView1.Columns[2].Width = -1;
                    this.listView1.EndUpdate();  //结束数据处理，UI界面一次性绘制。
                    treeView1.Nodes[0].Expand();
                }));

        }
        /// <summary>
        /// 递归创建treeview节点
        /// </summary>
        /// <param name="node1"></param>
        private void CreatTree2(TreeNode node1)
        {
            List<Google.Apis.Drive.v3.Data.File> list = service.Getfolderbyparentsid(node1.Name).list;
            list.Sort(delegate (Google.Apis.Drive.v3.Data.File p1, Google.Apis.Drive.v3.Data.File p2)
            {
                return p1.Name.CompareTo(p2.Name);//升序
            });
            foreach (var item in list)
            {
                TreeNode node = new TreeNode(); //定义根节点
                node.Name = item.Id; //将类Model的各个属性赋值给根节点
                node.Text = item.Name;
                if (!node1.Nodes.ContainsKey(node.Name))
                {
                    this.treeView1.BeginInvoke(
                          new Action(() =>
                          {
                              node1.Nodes.Add(node);
                              toolStripButton5.AutoCompleteCustomSource.Add(node.FullPath);
                          }
                          ));//将节点node作为treeView1的根节点
                }
                else
                {
                    node = node1.Nodes.Find(node.Name, false).First();
                }

                CreatTree2(node);

            }
            //  nownode.ExpandAll(); //展开所有节点

        }
        /// <summary>
        /// 创建下一级treeview节点
        /// </summary>
        /// <param name="node1"></param>
        private void CreatTree4(TreeNode node1)
        {
            List<Google.Apis.Drive.v3.Data.File> list = service.Getfolderbyparentsid(node1.Name).list;
            list.Sort(delegate (Google.Apis.Drive.v3.Data.File p1, Google.Apis.Drive.v3.Data.File p2)
            {
                return p1.Name.CompareTo(p2.Name);//升序
            });
            foreach (var item in list)
            {
                TreeNode node = new TreeNode(); //定义根节点
                node.Name = item.Id; //将类Model的各个属性赋值给根节点
                node.Text = item.Name;
                if (!node1.Nodes.ContainsKey(node.Name))
                {
                    this.treeView1.Invoke(
                          new Action(() =>
                          {
                              node1.Nodes.Add(node);
                              toolStripButton5.AutoCompleteCustomSource.Add(node.FullPath);
                          }
                          ));//将节点node作为treeView1的根节点
                }
                else
                {
                    node = node1.Nodes.Find(node.Name, false).First();
                }
            }
            //  nownode.ExpandAll(); //展开所有节点

        }
        /// <summary>
        /// 判断文件类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private FileType filetype(string type)
        {
            FileType fileType = new FileType();
            switch (type)
            {
                case "application/vnd.google-apps.folder":
                    fileType.image = 0;
                    fileType.type = "文件夾"; break;
                case "image/svg+xml":
                    fileType.image = 1;
                    fileType.type = "图片"; break;
                case "application/pdf":
                    fileType.image = 2;
                    fileType.type = "PDF"; break;
                case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                    fileType.image = 3;
                    fileType.type = "xlsx"; break;
                default:
                    fileType.image = 4;
                    fileType.type = "未知文件"; break;
            }
            return fileType;
        }
        /// <summary>
        /// listView双击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (this.listView1.SelectedItems[0].SubItems[1].Text != "文件夾")
                {
                    return;
                }
                string id = this.listView1.SelectedItems[0].SubItems[3].Text;
                nownode = treeView1.Nodes.Find(id, true).First();
                toolStripButton5.Text = nownode.FullPath;
                nowid = id;
                Task.Run(() =>
                {
                    try
                    {
                        CreatTrees(id); ;
                        getimg();
                    }
                    catch (Exception ex)
                    {
                        Log.WriteLogs("LOG", "INFO", $"查询出錯:{ex.Message}");
                        Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                        toolStripStatusLabel1.Text = "查询出錯";
                    }
                });

            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"查询出錯:{ex.Message}");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                toolStripStatusLabel1.Text = "查询出錯";
            }

        }
        /// <summary>
        /// 创建下一级treeview节点
        /// </summary>
        /// <param name="id"></param>
        private void CreatTrees(string id)
        {
            List<Google.Apis.Drive.v3.Data.File> list = service.GetFilesbyparentsid(id).list;
            list.Sort(delegate (Google.Apis.Drive.v3.Data.File p1, Google.Apis.Drive.v3.Data.File p2)
            {
                return p1.Name.CompareTo(p2.Name);//升序
            });
            this.treeView1.BeginInvoke(
                      new Action(() =>
                      {

                          this.listView1.Items.Clear();
                          this.listView1.BeginUpdate();
                      }));
            foreach (var item in list)
            {
                ListViewItem listView = new ListViewItem();
                listView.SubItems[0].Text = item.Name;
                FileType fileType = filetype(item.MimeType);
                listView.SubItems.Add(fileType.type);
                listView.ImageIndex = fileType.image;
                listView.SubItems.Add(treeView1.Nodes.Find(id, true).First().FullPath);
                listView.SubItems.Add(item.Id);
                listView.SubItems.Add(item.Parents == null || item.Parents.Count == 0 ? "" : item.Parents.First());
                this.treeView1.BeginInvoke(
                   new Action(() =>
                   {
                       listView1.Items.Add(listView);
                   }));

                if (fileType.type != "文件夾")
                {
                    continue;
                }
                TreeNode node = new TreeNode(); //定义根节点
                node.Name = item.Id; //将类Model的各个属性赋值给根节点
                node.Text = item.Name;
                if (nownode.Nodes.ContainsKey(node.Name))
                {
                    continue;
                }
                this.treeView1.Invoke(
                 new Action(() =>
                 {
                     nownode.Nodes.Add(node);//将节点node作为treeView1的根节点
                     toolStripButton5.AutoCompleteCustomSource.Add(node.FullPath);
                 }));

            }
            this.treeView1.BeginInvoke(
                    new Action(() =>
                    {
                        this.listView1.Columns[0].Width = -1;
                        this.listView1.Columns[1].Width = -2;
                        this.listView1.Columns[2].Width = -1;
                        this.listView1.EndUpdate();  //结束数据处理，UI界面一次性绘制。
                        nownode.Expand();
                    }));
        }
        /// <summary>
        /// 创建下一级treeview节点
        /// </summary>
        /// <param name="id"></param>
        private List<TreeNode> CreatTrees3(string id)
        {

            List<Google.Apis.Drive.v3.Data.File> list = service.Getfolderbyparentsid(id).list;
            list.Sort(delegate (Google.Apis.Drive.v3.Data.File p1, Google.Apis.Drive.v3.Data.File p2)
            {
                return p1.Name.CompareTo(p2.Name);//升序
            });
            List<TreeNode> TreeNode = new List<TreeNode>();
            foreach (var item in list)
            {
                TreeNode node = new TreeNode(); //定义根节点
                node.Name = item.Id; //将类Model的各个属性赋值给根节点
                node.Text = item.Name;
                TreeNode.Add(node);
            }
            return TreeNode;
        }
        /// <summary>
        /// treeView双击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (treeView1.SelectedNode == treeView1.Nodes[0])
                {
                    return;
                }
                string id = treeView1.SelectedNode.Name;
                nownode = treeView1.SelectedNode;
                toolStripButton5.Text = nownode.FullPath;
                nowid = id;
                Task.Run(() =>
                {
                    try
                    {
                        CreatTrees(id); ;
                        getimg();
                    }
                    catch (Exception ex)
                    {
                        Log.WriteLogs("LOG", "INFO", $"查询出錯:{ex.Message}");
                        Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                        toolStripStatusLabel1.Text = "查询出錯";
                    }
                });
            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"查询出錯:{ex.Message}");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                toolStripStatusLabel1.Text = "查询出錯";
            }
        }
        /// <summary>
        /// 获取当前所在节点的位置
        /// </summary>
        void getimg()
        {
            try
            {
                if (nownode == null)
                {
                    return;
                }
                if (nownode.Text.Contains("2-已完成切圖"))
                {
                    imgid_Unitplan.EstateName = GetEstateName(nownode.Parent.Text);
                    imgid_Unitplan.BuildingName = "a";
                    imgid_Unitplan.PhaseName = "";
                    imgid_Unitplan.Flat = null;
                    imgid_Unitplan.Floordesc = null;
                }
                else if (nownode.Nodes.Cast<TreeNode>().Where(a => a.Text.Contains("2-已完成切圖")).Count() > 0)
                {
                    imgid_Unitplan.EstateName = GetEstateName(nownode.Text);
                    imgid_Unitplan.BuildingName = null;
                    imgid_Unitplan.PhaseName = "";
                    imgid_Unitplan.Flat = null;
                    imgid_Unitplan.Floordesc = null;
                }
                else if (nownode.Parent.Text.Contains("2-已完成切圖"))
                {
                    imgid_Unitplan.EstateName = GetEstateName(nownode.Parent.Parent.Text);
                    imgid_Unitplan.PhaseName = GetPhaseName(nownode.Text);
                    imgid_Unitplan.BuildingName = null;
                    imgid_Unitplan.Flat = null;
                    imgid_Unitplan.Floordesc = null;
                }
                else if (nownode.Parent.Parent != null&& nownode.Parent.Parent.Text.Contains("2-已完成切圖"))
                {
                  
                        imgid_Unitplan.BuildingName = nownode.Text;
                        imgid_Unitplan.PhaseName = GetPhaseName(nownode.Parent.Text);
                        imgid_Unitplan.EstateName = GetEstateName(nownode.Parent.Parent.Parent.Text);
                        imgid_Unitplan.Flat = null;
                        imgid_Unitplan.Floordesc = null;

                }
                else
                {
                    imgid_Unitplan = new Imgid_unitplan();
                }
            }
            catch (Exception e)
            {
                Log.WriteLogs("LOG", "INFO", $"获取路径出错:{e.Message}");
                Log.WriteLogs("LOG", "ERROR", $"{e.ToString()}");
            }
           
          
        }
        /// <summary>
        /// 获取PhaseName
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        string GetPhaseName(string name)
        {
            if (name.Contains("沒有"))
            {
                return null;
            }
            else
            {
                return name;
            }
        }
        /// <summary>
        /// 获取EstateName
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        string GetEstateName(string name)
        {
            string[] vs = name.Split('_');
            if (vs.Length>1)
            {
                return vs[1];
            }
            else
            {
                return "";
            }
           
        }
        /// <summary>
        /// 获取保存副本的文件夹
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        bool GetSaveFolder(TreeNode node)
        {
            string Estateid = "";
            if (nownode == null)
            {
                Log.WriteLogs("LOG", "INFO", $"请重新选择文件");
                return false;
            }
            if (nownode.Text.Contains("2-已完成切圖"))
            {
                Estateid = nownode.Parent.Name;
            }
         
            else if ((nownode.Nodes.Cast<TreeNode>().Where(a=>a.Text.Contains("2-已完成切圖")).Count()>0)|| node.Text.Contains("2-已完成切圖"))
            {
                Estateid = nownode.Name;
            }
            else if (nownode.Parent.Text.Contains("2-已完成切圖"))
            {
                Estateid = nownode.Parent.Parent.Name;
            }
            else if (nownode.Parent.Parent != null&& nownode.Parent.Parent.Text.Contains("2-已完成切圖"))
            {
                    Estateid = nownode.Parent.Parent.Parent.Name;
            }
            else if (node.Nodes.Cast<TreeNode>().Where(a => a.Text.Contains("2-已完成切圖")).Count() > 0)
            {
                Estateid = node.Name;
            }
            else
            {
                Log.WriteLogs("LOG", "INFO", $"请重新选择文件");
                return false;
            }
            Google.Apis.Drive.v3.Data.File file = service.GetFloorFile(Estateid);
            service.FloorPlanid = file.Id;
            file = service.GetUnitFile(Estateid);
            service.UnitPlanid = file.Id;
            return true;
        }
        /// <summary>
        /// 保存副本
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (this.listView1.SelectedIndices.Count == 0)
            {
                Log.WriteLogs("LOG", "INFO", $"请先选择文件");
                return;
            }
          
            string id = this.listView1.SelectedItems[0].SubItems[3].Text;
            TreeNode node = treeView1.Nodes.Find(id, true).First();
            nownode = node.Parent;
            toolStripButton5.Text = nownode.FullPath;
            nowid = nownode.Name;
            getimg();
            try
            {
                CreatTree4(node);
                if (!GetSaveFolder(node))
                {
                    return;
                } 
            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"获取Unit和Floor文件夹");
                Log.WriteLogs("LOG", "INFO", $"{ex.ToString()}");
                return;
            }
         
            
           
          //  CreatTree2(node);
            if (this.listView1.SelectedItems[0].SubItems[1].Text != "文件夾")
            {
                string[] vs = this.listView1.SelectedItems[0].SubItems[0].Text.Split('.')[0].Split('_');
                imgid_Unitplan.Floordesc = vs[0].Replace("-", "/F-") + "/F";
                imgid_Unitplan.Flat = vs.Length == 2 ? vs[1] + "室" : null;
            }

            Task.Run(() =>
            {
                try
                {
                    Log.WriteLogs("LOG", "INFO", $"开始保存副本");
                    service.GetData(imgid_Unitplan.EstateName == null ? GetEstateName(node.Text) : imgid_Unitplan.EstateName);
                    service.CopyAll(id, imgid_Unitplan);
                    getimg();
                    service.SaveData();
                    Log.WriteLogs("LOG", "INFO", $"保存副本完成");
                  //  toolStripStatusLabel1.Text = "保存副本完成";

                }
                catch (Exception ex)
                {
                    Log.WriteLogs("LOG", "INFO", $"保存副本出错:{ex.Message}");
                    Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                }
                this.listView1.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        if (string.IsNullOrEmpty(nowid))
                        {
                            CreatTree();
                        }
                        else
                        {
                            CreatTrees(nowid);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.WriteLogs("LOG", "INFO", $"刷新出错:{ex.Message}");
                        Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                       // toolStripStatusLabel1.Text = "刷新出错";
                    }
                }));

            });
        }
        /// <summary>
        /// 生成excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton1_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.listView1.SelectedIndices.Count == 0)
                {
                    Log.WriteLogs("LOG", "INFO", $"请先选择文件");
                    return;
                }

                string id = this.listView1.SelectedItems[0].SubItems[3].Text;
                TreeNode node = treeView1.Nodes.Find(id, true).First();
                nownode = node.Parent;
                nowid = nownode.Name;
                getimg();
                CreatTree4(node);
                if (!string.IsNullOrWhiteSpace(GetEstateName(node.Text)))
                {
                    int flag=0;
                    string floorexcl = GetEstateName(node.Text) + "_imgid_floorplan.xlsx";
                    if (service.selectEcel(id, floorexcl))
                    {
                        flag = 1;
                    }
                    else
                    {
                        DialogResult d = MessageBox.Show(floorexcl+"已存在，是否替换", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (d == DialogResult.Yes)
                        {
                            flag = 2;
                        }
                        else
                        {
                            flag = 3;
                        }
                    }
                    service.SaveExcel(id,flag, GetEstateName(node.Text));
                }
                else
                {
                    Log.WriteLogs("LOG", "INFO", $"请重新选择文件夹");
                    return;
                }
              
            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"生成excel出錯:{ex.Message}");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                toolStripStatusLabel1.Text = "生成excel出錯";
            }
        }
        /// <summary>
        /// 刷新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(nowid))
                {
                    CreatTree();
                }
                else
                {
                    toolStripButton5.Text = nownode.FullPath;

                    Task.Run(() =>
                    {
                        try
                        {
                            CreatTrees(nowid);
                        }
                        catch (Exception ex)
                        {
                            Log.WriteLogs("LOG", "INFO", $"查询出錯:{ex.Message}");
                            Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                            toolStripStatusLabel1.Text = "查询出錯";
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"刷新出错:{ex.Message}");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                toolStripStatusLabel1.Text = "刷新出错";
            }
        }
        /// <summary>
        /// 根据路径跳转
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton5_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                try
                {
                    string[] path = toolStripButton5.Text.TrimStart('\\').TrimEnd('\\').Split('\\');
                    if (path.Length == 1 && string.IsNullOrEmpty(path[0]))
                    {
                        nowid = "";
                        imgid_Unitplan = new Imgid_unitplan();
                        nownode = null;
                        CreatTree();
                        return;
                    }
                    TreeNode[] treeNodes = treeView1.Nodes.Cast<TreeNode>().ToArray();
                    for (int i = 0; i < path.Length - 1; i++)
                    {
                        treeNodes = treeNodes.Where(a => a.Text == path[i]).First().Nodes.Cast<TreeNode>().ToArray();
                    }
                    nownode = treeNodes.Where(a => a.Text == path[path.Length - 1]).First();
                    string id = nownode.Name;
                    nowid = id;
                    Task.Run(() =>
                    {
                        try
                        {
                            CreatTrees(id); ;
                            getimg();
                        }
                        catch (Exception ex)
                        {
                            Log.WriteLogs("LOG", "INFO", $"查询出錯:{ex.Message}");
                            Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                            toolStripStatusLabel1.Text = "查询出錯";
                        }
                    });
                }
                catch (InvalidOperationException ex)
                {
                    Log.WriteLogs("LOG", "INFO", $"请检查路径是否正确");
                    Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                    toolStripStatusLabel1.Text = "请检查路径是否正确";
                }
            }
        }
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                service = new GoogleDriveDriveService();
            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"连接Google出錯");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
            }
            try
            {
                toolStripButton5.Text = "共享文件";
                toolStripButton5.AutoCompleteCustomSource.Add("共享文件");
                Log.WriteLogs("LOG", "INFO", $"开始获取文件夹");
                Task.Run(() =>
                {
                   // CreateListView();
                      CreatTree();
                    Log.WriteLogs("LOG", "INFO", $"获取文件完成");
                });

                //CreateListView();
                //  throw new Exception();

            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"初始化出錯:{ex.Message}");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                toolStripStatusLabel1.Text = "初始化出錯";
            }
        }

        /// <summary>
        /// 右键菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_MouseDown(object sender, MouseEventArgs e)
        {

            if (e.Button == MouseButtons.Right)//判断你点的是不是右键
            {
                Point ClickPoint = new Point(e.X, e.Y);
                TreeNode CurrentNode = treeView1.GetNodeAt(ClickPoint);
                if (CurrentNode != null)//判断你点的是不是一个节点
                {
                    CurrentNode.ContextMenuStrip = contextMenuStrip1;
                    //  name = treeView1.SelectedNode.Text.ToString();//存储节点的文本
                    treeView1.SelectedNode = CurrentNode;//选中这个节点
                }
            }
            else
            {
                Point ClickPoint = new Point(e.X, e.Y);
                TreeNode CurrentNode = treeView1.GetNodeAt(ClickPoint);
                if (CurrentNode == treeView1.Nodes[0])//判断你点的是不是一个节点
                {
                    CurrentNode.Toggle();
                }
            }
        }
        /// <summary>
        /// 获取所有子文件夹
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void 获取所有子文件夹ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Log.WriteLogs("LOG", "INFO", $"开始获取子文件夹");
            TreeNode selectedNode = treeView1.SelectedNode;
            Task.Run(() =>
            {
                CreatTree2(selectedNode);

                Log.WriteLogs("LOG", "INFO", $"获取完成");
            });
        }
        /// <summary>
        /// 搜索文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripComboBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                Log.WriteLogs("LOG", "INFO", $"开始搜索");
                string name = toolStripComboBox1.Text;
                if (string.IsNullOrEmpty(name))
                {
                    return;
                }
                Task.Run(() =>
                {
                    List<Google.Apis.Drive.v3.Data.File> list = service.GetAllFilesbyname(name).list;
                    list.Sort(delegate (Google.Apis.Drive.v3.Data.File p1, Google.Apis.Drive.v3.Data.File p2)
                    {
                        return p1.Name.CompareTo(p2.Name);//升序
                    });

                    this.treeView1.BeginInvoke(
                        new Action(() =>
                        {
                            this.listView1.Items.Clear();
                            this.listView1.BeginUpdate();
                        }));
                    foreach (var item in list)
                    {
                        if (item.Shared==false)
                        {
                            continue;
                        }
                        ListViewItem listView = new ListViewItem();
                        listView.SubItems[0].Text = item.Name;
                        FileType fileType = filetype(item.MimeType);
                        listView.SubItems.Add(fileType.type);
                        listView.ImageIndex = fileType.image;
                        listView.SubItems.Add(getpath(item.Parents == null || item.Parents.Count == 0 ? "" : item.Parents.First(), item.Id));
                        listView.SubItems.Add(item.Id);
                        listView.SubItems.Add(item.Parents == null || item.Parents.Count == 0 ? "" : item.Parents.First());
                        this.treeView1.BeginInvoke(
                     new Action(() =>
                     {
                         listView1.Items.Add(listView);
                     }));
                    }
                    this.treeView1.BeginInvoke(
                   new Action(() =>
                   {
                       this.listView1.Columns[0].Width = -1;
                       this.listView1.Columns[1].Width = -2;
                       this.listView1.Columns[2].Width = -1;
                       this.listView1.EndUpdate();  //结束数据处理，UI界面一次性绘制。
                       Log.WriteLogs("LOG", "INFO", $"搜索完成");
                   }));
                });
            }
        }
        /// <summary>
        /// 获取文件路径
        /// </summary>
        /// <param name="Parentsid"></param>
        /// <param name="id"></param>
        /// <returns></returns>
        private string getpath(string Parentsid,string id)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(Parentsid))
                {
                    return "                  ";
                }
                TreeNode[] treeNodes = treeView1.Nodes.Find(id, true);
                if (treeNodes.Length > 0)
                {
                    treeNodes = treeView1.Nodes.Find(Parentsid, true);
                    return treeNodes[0].FullPath;
                }
                getnode(Parentsid, new List<TreeNode>(), id);

                treeNodes = treeView1.Nodes.Find(Parentsid, true);
                if (treeNodes.Length > 0)
                {
                    return treeNodes[0].FullPath;
                }
                else
                    return "";
            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"搜索时，获取文件路径出错");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
                return "";
            }
            
        }
        /// <summary>
        /// 反向递归生成节点
        /// </summary>
        /// <param name="id"></param>
        /// <param name="treeNodes2"></param>
        /// <param name="id1"></param>
        private void getnode(string id, List<TreeNode> treeNodes2,string id1)
        {
            try
            {
                List<TreeNode> treeNodes1 = CreatTrees3(id);
                TreeNode treeNode = treeNodes1.Find(a => a.Name == id1);
                if (treeNode != null)
                {
                    treeNode.Nodes.AddRange(treeNodes2.ToArray());
                }
                TreeNode[] treeNodes = treeView1.Nodes.Find(id, true);
                if (treeNodes.Length > 0)
                {
                    this.treeView1.Invoke(
                      new Action(() =>
                      {
                          treeNodes[0].Nodes.AddRange(treeNodes1.ToArray());
                      }));
                    return;
                }
                else
                {
                    Google.Apis.Drive.v3.Data.File file = service.GetFilesbyId(id);
                    getnode(file.Parents[0], treeNodes1, id);
                }
            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"反向递归生成节点时出错");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
            }
           
        }
    }
}
