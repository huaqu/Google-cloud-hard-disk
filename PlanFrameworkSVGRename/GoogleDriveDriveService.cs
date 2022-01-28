using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace PlanFrameworkSVGRename
{
    public class GoogleDriveDriveService
    {
        static string[] Scopes = { DriveService.Scope.DriveReadonly, DriveService.Scope.Drive, DriveService.Scope.DriveFile, DriveService.Scope.DriveAppdata, DriveService.Scope.DrivePhotosReadonly };
        static string ApplicationName = " My First Project ";
        public   FilesResource files1;
        DataTable floorplandataTable;
        DataTable unitplandataTable;
        public string UnitPlanid = "";
        public string FloorPlanid = "";
        public GoogleDriveDriveService()
        {
            UserCredential credential;
            string basePath2 = Path.GetDirectoryName(typeof(Program).Assembly.Location);
            string credPath = Path.Combine(basePath2, "credentials.json");
            string tokenPath = Path.Combine(basePath2, "token.json");
            using (var stream =
                new FileStream(credPath, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(tokenPath, true)).Result;
            }

            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            files1 = service.Files;
        }
        public Result GetFilesbyparentsid( string id,string token=null)
        {
            Result result = new Result();
            FilesResource.ListRequest listRequest = files1.List();
            listRequest.PageToken = token;
            listRequest.PageSize = 1000;
            listRequest.Q = $"'{id}' in parents and not trashed";//
            listRequest.Fields = "nextPageToken, files";
            // List files.
            Google.Apis.Drive.v3.Data.FileList fileList;
            fileList = listRequest.Execute();
            result.list = (List<Google.Apis.Drive.v3.Data.File>)fileList.Files;
            if (!string.IsNullOrEmpty(fileList.NextPageToken))
            {
                result.list.AddRange(GetFilesbyparentsid(id, fileList.NextPageToken).list);
            }
            return result;
        }

        public Result Getfolderbyparentsid(string id, string token = null)
        {
            Result result = new Result();
            FilesResource.ListRequest listRequest = files1.List();
            listRequest.PageToken = token;
            listRequest.PageSize = 1000;
            listRequest.Q = $"'{id}' in parents and mimeType = 'application/vnd.google-apps.folder' and not trashed";
            listRequest.Fields = "nextPageToken, files(id, name, mimeType,parents)";
            // List files.
            Google.Apis.Drive.v3.Data.FileList fileList = listRequest.Execute();
            result.list = (List<Google.Apis.Drive.v3.Data.File>)fileList.Files;
            if (!string.IsNullOrEmpty(fileList.NextPageToken))
            {
                result.list.AddRange(Getfolderbyparentsid(id, fileList.NextPageToken).list);
            }
            return result;
        }
        public Result GetAllFilesbyparentsid(string id)
        {
            Result result = new Result();
            FilesResource.ListRequest listRequest = files1.List();
            listRequest.PageSize = 1000;
            listRequest.Q = $"'{id}' in parents";
            listRequest.Fields = "nextPageToken, files(id, name, mimeType,parents)";
            // List files.
            Google.Apis.Drive.v3.Data.FileList fileList = listRequest.Execute();
            result.list = (List<Google.Apis.Drive.v3.Data.File>)fileList.Files;

            foreach (var item in new List<Google.Apis.Drive.v3.Data.File>(result.list))
            {
                if (item.MimeType == "application/vnd.google-apps.folder")
                {
                    IList<Google.Apis.Drive.v3.Data.File> list = GetAllFilesbyparentsid(item.Id).list;
                    result.list.AddRange(list);
                }
            }
            return result;
        }
        public Result GetAllFiles(string PageToken)
        {
            Result result = new Result();
            FilesResource.ListRequest listRequest = files1.List();
            listRequest.PageToken = PageToken;
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(id, name, mimeType,parents)";
            // List files.
            Google.Apis.Drive.v3.Data.FileList fileList = listRequest.Execute();
            result.list = (List<Google.Apis.Drive.v3.Data.File>)fileList.Files;
            result.nextPageToken = fileList.NextPageToken;
            return result;
        }
        public Result GetSharedWithMeFileList()
        {
            Result result = new Result();
            FilesResource.ListRequest listRequest = files1.List();
            listRequest.PageSize = 1000;
            listRequest.Q = $"sharedWithMe and not trashed";
            listRequest.Fields = "files(id, name, mimeType)";
            // List files.
            Google.Apis.Drive.v3.Data.FileList fileList = listRequest.Execute();
            result.list = (List<Google.Apis.Drive.v3.Data.File>)fileList.Files;
            return result;
        }
        public Result GetAllFolders(string PageToken)
        {
            Result result = new Result();
            FilesResource.ListRequest listRequest = files1.List();
            listRequest.Q = $" mimeType = 'application/vnd.google-apps.folder'";
            listRequest.PageToken = PageToken;
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(id, name, mimeType,parents)";
            // List files.
            Google.Apis.Drive.v3.Data.FileList fileList = listRequest.Execute();
            result.list = (List<Google.Apis.Drive.v3.Data.File>)fileList.Files;
            result.nextPageToken = fileList.NextPageToken;
            return result;
        }
        public Result GetAllFilesbyname(string name)
        {
            
            Result result = new Result();
            FilesResource.ListRequest listRequest = files1.List();
            listRequest.PageSize = 1000;
            listRequest.Q = $"name contains '{name}' and not trashed";
            listRequest.Fields = "nextPageToken, files(shared,id, name, mimeType,parents)";
            // List files.
            Google.Apis.Drive.v3.Data.FileList fileList = listRequest.Execute();
            result.list = (List<Google.Apis.Drive.v3.Data.File>)fileList.Files;
            return result;
        }
        public Google.Apis.Drive.v3.Data.File GetFilesbyId(string id)
        {
            FilesResource.GetRequest getRequest = files1.Get(id); 
            getRequest.Fields = "shared,id, name, mimeType,parents";
            return getRequest.Execute();
        }
        public void GetData(string EstateName)
        {
          
            List<Imgid_floorplan> imgid_floorplans1 = DB.SqlServer.Select<Imgid_floorplan>().Where(a => (a.EstateName == "" ? null : a.EstateName) == EstateName).ToList();
            floorplandataTable = Ext.ToDataTable(imgid_floorplans1);

            List<Imgid_unitplan> Imgid_unitplan1 = DB.SqlServer.Select<Imgid_unitplan>().Where(a => (a.EstateName == "" ? null : a.EstateName) == EstateName).ToList();
            unitplandataTable = Ext.ToDataTable(Imgid_unitplan1);
        
            
        }
        public void GetData2(string EstateName)
        {

            List<Imgid_floorplan> imgid_floorplans1 = DB.SqlServer.Select<Imgid_floorplan>().Where(a => (a.EstateName == "" ? null : a.EstateName) == EstateName).ToList();
            floorplandataTable = Ext.ToDataTable2(imgid_floorplans1);

            List<Imgid_unitplan> Imgid_unitplan1 = DB.SqlServer.Select<Imgid_unitplan>().Where(a => (a.EstateName == "" ? null : a.EstateName) == EstateName).ToList();
            unitplandataTable = Ext.ToDataTable2(Imgid_unitplan1);


        }
        public void SaveData()
        
        {
            
            Google.Apis.Drive.v3.Data.File file = new Google.Apis.Drive.v3.Data.File();
            file.Name = "floorplan_mapping_result.xlsx";
            file.MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            file.Parents = new List<string>() { FloorPlanid };
            Stream stream = Ext.ExportExcel(floorplandataTable, file.Name);
            Google.Apis.Upload.IUploadProgress uploadProgress = files1.Create(file, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet").Upload();
            Ext.close();
          
            file.Name = "unitplan_mapping_result.xlsx";
            file.MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            file.Parents = new List<string>() { UnitPlanid };
            stream = Ext.ExportExcel(unitplandataTable, file.Name);
            uploadProgress = files1.Create(file, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet").Upload();
            Ext.close();
            floorplandataTable = null;
            unitplandataTable = null;
        }
        public void SaveData2(string id,string name)
        {

            Google.Apis.Drive.v3.Data.File file = new Google.Apis.Drive.v3.Data.File();
            file.Name = name;
            file.MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            file.Parents = new List<string>() { id };
            Stream stream=null;
            if (name.Contains("floorplan"))
            {
                 stream = Ext.ExportExcel(floorplandataTable, "test.xlsx");
            }
            else
            {
                 stream = Ext.ExportExcel(unitplandataTable, "test.xlsx");
            }
           
            Google.Apis.Upload.IUploadProgress uploadProgress = files1.Create(file, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet").Upload();
            Ext.close();
        }
        public void SaveExcel(string id,int flag, string name)
        {
            GetData2(name);
            string floorexcl = name + "_imgid_floorplan.xlsx";
            string unitexcl = name + "_imgid_unitplan.xlsx";
            List<Google.Apis.Drive.v3.Data.File> list = GetFilesbyparentsid(id).list;
            Google.Apis.Drive.v3.Data.File file;
            if (flag==2)
            {
                 file = list.Find(a => a.Name == floorexcl);
                if (file != null)
                {
                    files1.Delete(file.Id).Execute();
                }
                file = list.Find(a => a.Name == unitexcl);
                if (file != null)
                {
                    files1.Delete(file.Id).Execute();
                }
            }
            int i = 2;
            while (true)
            {
                file = list.Find(a => a.Name == floorexcl);
                if (file != null)
                {
                    floorexcl= name + "_imgid_floorplan"+ "_v" + i+ ".xlsx";
                    unitexcl= name + "_imgid_unitplan"+ "_v" + i + ".xlsx";
                    i++;
                }
                else
                {
                    SaveData2(id, floorexcl);
                    SaveData2(id, unitexcl);
                    Log.WriteLogs("LOG", "INFO", $"生成excel完成");
                    floorplandataTable = null;
                    unitplandataTable = null;
                    return;
                }
                
            }
        }
        public bool selectEcel(string id,string name)
        {
            List<Google.Apis.Drive.v3.Data.File> list = GetFilesbyparentsid(id).list;
            Google.Apis.Drive.v3.Data.File file = list.Find(a => a.Name == name );
                if (file != null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        public void CopyAll(string id, Imgid_unitplan img)
        {
            if (img.EstateName==null)
            {
                img.EstateName = GetEstateName(files1.Get(id).Execute().Name);
                List<Google.Apis.Drive.v3.Data.File> list = GetFilesbyparentsid(id).list;
                foreach (var item in list)
                {
                    img.PhaseName = "";
                    CopyAll(item.Id, img);
                }
                return;
            }
            else if (img.BuildingName == "a")
            {
                img.BuildingName = "";
                img.PhaseName = GetPhaseName(files1.Get(id).Execute().Name);
                CopyAll(id, img);
                return;
            }
            else  if (img.BuildingName == "")
            {
                List<Google.Apis.Drive.v3.Data.File> list = GetFilesbyparentsid(id).list;
                foreach (var item in list)
                {
                    img.BuildingName = item.Name;
                    CopyAll(item.Id, img);
                }
                return;
            }
            else if (img.PhaseName == "")
            {
             
                List<Google.Apis.Drive.v3.Data.File> list = GetFilesbyparentsid(id).list;
                foreach (var item in list)
                {
                    img.BuildingName = "";
                    img.PhaseName = GetPhaseName(item.Name);
                    CopyAll(item.Id, img);
                }
                return;
            }
            else if (img.BuildingName == null)
            {
                img.BuildingName= files1.Get(id).Execute().Name;
                CopyAll(id, img);
                return;
            }
            else if (img.Floordesc == null)
            {
                List<Google.Apis.Drive.v3.Data.File> list = GetFilesbyparentsid(id).list;
                foreach (var item in list)
                {
                    string[] vs = item.Name.Split('.')[0].Split('_');
                    img.Floordesc = vs[0].Replace("-", "/F-").Replace(",", "/F,") + "/F";
                    img.Flat = vs.Length == 2 ? vs[1].TrimStart('0') + "室" : null;
                    Log.WriteLogs("LOG", "INFO", $"开始复制文件：{item.Name}");
                    CopyAll(item.Id, img);
                    img.Floordesc = null;
                    img.Flat = null;
                    Log.WriteLogs("LOG", "INFO", $"结束复制文件：{item.Name}");
                }
                return;
            }
            try
            {
                Google.Apis.Drive.v3.Data.File file = files1.Get(id).Execute();
                if (string.IsNullOrWhiteSpace(img.Flat))
                {
                    var Imgid_floorplans = floorplandataTable.AsEnumerable().Where(a => (a["PhaseName"].ToString().Trim() == "" ? null : a["PhaseName"].ToString().Trim()) == img.PhaseName).Where(a => (a["BuildingName"].ToString().Trim() == "" ? null : a["BuildingName"].ToString().Trim()) == img.BuildingName).Where(a => (a["Floordesc"].ToString().Trim() == "" ? null : a["Floordesc"].ToString().Trim()) == img.Floordesc).ToList();
                    if (Imgid_floorplans.Count == 0)
                    {
                        Log.WriteLogs("LOG", "INFO", $"没有在Imgid_floorplan查询到该文件的数据");
                    }
                    else
                    {
                        file.Name = Imgid_floorplans.First()["imgid"].ToString().Trim() + "." + file.Name.Split('.').Last();
                        Imgid_floorplans.First()["是否已配对"] = "Y";
                    }
                    file.Parents = new List<string>() { FloorPlanid };
                }
                else
                {
                    var imgid_unitplans = unitplandataTable.AsEnumerable().Where(a => (a["PhaseName"].ToString().Trim() == "" ? null : a["PhaseName"].ToString().Trim()) == img.PhaseName).Where(a => (a["BuildingName"].ToString().Trim() == "" ? null : a["BuildingName"].ToString().Trim()) == img.BuildingName).Where(a => (a["Floordesc"].ToString().Trim() == "" ? null : a["Floordesc"].ToString().Trim()) == img.Floordesc).Where(a => (a["Flat"].ToString().Trim() == "" ? null : a["Flat"].ToString().Trim()) == img.Flat).ToList();
                    if (imgid_unitplans.Count == 0)
                    {
                        Log.WriteLogs("LOG", "INFO", $"没有在imgid_unitplan查询到该文件的数据");
                    }
                    else
                    {
                        file.Name = imgid_unitplans.First()["imgid"].ToString().Trim() + "." + file.Name.Split('.').Last();
                        imgid_unitplans.First()["是否已配对"] = "Y";
                    }
                    file.Parents = new List<string>() { UnitPlanid };
                }
                file.Id = null;

                Google.Apis.Drive.v3.Data.File file1 = files1.Copy(file, id).Execute();
            }
            catch (Exception ex)
            {
                Log.WriteLogs("LOG", "INFO", $"复制时出错");
                Log.WriteLogs("LOG", "ERROR", $"{ex.ToString()}");
            }
          
        }
        public Google.Apis.Drive.v3.Data.File GetFloorFile(string id)
        {
            List<Google.Apis.Drive.v3.Data.File> list = GetFilesbyparentsid(id).list;
            Google.Apis.Drive.v3.Data.File file = list.Find(a => a.Name == "3-PDF-SVG-Floor-Plan");
            if (file != null)
            {
                files1.Delete(file.Id).Execute();
            }
                file = new Google.Apis.Drive.v3.Data.File();
                file.Name = "3-PDF-SVG-Floor-Plan";
                file.MimeType = "application/vnd.google-apps.folder";
                file.Parents=new List<string>() { id };
                return files1.Create(file).Execute();
        }
        public Google.Apis.Drive.v3.Data.File GetUnitFile(string id)
        {
            List<Google.Apis.Drive.v3.Data.File> list = GetFilesbyparentsid(id).list;
            Google.Apis.Drive.v3.Data.File file = list.Find(a => a.Name == "4-PDF-SVG-Unit-Plan");
            if (file != null)
            {
                files1.Delete(file.Id).Execute();
            }
         
                file = new Google.Apis.Drive.v3.Data.File();
                file.Name = "4-PDF-SVG-Unit-Plan";
                file.MimeType = "application/vnd.google-apps.folder";
                file.Parents = new List<string>() { id };
                return files1.Create(file).Execute();
        }
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
        string GetEstateName(string name)
        {
            return name.Split('_')[1];
        }
    }
}
