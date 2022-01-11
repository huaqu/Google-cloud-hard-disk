using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlanFrameworkSVGRename
{
    public class Result
    {
        public string nextPageToken { get; set; }
        public List<Google.Apis.Drive.v3.Data.File> list { get; set; }
    }
}
