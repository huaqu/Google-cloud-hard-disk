using FreeSql.DataAnnotations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlanFrameworkSVGRename
{
    public class Imgid_unitplan
    {
        [Column(IsIdentity = true, IsPrimary = true)]
        public int imgid { get; set; }
        public string EstateName { get; set; }
        public string PhaseName { get; set; }
        public string BuildingName { get; set; }
        public string Floordesc { get; set; }
        public string Flat { get; set; }
        public string ReferenceURL { get; set; }
    }
    public class Imgid_floorplan
    {
        [Column(IsIdentity = true, IsPrimary = true)]
        public int imgid { get; set; }
        public string EstateName { get; set; }
        public string PhaseName { get; set; }
        public string BuildingName { get; set; }
        public string Floordesc { get; set; }
        public string ReferenceURL { get; set; }
    }
}
