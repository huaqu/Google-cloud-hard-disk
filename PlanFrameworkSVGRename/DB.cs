using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlanFrameworkSVGRename
{
    public class DB
    {
        static Lazy<IFreeSql> SqlServerLazy = new Lazy<IFreeSql>(() => new FreeSql.FreeSqlBuilder().UseConnectionString(FreeSql.DataType.SqlServer, @"Data Source=10.29.204.34;User Id=sa;Password=centadata2.0;Initial Catalog=HKData;Pooling=true;Min Pool Size=1").Build());
        public static IFreeSql SqlServer => SqlServerLazy.Value;
    }
}
