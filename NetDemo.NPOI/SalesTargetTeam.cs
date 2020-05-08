using System;
using System.Collections.Generic;
using System.Text;

namespace NetDemo.NPOI
{
    public class SalesTargetTeam
    {
        public string Month { get; set; }
        public string RegionDepartmentName { get; set; }
        public string AreaDepartmentName { get; set; }
        public string BranchDepartmentName { get; set; }
        public string TeamDepartmentName { get; set; }
        public string CustomerClassName { get; set; }
        public List<SalesTargetByTeam> ProjectModels { get; set; }
    }

    public class SalesTargetProject
    {
        public string ProjectName { get; set; }
        public string ProjectType { get; set; }
        public int ColumnsCount { get; set; }
        public int Sequence { get; set; }
    }

    public class SalesTargetByTeam
    {
        public string ProjectName { get; set; }
        public decimal SalesAmount { get; set; }
        public decimal ProfitRate { get; set; }
        public decimal ProfitAmount { get; set; }
        public int NewCustomerCount { get; set; }
        public int SalesCount { get; set; }
    }

    public class GroupModel
    {
        public string Month { get; set; }

        public int RowIndexStart { get; set; }

        public int RowIndexEnd { get; set; }

        public List<int> RowIndexs { get; set; }

        public Dictionary<string, GroupModel> ChildModels { get; set; }

        public List<SalesTargetTeam> TeamModels { get; set; }
    }
}
