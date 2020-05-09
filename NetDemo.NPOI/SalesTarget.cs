using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace NetDemo.NPOI
{
    public class SalesTarget
    {
        private int _columnPrefix = 5; //前缀列数
        private int _columnSuffix = 4; //后缀列数
        private int _rowTitleCount = 3; //列头占用行数
        private string _normalKeyWord = "常规";
        private string _smithKeyWord = "史密斯";
        private string _profitRate = "毛利率";

        public List<SalesTargetTeam> Get()
        {
            var salesTargetTeams = new List<SalesTargetTeam>();
            var dt = ExcelHelper.ExcelToDataTable(@"D:\文档\20200423销售指标线上化\业绩指标计划导入模板1.xlsx", false);
            if (dt == null || dt.Rows.Count < this._rowTitleCount || dt.Columns.Count < (this._columnPrefix + this._columnSuffix))
                return salesTargetTeams; //如果没有数据，或者行列小于最小数据要求，不做处理

            #region 所在的列及对应项目
            var dicProject = new Dictionary<int, string>(); //所在的列及对应项目，用来查找当前列对应的列名

            DataRow drTitle = dt.Rows[0]; //第二列标题列
            for (var j = 0; j < dt.Columns.Count; j++)
            {
                var value = drTitle[j].ToString().Trim();
                if (string.IsNullOrEmpty(value))
                    continue;

                dicProject.Add(j, value);
            } 
            #endregion

            for (var i = this._rowTitleCount; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                var salesTargetTeam = this.Mapping(dr);
                if (string.IsNullOrEmpty(salesTargetTeam.Month) || string.IsNullOrEmpty(salesTargetTeam.RegionDepartmentName) || string.IsNullOrEmpty(salesTargetTeam.AreaDepartmentName) || string.IsNullOrEmpty(salesTargetTeam.BranchDepartmentName) || string.IsNullOrEmpty(salesTargetTeam.TeamDepartmentName))
                    continue; //关键字段为空，可能是合计行，跳过不处理

                var colIndex = this._columnPrefix;
                var lastColumnIndex = dt.Columns.Count - this._columnSuffix; //最后数据列，排除最后四行合计列
                while (dicProject.ContainsKey(colIndex) && colIndex < lastColumnIndex)
                {
                    var colName = dicProject[colIndex];
                    var projectModel = new SalesTargetByTeam() { ProjectName = colName }; //项目名称
                    if (colName.Contains(this._normalKeyWord))
                    {
                        var hasData = this.SetData(projectModel, dr, colIndex, ProjectTypeEnum.Normal);
                        colIndex += 4; //常规占用4行：销售额、毛利率、毛利额、新客户数
                        if (!hasData)
                            continue;
                    }
                    else if (colName.Contains(this._smithKeyWord))
                    {
                        var hasData = this.SetData(projectModel, dr, colIndex, ProjectTypeEnum.Smith);
                        colIndex += 1; //史密斯占用4行：台数
                        if (!hasData)
                            continue;
                    }
                    else
                    {
                        var hasData = this.SetData(projectModel, dr, colIndex, ProjectTypeEnum.Major);
                        colIndex += 3; //核心占用3行：销售额、毛利率、毛利额
                        if (!hasData)
                            continue;
                    }

                    salesTargetTeam.ProjectModels.Add(projectModel);
                }

                if (salesTargetTeam.ProjectModels.Count > 0)
                    salesTargetTeams.Add(salesTargetTeam);
            }

            return salesTargetTeams;
        }

        /// <summary>
        /// 实现逻辑：
        /// 1. 由于需要实现分组合计，所以提前把数据进行分组
        /// 2. 根据分组填充数据，如果对应的项目是空，补充空空数据
        /// 3. 然后根据分组的合计列，显示合计
        /// </summary>
        /// <param name="salesTargetTeams"></param>
        public void Write(List<SalesTargetTeam> salesTargetTeams)
        {
            #region 分组数据
            var groupModel = new GroupModel() { ChildModels = new Dictionary<string, GroupModel>(), RowIndexs = new List<int>() };
            foreach (var teamModel in salesTargetTeams)
            {
                if (!groupModel.ChildModels.ContainsKey(teamModel.Month))
                    groupModel.ChildModels.Add(teamModel.Month, new GroupModel() { ChildModels = new Dictionary<string, GroupModel>(), RowIndexs = new List<int>() });

                var monthModel = groupModel.ChildModels[teamModel.Month]; //月份
                if (!monthModel.ChildModels.ContainsKey(teamModel.RegionDepartmentName))
                    monthModel.ChildModels.Add(teamModel.RegionDepartmentName, new GroupModel() { ChildModels = new Dictionary<string, GroupModel>(), RowIndexs = new List<int>() });

                var area1Model = monthModel.ChildModels[teamModel.RegionDepartmentName]; //区域
                if (!area1Model.ChildModels.ContainsKey(teamModel.AreaDepartmentName))
                    area1Model.ChildModels.Add(teamModel.AreaDepartmentName, new GroupModel() { ChildModels = new Dictionary<string, GroupModel>(), RowIndexs = new List<int>() });

                var area2Model = area1Model.ChildModels[teamModel.AreaDepartmentName]; //大区
                if (!area2Model.ChildModels.ContainsKey(teamModel.BranchDepartmentName))
                    area2Model.ChildModels.Add(teamModel.BranchDepartmentName, new GroupModel() { ChildModels = new Dictionary<string, GroupModel>(), TeamModels = new List<SalesTargetTeam>() });

                var branchModel = area2Model.ChildModels[teamModel.BranchDepartmentName]; //分公司
                branchModel.TeamModels.Add(teamModel);
            }
            #endregion

            #region 项目数据
            var projectInfos = new List<SalesTargetProject>();
            projectInfos.Add(new SalesTargetProject() { ProjectName = "史密斯", ColumnsCount = 1, Sequence = 1 });
            projectInfos.Add(new SalesTargetProject() { ProjectName = "央企常规指标", ColumnsCount = 4, Sequence = 2 });
            projectInfos.Add(new SalesTargetProject() { ProjectName = "政府常规指标", ColumnsCount = 4, Sequence = 4 });
            projectInfos.Add(new SalesTargetProject() { ProjectName = "国网", ColumnsCount = 3, Sequence = 8 });
            projectInfos.Add(new SalesTargetProject() { ProjectName = "南网", ColumnsCount = 3, Sequence = 9 }); 
            #endregion

            var strFile = @"D:\文档\20200423销售指标线上化\业绩指标计划导入模板2.xlsx";
            IWorkbook workbook = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;

            FileStream fs = null;
            var rowIndex = 0; //当前行
            try
            {
                workbook = new XSSFWorkbook(); //2007版本
                sheet = workbook.CreateSheet("Sheet1");//创建一个名称为Sheet1的表

                int columnCount = projectInfos.Sum(o => o.ColumnsCount) + this._columnPrefix + this._columnSuffix; //列数
                #region 设置列头
                row = sheet.CreateRow(rowIndex);//excel第1行设为列头
                rowIndex++;

                var colIndex = 0;
                var colTitles = new List<string>() { "月份", "区域", "大区", "分公司", "团队" };
                #region 设置第一列
                //前置列设置为空
                for (colIndex = 0; colIndex < this._columnPrefix; colIndex++)
                {
                    cell = row.CreateCell(colIndex);
                    cell.SetCellValue("");
                }

                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, this._columnPrefix - 1)); //合计列

                //中间设置项目
                foreach (var projectInfo in projectInfos)
                {
                    for (int j = 0; j < projectInfo.ColumnsCount; j++)
                    {
                        cell = row.CreateCell(colIndex);
                        colIndex++;
                        cell.SetCellValue(j == 0 ? projectInfo.ProjectName : "");
                    }

                    if (projectInfo.ProjectName.Contains("常规"))
                    {
                        colTitles.AddRange(new string[] { "销售额", "毛利率", "毛利额", "新客户数" });
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, colIndex - 4, colIndex - 1)); //合计列
                    }
                    else if (projectInfo.ProjectName.Contains("史密斯"))
                    {
                        colTitles.Add("台数");
                    }
                    else
                    {
                        colTitles.AddRange(new string[] { "销售额", "毛利率", "毛利额" });
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, colIndex - 3, colIndex - 1)); //合计列
                    }
                }

                colTitles.AddRange(new string[] { "销售额", "毛利率", "毛利额", "新客户数" }); //最后合计列

                //后置列设置合计
                while (colIndex < columnCount)
                {
                    cell = row.CreateCell(colIndex);
                    cell.SetCellValue("合计");

                    colIndex++;
                }
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, colIndex - 4, colIndex - 1)); //合计列
                #endregion

                #region 设置第二列
                row = sheet.CreateRow(rowIndex);//excel第2行设为列头
                rowIndex++;

                for (colIndex = 0; colIndex < colTitles.Count; colIndex++)
                {
                    cell = row.CreateCell(colIndex);
                    cell.SetCellValue(colTitles[colIndex]);
                }
                #endregion

                #region 设置第三列
                var rowTotal = sheet.CreateRow(rowIndex);//excel第3行设为总合计
                rowIndex++;

                for (colIndex = 0; colIndex < colTitles.Count; colIndex++)
                {
                    cell = rowTotal.CreateCell(colIndex);
                    if (colIndex == 0)
                        cell.SetCellValue("TOTAL");
                }

                sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, this._columnPrefix - 1)); //合计列
                #endregion
                #endregion

                #region 合计处理
                Action<int, string, List<int>, int, int> sum = (colIndex, calName, colIndexs, colIndexStart, colIndexEnd) =>
                {
                    row = sheet.CreateRow(rowIndex);
                    rowIndex++;

                    for (int j = 0; j < columnCount; j++)
                    {
                        cell = row.CreateCell(j);
                        if (j >= this._columnPrefix)
                        {
                            var colLetter = this.ConvertToLetter(j + 1); //列号起始为1
                            if (colIndexs == null || colIndexs.Count == 0)
                                cell.CellFormula = $"SUM({colLetter}{colIndexStart}:{colLetter}{colIndexEnd})"; //对应团队合计，采用冒号分隔
                            else
                                cell.CellFormula = $"SUM({string.Join(',', colIndexs.Select(o => colLetter + o))})"; //对于其他合计，采用逗号分隔

                            if (colTitles[j] == this._profitRate)
                                cell.CellFormula = $"{this.ConvertToLetter(j + 2)}{rowIndex}/{this.ConvertToLetter(j)}{rowIndex}"; //毛利采用毛利额/销售额
                        }
                        else
                        {
                            cell.SetCellValue(j == colIndex ? calName : ""); //在第几列设置合计标题
                        }
                    }
                };
                #endregion

                foreach (var monthModel in groupModel.ChildModels)
                {
                    foreach (var area1Model in monthModel.Value.ChildModels)
                    {
                        foreach (var area2Model in area1Model.Value.ChildModels)
                        {
                            #region 分公司维度
                            foreach (var brachModel in area2Model.Value.ChildModels)
                            {
                                var rowIndexStart = rowIndex + 1; //分组开始行，注意Excel的序号从1开始
                                foreach (var teamModel in brachModel.Value.TeamModels)
                                {
                                    #region 新增并设置固定列
                                    row = sheet.CreateRow(rowIndex);
                                    rowIndex++;

                                    row.CreateCell(0).SetCellValue(teamModel.Month);
                                    row.CreateCell(1).SetCellValue(teamModel.RegionDepartmentName);
                                    row.CreateCell(2).SetCellValue(teamModel.AreaDepartmentName);
                                    row.CreateCell(3).SetCellValue(teamModel.BranchDepartmentName);
                                    row.CreateCell(4).SetCellValue(teamModel.TeamDepartmentName);
                                    #endregion

                                    #region 数据列
                                    var arrSalesAmountIndexs = new List<int>();
                                    var arrProfitAmountIndexs = new List<int>();
                                    var arrNewCustomerCountIndexs = new List<int>();
                                    colIndex = this._columnPrefix; //设置开始列
                                    foreach (var projectInfo in projectInfos)
                                    {
                                        var salesTargetByTeam = teamModel.ProjectModels.FirstOrDefault(o => o.ProjectName == projectInfo.ProjectName);
                                        if (projectInfo.ProjectName.Contains(this._normalKeyWord))
                                        {
                                            arrSalesAmountIndexs.Add(colIndex);
                                            cell = row.CreateCell(colIndex); colIndex++;
                                            if (salesTargetByTeam != null)
                                                cell.SetCellValue(Convert.ToDouble(salesTargetByTeam.SalesAmount));

                                            cell = row.CreateCell(colIndex); colIndex++;
                                            if (salesTargetByTeam != null)
                                                cell.SetCellValue(Convert.ToDouble(salesTargetByTeam.ProfitRate));

                                            arrProfitAmountIndexs.Add(colIndex);
                                            cell = row.CreateCell(colIndex); colIndex++;
                                            cell.CellFormula = $"{this.ConvertToLetter(colIndex - 2) + rowIndex.ToString()}*{this.ConvertToLetter(colIndex - 1) + rowIndex.ToString()}";

                                            arrNewCustomerCountIndexs.Add(colIndex);
                                            cell = row.CreateCell(colIndex); colIndex++;
                                            if (salesTargetByTeam != null)
                                                cell.SetCellValue(salesTargetByTeam.NewCustomerCount);
                                        }
                                        else if (projectInfo.ProjectName.Contains(this._smithKeyWord))
                                        {
                                            cell = row.CreateCell(colIndex); colIndex++;
                                            if (salesTargetByTeam != null)
                                                cell.SetCellValue(salesTargetByTeam.SalesCount);
                                        }
                                        else
                                        {
                                            arrSalesAmountIndexs.Add(colIndex);
                                            cell = row.CreateCell(colIndex); colIndex++;
                                            if (salesTargetByTeam != null)
                                                cell.SetCellValue(Convert.ToDouble(salesTargetByTeam.SalesAmount));

                                            cell = row.CreateCell(colIndex); colIndex++;
                                            if (salesTargetByTeam != null)
                                                cell.SetCellValue(Convert.ToDouble(salesTargetByTeam.ProfitRate));

                                            arrProfitAmountIndexs.Add(colIndex);
                                            cell = row.CreateCell(colIndex); colIndex++;
                                            cell.CellFormula = $"{this.ConvertToLetter(colIndex - 2) + rowIndex.ToString()}*{this.ConvertToLetter(colIndex - 1) + rowIndex.ToString()}";
                                        }
                                    }
                                    #endregion

                                    #region 统计列
                                    cell = row.CreateCell(colIndex); colIndex++;
                                    cell.CellFormula = $"SUM({string.Join(',', arrSalesAmountIndexs.Select(o => this.ConvertToLetter(o + 1) + rowIndex))})";

                                    cell = row.CreateCell(colIndex); colIndex++;
                                    cell.SetCellValue("");

                                    cell = row.CreateCell(colIndex); colIndex++;
                                    cell.CellFormula = $"SUM({string.Join(',', arrProfitAmountIndexs.Select(o => this.ConvertToLetter(o + 1) + rowIndex))})";

                                    cell = row.CreateCell(colIndex); colIndex++;
                                    cell.CellFormula = $"SUM({string.Join(',', arrNewCustomerCountIndexs.Select(o => this.ConvertToLetter(o + 1) + rowIndex))})";
                                    #endregion
                                }

                                sum(3, brachModel.Key, null, rowIndexStart, rowIndex); //分公司合计在第4行
                                area2Model.Value.RowIndexs.Add(rowIndex); //把当前列，也就是分公司合计列，加入到上级集合中，方便做合计
                            }
                            #endregion

                            sum(2, area2Model.Key, area2Model.Value.RowIndexs, 0, 0); //大区合计在第3行
                            area1Model.Value.RowIndexs.Add(rowIndex);
                        }

                        sum(1, area1Model.Key, area1Model.Value.RowIndexs, 0, 0); //区域合计在第2行
                        monthModel.Value.RowIndexs.Add(rowIndex);
                    }

                    sum(0, monthModel.Key, monthModel.Value.RowIndexs, 0, 0); //月份合计在第1行
                    groupModel.RowIndexs.Add(rowIndex);
                }

                #region 把最后的汇总列，赋值到第三列
                for (int j = this._columnPrefix; j < columnCount; j++)
                {
                    var colLetter = this.ConvertToLetter(j + 1); //列号起始为1
                    cell = rowTotal.GetCell(j);
                    cell.CellFormula = $"SUM({string.Join(',', groupModel.RowIndexs.Select(o => colLetter + o))})";
                    if (colTitles[j] == this._profitRate)
                        cell.CellFormula = $"{this.ConvertToLetter(j + 2)}{rowTotal.RowNum + 1}/{this.ConvertToLetter(j)}{rowTotal.RowNum + 1}"; //毛利采用毛利额/销售额
                }
                #endregion

                using (fs = File.OpenWrite(strFile))
                {
                    workbook.Write(fs);//向打开的这个xls文件中写入数据
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }

        private bool SetData(SalesTargetByTeam projectModel, DataRow dr, int colIndex, ProjectTypeEnum projectType)
        {
            if (projectType == ProjectTypeEnum.Smith)
            {
                if (string.IsNullOrWhiteSpace(dr[colIndex].ToString()))
                    return false;
                else
                    projectModel.SalesCount = Convert.ToInt32(dr[colIndex].ToString().Trim()); //台数
            }
            else
            {
                if (string.IsNullOrWhiteSpace(dr[colIndex].ToString()))
                    return false;
                else
                    projectModel.SalesAmount = Convert.ToDecimal(dr[colIndex].ToString().Trim()); //销售额

                if (string.IsNullOrWhiteSpace(dr[colIndex + 1].ToString()))
                    return false;
                else
                    projectModel.ProfitRate = Convert.ToDecimal(dr[colIndex + 1].ToString().Trim()); //毛利率

                if (string.IsNullOrWhiteSpace(dr[colIndex + 2].ToString()))
                    return false;
                else
                    projectModel.ProfitAmount = Convert.ToDecimal(dr[colIndex + 2].ToString().Trim()); //毛利额

                if (projectType == ProjectTypeEnum.Normal)
                {
                    if (string.IsNullOrWhiteSpace(dr[colIndex + 3].ToString()))
                        return false;
                    else
                        projectModel.NewCustomerCount = Convert.ToInt32(dr[colIndex + 3].ToString().Trim()); //新客户数
                }
            }

            return true;
        }

        private SalesTargetTeam Mapping(DataRow dr)
        {
            var teamModel = new SalesTargetTeam();
            teamModel.Month = dr[0].ToString().Trim(); //月份
            teamModel.RegionDepartmentName = dr[1].ToString().Trim(); //区域
            teamModel.AreaDepartmentName = dr[2].ToString().Trim(); //大区
            teamModel.BranchDepartmentName = dr[3].ToString().Trim(); //分公司
            teamModel.TeamDepartmentName = dr[4].ToString().Trim(); //团队
            //teamModel.CustomerClassName = dr[5].ToString().Trim(); //行业
            teamModel.ProjectModels = new List<SalesTargetByTeam>();

            return teamModel;
        }

        /// <summary>
        /// 列的序号转换为Excel的字母字符
        /// </summary>
        public string ConvertToLetter(int colomnIndex)
        {
            int a;
            int b;

            var convertToLetter = "";
            while (colomnIndex > 0)
            {
                a = Convert.ToInt32((colomnIndex - 1) / 26);
                b = (colomnIndex - 1) % 26;
                convertToLetter = Convert.ToChar(b + 65).ToString() + convertToLetter;
                colomnIndex = a;
            }

            return convertToLetter;
        }

        private enum ProjectTypeEnum
        {
            Normal = 1,
            Major = 2,
            Smith = 3
        }
    }
}
