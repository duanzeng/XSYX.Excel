using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Data;

namespace XSYX.Excel.Tests
{
    [TestClass]
    public class ExportTests
    {
        [TestMethod]
        public void EEPlus_List_Test()
        {
            var list = new List<dynamic>();
            string[] columnNames = new string[] { "≤÷ø‚ID", "≤÷ø‚√˚≥∆" };
            var count = 1000000;
            for (int i = 0; i < count; i++)
            {
                list.Add(new
                {
                    Id = i,
                    Name = $"≥§…≥{i}"
                });
            }
            ExportHelper.ExportExcel(list, columnNames, "EEPlus_List.xlsx");
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void EPPlus_DataTable_Test()
        {
            var list = new List<dynamic>();
            string[] columnNames = new string[] { "≤÷ø‚ID", "≤÷ø‚√˚≥∆" };
            var count = 1000000;
            var dt = new DataTable();
            dt.Columns.Add(new DataColumn("id", typeof(int)));
            dt.Columns.Add(new DataColumn("name", typeof(string)));
            for (int j = 0; j < count; j++)
            {
                var dr = dt.NewRow();
                dr["id"] = j;
                dr["name"] = $"≥§…≥{j}";
                dt.Rows.Add(dr);
            }
            ExportHelper.ExportExcel(dt, columnNames, "EPPlus.DataTable.xlsx");

            Assert.IsTrue(true);
        }
    }
}
