using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;

namespace XSYX.Excel
{
    /// <summary>
    /// 导出帮助类
    /// </summary>
    public class ExportHelper
    {
        /// <summary>
        /// sheet页最大行数
        /// </summary>
        public static readonly int sheetMaxRows = 1048575;

        /// <summary>
        /// 导出Excel文件,超过sheet页行数会自动增加sheet
        /// </summary>
        /// <param name="dt">数据集合</param>
        /// <param name="columnNames">列名称</param>
        /// <param name="fileName">文件名称，需要以xlsx结尾</param>
        /// <param name="sheetName">sheet名称</param>
        public static void ExportExcel(DataTable dt, string[] columnNames, string fileName, string sheetName = "Sheet")
        {
            using (var fileStream = File.Create(fileName))
            {
                var excelBytes = ExportExcelBytes(dt, columnNames, sheetName);
                fileStream.Write(excelBytes, 0, excelBytes.Length);
            }
        }

        /// <summary>
        /// 导出Excel,超过sheet页行数会自动增加sheet
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="columnNames"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static byte[] ExportExcelBytes(DataTable dt, string[] columnNames, string sheetName = "Sheet")
        {
            using (var excel = new ExcelPackage())
            {
                var sheetCount = Convert.ToInt32(dt.Rows.Count / sheetMaxRows) + 1;//根据总数和sheet最大行数计算sheet页数量
                for (int o = 1; o <= sheetCount; o++)
                {
                    var sheet = excel.Workbook.Worksheets.Add($"{sheetName ?? "Sheet"}{o}");
                    if (columnNames.Length > 0)
                    {
                        for (int k = 0; k < columnNames.Length; k++)
                        {
                            sheet.Cells[1, k + 1].Value = columnNames[k].Trim();
                            sheet.Cells[1, k + 1].Style.Font.Bold = true;//字体为粗体
                        }
                    }

                    var currCount = dt.Rows.Count - (o - 1) * sheetMaxRows;//剩余行数
                    var sheetRows = currCount > sheetMaxRows ? sheetMaxRows : currCount;//当前sheet可显示行数

                    for (int i = 0; i < sheetRows; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            //!!注意NPPlus下标是从1开始的,要计算列头
                            sheet.Cells[i + 1 + 1, j + 1].Value = dt.Rows[(o - 1) * sheetMaxRows + i][j];
                        }
                    }
                }
                return excel.GetAsByteArray();
            }
        }

        /// <summary>
        /// 导出Excel文件,超过sheet页行数会自动增加sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList">数据集合</param>
        /// <param name="columnNames">列头</param>
        /// <param name="fileName">文件名称，需要以xlsx结尾</param>
        /// <param name="sheetName">sheet名称</param>
        public static void ExportExcel<T>(List<T> dataList, string[] columnNames, string fileName, string sheetName = "Sheet")
        {
            using (var fileStream = File.Create(fileName))
            {
                var excelBytes = ExportExcelBytes(dataList, columnNames, sheetName);
                fileStream.Write(excelBytes, 0, excelBytes.Length);
            }
        }

        /// <summary>
        /// 导出Excel,超过sheet页行数会自动增加sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList">数据集合</param>
        /// <param name="columnNames">列头集合</param>
        /// <param name="sheetName">sheet名称</param>
        /// <returns>文件字节</returns>
        public static byte[] ExportExcelBytes<T>(List<T> dataList, string[] columnNames, string sheetName = "Sheet")
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                //sheet页根据最大行数分多个sheet页
                var sheetCount = Convert.ToInt32(dataList.Count / sheetMaxRows) + 1;
                for (int o = 1; o <= sheetCount; o++)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"{sheetName}{o}");

                    //添加表头
                    int column = 1;
                    if (columnNames.Length > 0)
                    {
                        foreach (string cn in columnNames)
                        {
                            worksheet.Cells[1, column].Value = cn.Trim();
                            worksheet.Cells[1, column].Style.Font.Bold = true;//字体为粗体
                                                                              //worksheet.Cells[1, column].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;//水平居中
                            column++;
                        }
                    }

                    int row = 2;    //数据列从1开始+列头=2
                    var currCount = dataList.Count - (o - 1) * sheetMaxRows;    //剩余数量=总数减去sheet最大行数
                    var sheetRows = currCount > sheetMaxRows ? sheetMaxRows : currCount;    //剩余数量>最大行数则取最大行数，否则取当前行数
                    foreach (T ob in dataList.GetRange((o - 1) * sheetMaxRows, sheetRows))  //
                    {
                        int col = 1;
                        foreach (PropertyInfo property in ob.GetType().GetRuntimeProperties())
                        {
                            worksheet.Cells[row, col].Value = property.GetValue(ob);
                            col++;
                        }
                        row++;
                    }
                }
                return package.GetAsByteArray();
            }
        }
    }
}
