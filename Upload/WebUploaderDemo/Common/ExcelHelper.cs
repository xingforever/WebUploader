using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace WebUploaderDemo.Common
{
    public  static class ExcelHelper
    {

        /// <summary>
        /// 读取excel
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="flag">错误标志</param>
        /// <returns>数据表</returns>
        public static System.Data.DataTable ReadExcel(string path)
        {
            System.Data.DataTable dt = new System.Data.DataTable();//系统表
            Application excel = new Application();
            excel.Visible = false;
            excel.UserControl = true;
            var missming = System.Reflection.Missing.Value;
            Workbook workbook = excel.Application.Workbooks.Open(path, missming, true, missming, missming, missming, missming, missming, missming, missming, missming, missming, missming, missming, missming);
            Worksheet worksheet = (Worksheet)workbook.Worksheets.get_Item(1);//默认第一张表
            int rowint = worksheet.UsedRange.Cells.Rows.Count;
            int columnint = worksheet.UsedRange.Cells.Columns.Count;//一共多少列
            int A = 65;//从A开始
            byte[] array = new byte[1];  ///excel 是从A开始  一直往前加  
            A += columnint - 1;//
            array[0] = (byte)A;
            string s = Convert.ToString(System.Text.Encoding.ASCII.GetString(array));
            Range rng = worksheet.Cells.get_Range("A1", s + rowint);//从A读取行范围
            for (int i = 0; i < columnint; i++)
            {
                dt.Columns.Add();
            }
            for (int i = 0; i < rowint; i++)
            {
                dt.Rows.Add();
            }

            object[,] arrayItem = (object[,])rng.Value2;
            for (int i = 1; i < rowint + 1; i++)
            {
                for (int j = 1; j < columnint + 1; j++)
                {
                    try
                    {
                        dt.Rows[i - 1][j - 1] = arrayItem[i, j];

                    }
                    catch (Exception)
                    {

                        continue;
                    }

                }

            }
            Process[] pros = Process.GetProcessesByName("excel");//关闭Excel程序
            foreach (var pro in pros)
            {
                pro.Kill();
            }
            GC.Collect();
            return dt;

        }
        /// <summary>
        /// 保存excel
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="dt">数据表</param>
        public static void SaveExcel(string path, System.Data.DataTable dt)
        {
            Application excel = new Application();
            Workbooks workbooks = excel.Workbooks;
            Workbook workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet worksheet = workbook.Sheets[1];
            string[] headeetxt = new string[] { "不规则点号", "不规则点x", "不规则点y", "绝对高", "三角网编号", "顶点编号p1", "顶点编号p2", "顶点编号p3", "体积" };
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = headeetxt[i];
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dt.Rows[i][j];
                }
            }

            worksheet.UsedRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.UsedRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.UsedRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            workbook.Saved = true;
            workbook.SaveAs(path);
            Process[] pros = Process.GetProcessesByName("excel");
            foreach (var pro in pros)
            {
                pro.Kill();
            }


        }
    }
}