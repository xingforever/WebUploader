using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebUploaderDemo.Common;

namespace WebUploaderDemo.Controllers
{
    public class DisplayExcelController : Controller
    {
        // GET: DisplayExcel
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExcelInfo() {

            return View();
        }


        public ActionResult GetCityInfoList()

        {
            //获取页码
            int pageIndex = Request["page"] != null ? int.Parse(Request["page"]) : 1;
            int pageSize = Request["rows"] != null ? int.Parse(Request["rows"]) : 5;
            
            //Excel地址
            string Adress = Request["Adress"];
            string newfilepath = System.Web.HttpContext.Current.Server.MapPath(Adress);
            //读取Excel 
            DataTable dt=  NPOIHelper.ReadExcel(newfilepath);

            var count = dt.Rows.Count;
            //将DataTable 转换为Json
            string json = DatableToJsonHelp.DataTableToJson(dt);
            return Json(new { row=json,total= count}, JsonRequestBehavior.AllowGet);


            ////获取第一行表头
            //IRow headRow = sheet.GetRow(0);
            ////列数
            //int columnCount = headRow.LastCellNum;//LastCellNum=PhysicalNumberOfCells
            //int rowCount = sheet.LastRowNum;//LastRowNum=PhysicalNumberOfCellsRow-1
            //                                //创建DataTable的表头
            //for (int i = headRow.FirstCellNum; i < columnCount; i++)
            //{
            //    DataColumn dc = new DataColumn(headRow.GetCell(i).StringCellValue.ToString());
            //    dt.Columns.Add(dc);
            //}


        }


        public string  DisplayExcelInfo() {


            string Adress = Request["Adress"];
            string newfilepath = System.Web.HttpContext.Current.Server.MapPath(Adress);
            //读取Excel 
            DataTable dt = NPOIHelper.ExcelToDataTable(newfilepath,true);

            var count = dt.Rows.Count;
            //将DataTable 转换为Json
            string json = DatableToJsonHelp.DataTableToJson(dt);//多了双引号

            return json;
            // Json(json, JsonRequestBehavior.AllowGet,);
            ////return Json(new { row = json, total = count }, JsonRequestBehavior.AllowGet);

            //return Json("fsd");


        }
    }
}