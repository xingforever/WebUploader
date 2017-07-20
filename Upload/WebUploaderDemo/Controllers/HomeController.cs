using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebUploaderDemo.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Excel()
        {
            return View();
        }

        public ActionResult Image() {

            return View();
        }

        public ActionResult  UploadExcel(HttpPostedFileBase file)
        {
            string fileName = string .Empty;//文件名
            string dir = string.Empty;//文件相对路径
            if (file != null)
            {
                fileName = Path.GetFileName(file.FileName);
                string fileExt = Path.GetExtension(fileName);//文件扩展名
                //if (fileExt == ".xls" || fileExt == ".xlsx")
                //{
                    //文件存储在Data/Excel/年/月/日  下  
                     dir = "\\Data\\Excel\\" + DateTime.Now.Year + "\\" + DateTime.Now.Month + "\\" + DateTime.Now.Day;
                   string newfilepath = System.Web.HttpContext.Current.Server.MapPath(dir);//获取物理路径,很重要
                    if (!Directory.Exists(newfilepath)) //   创建文件夹
                    {
                        Directory.CreateDirectory(newfilepath);
                    }
                    string path = newfilepath + "\\" + fileName;    //真实地址                
                    file.SaveAs(path);//存储文件
                //}

            }
            return Json(new
            {
                jsonrpc = "2.0",             
                filePath = dir + "/" + fileName   //相对位置
            });


        }


        public ActionResult UpLoadImage(string id, string name, string type, string lastModifiedDate, int size, HttpPostedFileBase file)
        {
            string fileName = string.Empty;//文件名
            string dir = string.Empty;//文件相对路径
            if (file!=null)
            {
               
                fileName = file.FileName;
                dir = "\\Data\\Images\\" + DateTime.Now.Year + "\\" + DateTime.Now.Month + "\\" + DateTime.Now.Day;
                string fileExt = Path.GetExtension(fileName);//文件扩展名
                string newfilepath = System.Web.HttpContext.Current.Server.MapPath(dir);//获取物理路径,很重要
                if (!Directory.Exists(newfilepath)) //   创建文件夹
                {
                    Directory.CreateDirectory(newfilepath);
                }
                if (Request.Files.Count == 0)
                {
                    return Json(new { jsonrpc = 2.0, error = new { code = 102, message = "保存失败" }, id = "id" });
                }                              
               string   filePathName = newfilepath + "\\"+ fileName;
                
                file.SaveAs(filePathName);
            }           

            return Json(new
            {
                jsonrpc = "2.0",
                id = id,
                filePath = dir + "/" + fileName   //相对位置
            });

        }


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}