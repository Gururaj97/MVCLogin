﻿using MVCLogin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using Table = MVCLogin.Models.Table;

namespace MVCLogin.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Autherize(MVCLogin.Models.Table userModel)
        {
            using(LoginEntities db = new LoginEntities())
            {
                var userDetails = db.Tables.Where(x => x.UserName == userModel.UserName && x.PassWord == userModel.PassWord).FirstOrDefault();
                if(userDetails == null)
                {
                    userModel.LoginErrorMessage = "Wrong username or password";
                    return View("Index",userModel);
                }
                else
                {
                    Session["userid"] = userDetails.Id;
                    Session["userName"] = userDetails.UserName;
                    return RedirectToAction("Index", "Home");
                }
            }
         
        }
        public ActionResult LogOut()
        {
            int userid = (int)Session["userid"];
            Session.Abandon();
            return RedirectToAction("Index", "Login");
        }
        
        //public ActionResult ImportExcelsheet(HttpPostedFileBase excelfile)
        //{
        //    try
        //    {
        //        if (excelfile == null || excelfile.ContentLength == 0)
        //        {
        //            ViewBag.Error = "Please select a excel file<br>";
        //            return View("Index");
        //        }
        //        else if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
        //        {
        //            string path = Server.MapPath("~/content/" + excelfile.FileName);
        //            if (System.IO.File.Exists(path))
        //            {
        //                System.IO.File.Delete(path);
        //            }
        //            else
        //            {
        //                excelfile.SaveAs(path);
        //            }


        //            Excel.Application application = new Excel.Application();
        //            Excel.Workbook workbook = application.Workbooks.Open(path);
        //            Excel.Worksheet worksheet = workbook.ActiveSheet;
        //            Excel.Range range = worksheet.UsedRange;
        //            List<Table> listproducts = new List<Table>();
        //            for (int row = 1; row < range.Rows.Count; row++)
        //            {
        //                Table p = new Table();
        //                p.SNo = int.Parse(((Excel.Range)range.Cells[row, 1]).Text);
        //                p.Name = ((Excel.Range)range.Cells[row, 2]).Text;
        //                p.Age = int.Parse(((Excel.Range)range.Cells[row, 3]).Text);
        //                listproducts.Add(p);
        //            }
        //            ViewBag.ListProducts = listproducts;
        //            return View("Sucess");
        //        }
        //        else
        //        {
        //            ViewBag.Error = "File type is incorrect<br>";
        //            return View("Index");
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //        Response.Redirect(ex.Message);
        //    }
        //    return View();
        //}
    }
   
}