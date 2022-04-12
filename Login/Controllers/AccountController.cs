using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Login.Models;
using System.Data.SqlClient;
using System.Data;
using OfficeOpenXml;
using System.Drawing;
using System.IO;
using Microsoft.Win32;
using System.Text;
using System.Web.UI;
using Microsoft.SharePoint.Client.Search.Query;
using System.Windows.Controls;

namespace Login.Controllers
{

    public class AccountController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();
        SqlDataReader dr;

        [HttpGet]

        public ActionResult Login()
        {
            
            return View();
        }
        void connectionString()
        {
            //con.ConnectionString = "Data Source=PERSONALSRV-KAR\\SQL2016;Initial Catalog=of1; Persist Security Info=True;User ID=sa;Password=12341234; MultipleActiveResultSets=True;";
            con.ConnectionString = "Data Source=(local);Initial Catalog=of1;Integrated Security=True;";
            //con.ConnectionString = "Data Source=(localdb)\\mssqllocaldb;Initial Catalog=of1;Integrated Security=True;";

        }
        [HttpPost]
        public ActionResult verify(UserAccounts acc)
        {
            DataTable dt = new DataTable();
            connectionString();
            con.Open();
            com.Connection = con;
            var ff = acc.Username;
            TempData["idd"] = ff;
            TempData.Keep("idd");
         
            //usertable
            com.CommandText = "SELECT * FROM [of1].[dbo].[User] where Username='" + acc.Username + "' and Pass='" + acc.Pass + "'";
           //
            dr = com.ExecuteReader();
            if (dr.Read())
            {
                con.Close();
                getDropDown();
                return View("getDropDown");
            }
            else
            {
                con.Close();
                return View("Error");
            }

        }
        //conectionchange
        Entities3 db = new Entities3();
        //
        public ActionResult getDropDown()
        {
            List<year> yearlist = db.years.ToList();
            List<month> monthlist = db.months.ToList();
            ViewBag.yearlist = new SelectList(yearlist, "id", "yearname", "yearnum");
            ViewBag.monthlist = new SelectList(monthlist, "id", "monthname", "monthnum");

            return View();
        }

        [HttpPost]
        public void submit( Yearmonth yearmonthh)
        {

            //conectionchange
            Entities3 db = new Entities3();
            //
            ViewBag.mmm = yearmonthh.selmonthId;
            ViewBag.yyy = yearmonthh.selyearId;
            var t = yearmonthh.selyearId;
            var t2 = yearmonthh.selmonthId;

           string yenum= (from years in db.years
            where
              years.id == t
             select new
            {
                yenum=years.yearnum
             }).ToList().FirstOrDefault().yenum;
            string monum = (from months in db.months
                       where
                         months.id == t2
                       select new
                       {
                           monthnum = months.monthnum
                       }).ToList().FirstOrDefault().monthnum;
            string conc = yenum + monum;
            ViewBag.uu= conc;
            var Userid = TempData["idd"];
            TempData.Keep("idd");
            DataTable dt = new DataTable();
            //usertable
            var compst = (from Users in db.Users
                          where
                            Users.Username == Userid
                          select new
                          {
                              Users.CompanyStatus
                          }).FirstOrDefault().ToString();
            //
            TempData["compst"] = compst;
            TempData.Keep("compst");
            //empltable+ usertable
            if (compst.Contains("1"))
            {
                List<Employee> query1 = (from Employees in db.Employees
                                        where
                                              (from Users in db.Users
                                               where
                                     Users.Username == Userid
                                               select new
                                               {
                                                   Users.CompanyCode
                                               }).Contains(new { CompanyCode = Employees.COMPANY_CODE }) &&
              Employees.SALARY_YYMM.Contains(conc)
                                        select Employees).ToList();
               
                //return View(query1);
                ExcelPackage p1 = new ExcelPackage();
                ExcelWorksheet ew = p1.Workbook.Worksheets.Add("Report");
                ew.Cells["A2"].Value = "Report";
                ew.Cells["B2"].Value = "Report1";
                ew.Cells["A3"].Value = "Date";
                ew.Cells["B3"].Value = string.Format("{0:dd MMMM yyyy} at {0:H: mm tt}", DateTimeOffset.Now);
                //emptable
                ew.Cells["A6"].Value = "REG_NO";
                ew.Cells["B6"].Value = "COMPANY_CODE";
                ew.Cells["C6"].Value = "MG_CODE";
                ew.Cells["D6"].Value = "PYRLCMP_CODE";
                ew.Cells["E6"].Value = "NATIONAL_NO";
                int rowStart = 7;
                foreach (var item in query1)
                {
                    ew.Cells[String.Format("A{0}", rowStart)].Value = item.REG_NO;
                    ew.Cells[String.Format("B{0}", rowStart)].Value = item.COMPANY_CODE;
                    ew.Cells[String.Format("C{0}", rowStart)].Value = item.MG_CODE;
                    ew.Cells[String.Format("D{0}", rowStart)].Value = item.PYRLCMP_CODE;
                    ew.Cells[String.Format("E{0}", rowStart)].Value = item.NATIONAL_NO;
                    rowStart++;
                }
                //
                ew.Cells["A:AZ"].AutoFitColumns();
                string filename = "Results_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                Response.Clear();
                Response.ContentType = "application/vnd.ms-excel";
                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", "attachment; filename=Report.xlsx");
                Response.ContentEncoding = Encoding.UTF8;
                StringWriter stringWriter = new StringWriter();
                HtmlTextWriter hw = new HtmlTextWriter(stringWriter);
                Response.Write(stringWriter.ToString());
                Response.BinaryWrite(p1.GetAsByteArray());
                Response.End();
            }

            else
            {
                //empltable+ usertable
                List<Employee> query2 = (from Employees in db.Employees
                                        where
                                              (from Users in db.Users
                                               where
                                                Users.Username == Userid
                                               select new
                                               {
                                                   Users.PayrollCode
                                               }).Contains(new { PayrollCode = Employees.PYRLCMP_CODE }) &&
              Employees.SALARY_YYMM.Contains(conc)
                                        select Employees).ToList();
                //
                //return View(query2);
                ExcelPackage p1 = new ExcelPackage();
                ExcelWorksheet ew = p1.Workbook.Worksheets.Add("Report");
                ew.Cells["A2"].Value = "Report";
                ew.Cells["B2"].Value = "Report1";
                ew.Cells["A3"].Value = "Date";
                ew.Cells["B3"].Value = string.Format("{0:dd MMMM yyyy} at {0:H: mm tt}", DateTimeOffset.Now);
                //emptable
                ew.Cells["A6"].Value = "REG_NO";
                ew.Cells["B6"].Value = "COMPANY_CODE";
                ew.Cells["C6"].Value = "MG_CODE";
                ew.Cells["D6"].Value = "PYRLCMP_CODE";
                ew.Cells["E6"].Value = "NATIONAL_NO";
                int rowStart = 7;
                foreach (var item in query2)
                {
                    ew.Cells[String.Format("A{0}", rowStart)].Value = item.REG_NO;
                    ew.Cells[String.Format("B{0}", rowStart)].Value = item.COMPANY_CODE;
                    ew.Cells[String.Format("C{0}", rowStart)].Value = item.MG_CODE;
                    ew.Cells[String.Format("D{0}", rowStart)].Value = item.PYRLCMP_CODE;
                    ew.Cells[String.Format("E{0}", rowStart)].Value = item.NATIONAL_NO;
                    rowStart++;
                }
                //
                ew.Cells["A:AZ"].AutoFitColumns();
                string filename = "Results_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                Response.Clear();
                Response.ContentType = "application/vnd.ms-excel";
                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", "attachment; filename=Report.xlsx");
                Response.ContentEncoding = Encoding.UTF8;
                StringWriter stringWriter = new StringWriter();
                HtmlTextWriter hw = new HtmlTextWriter(stringWriter);
                Response.Write(stringWriter.ToString());
                Response.BinaryWrite(p1.GetAsByteArray());
                Response.End();
            }
            ViewBag.yy = yenum;
            ViewBag.yy2 = monum;
            //return View();
        }

        public void submit2(Yearmonth yearmonthh)
        {

            //conectionchange
            Entities3 db = new Entities3();
            //
            ViewBag.mmm = yearmonthh.selmonthId;
            ViewBag.yyy = yearmonthh.selyearId;
            var t = yearmonthh.selyearId;
            var t2 = yearmonthh.selmonthId;

            string yenum = (from years in db.years
                            where
                              years.id == t
                            select new
                            {
                                yenum = years.yearnum
                            }).ToList().FirstOrDefault().yenum;
            string monum = (from months in db.months
                            where
                              months.id == t2
                            select new
                            {
                                monthnum = months.monthnum
                            }).ToList().FirstOrDefault().monthnum;
            string conc = yenum + monum;
            ViewBag.uu = conc;

            var Userid = TempData["idd"];
            TempData.Keep("idd");
            DataTable dt = new DataTable();
            //usertable
            var compst = (from Users in db.Users
                          where
                            Users.Username == Userid
                          select new
                          {
                              Users.CompanyStatus
                          }).FirstOrDefault().ToString();
            //
            TempData["compst"] = compst;
            TempData.Keep("compst");

            //empltable+ usertable
            if (compst.Contains("1"))
            {
                List<Employee> query1 = (from Employees in db.Employees
                                         where
                                               (from Users in db.Users
                                                where
                                      Users.Username == Userid
                                                select new
                                                {
                                                    Users.CompanyCode
                                                }).Contains(new { CompanyCode = Employees.COMPANY_CODE }) &&
               Employees.SALARY_YYMM.Contains(conc)
                                         select Employees).ToList();

                //return View(query1);
                ExcelPackage p1 = new ExcelPackage();
                ExcelWorksheet ew = p1.Workbook.Worksheets.Add("Report");
                ew.Cells["A2"].Value = "Report";
                ew.Cells["B2"].Value = "Report1";
                ew.Cells["A3"].Value = "Date";
                ew.Cells["B3"].Value = string.Format("{0:dd MMMM yyyy} at {0:H: mm tt}", DateTimeOffset.Now);
                //emptable
                ew.Cells["A6"].Value = "REG_NO";
                ew.Cells["B6"].Value = "COMPANY_CODE";
                ew.Cells["C6"].Value = "MG_CODE";
                ew.Cells["D6"].Value = "PYRLCMP_CODE";
                ew.Cells["E6"].Value = "NATIONAL_NO";
                int rowStart = 7;
                foreach (var item in query1)
                {
                    ew.Cells[String.Format("A{0}", rowStart)].Value = item.REG_NO;
                    ew.Cells[String.Format("B{0}", rowStart)].Value = item.COMPANY_CODE;
                    ew.Cells[String.Format("C{0}", rowStart)].Value = item.MG_CODE;
                    ew.Cells[String.Format("D{0}", rowStart)].Value = item.PYRLCMP_CODE;
                    ew.Cells[String.Format("E{0}", rowStart)].Value = item.NATIONAL_NO;
                    rowStart++;
                }
                //
                ew.Cells["A:AZ"].AutoFitColumns();
                string filename = "Results_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                Response.Clear();
                Response.ContentType = "application/vnd.ms-excel";
                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", "attachment; filename=Report.xlsx");
                Response.ContentEncoding = Encoding.UTF8;
                StringWriter stringWriter = new StringWriter();
                HtmlTextWriter hw = new HtmlTextWriter(stringWriter);
                Response.Write(stringWriter.ToString());
                Response.BinaryWrite(p1.GetAsByteArray());
                Response.End();
            }

            else
            {
                //empltable+ usertable
                List<Employee> query2 = (from Employees in db.Employees
                                         where
                                               (from Users in db.Users
                                                where
                                                 Users.Username == Userid
                                                select new
                                                {
                                                    Users.PayrollCode
                                                }).Contains(new { PayrollCode = Employees.PYRLCMP_CODE }) &&
               Employees.SALARY_YYMM.Contains(conc)
                                         select Employees).ToList();
                //
                //return View(query2);
                ExcelPackage p1 = new ExcelPackage();
                ExcelWorksheet ew = p1.Workbook.Worksheets.Add("Report");
                ew.Cells["A2"].Value = "Report";
                ew.Cells["B2"].Value = "Report1";
                ew.Cells["A3"].Value = "Date";
                ew.Cells["B3"].Value = string.Format("{0:dd MMMM yyyy} at {0:H: mm tt}", DateTimeOffset.Now);
                //emptable
                ew.Cells["A6"].Value = "REG_NO";
                ew.Cells["B6"].Value = "COMPANY_CODE";
                ew.Cells["C6"].Value = "MG_CODE";
                ew.Cells["D6"].Value = "PYRLCMP_CODE";
                ew.Cells["E6"].Value = "NATIONAL_NO";
                int rowStart = 7;
                foreach (var item in query2)
                {
                    ew.Cells[String.Format("A{0}", rowStart)].Value = item.REG_NO;
                    ew.Cells[String.Format("B{0}", rowStart)].Value = item.COMPANY_CODE;
                    ew.Cells[String.Format("C{0}", rowStart)].Value = item.MG_CODE;
                    ew.Cells[String.Format("D{0}", rowStart)].Value = item.PYRLCMP_CODE;
                    ew.Cells[String.Format("E{0}", rowStart)].Value = item.NATIONAL_NO;
                    rowStart++;
                }
                //
                ew.Cells["A:AZ"].AutoFitColumns();
                string filename = "Results_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";
                Response.Clear();
                Response.ContentType = "application/vnd.ms-excel";
                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", "attachment; filename=Report.xlsx");
                Response.ContentEncoding = Encoding.UTF8;
                StringWriter stringWriter = new StringWriter();
                HtmlTextWriter hw = new HtmlTextWriter(stringWriter);
                Response.Write(stringWriter.ToString());
                Response.BinaryWrite(p1.GetAsByteArray());
                Response.End();
            }
            ViewBag.yy = yenum;
            ViewBag.yy2 = monum;
            //return View();
        }
    }
}