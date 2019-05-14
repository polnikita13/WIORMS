using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using finalProject.Models;
using OfficeOpenXml;
using PagedList;
using PagedList.Mvc;

namespace finalWIORMS.Controllers
{
    public class Inward_RegisterController : Controller
    {
        private WIORMSDBEntities db = new WIORMSDBEntities();

        // GET: Inward_Register
        // public ActionResult Index(DateTime? fromDate, DateTime? toDate)

        /* if (option == "InwardNumber")
         {
             //Index action method will return a view with a student records based on what a user specify the value in textbox  
             return View(db.Inward_Register.Where(m => m.InwardNumber.Equals(search) || search == null).ToList());
         }
         else 
         {
             return View(db.Inward_Register.Where(m => m.Subject.StartsWith(search) || search == null).ToList());
         }
         */
        //  return View(db.Inward_Register.ToList());

        public ActionResult Index(DateTime? fromDate, DateTime? toDate)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }

            else
            {
                if (!fromDate.HasValue) fromDate = DateTime.Now.Date;
                if (!toDate.HasValue) toDate = fromDate.GetValueOrDefault(DateTime.Now.Date).Date.AddDays(1);
                if (toDate < fromDate) toDate = fromDate.GetValueOrDefault(DateTime.Now.Date).Date.AddDays(1);
                ViewBag.fromDate = fromDate;
                ViewBag.toDate = toDate;
                var Meeting = db.Inward_Register.Where(c => c.ReceivedOn >= fromDate && c.ReceivedOn < toDate).ToList();

                return View(Meeting);
            }

        }

        //To convert into Excel sheet
        public void ExportToExcel()
        {
            List<InwardViewModel> inwardList = db.Inward_Register.Select(x => new InwardViewModel
            {
                InwardNumber = x.InwardNumber,
                ReceivedOn = x.ReceivedOn,
                LetterDated = x.LetterDated,
                ReceivedFrom = x.ReceivedFrom,
                Subject = x.Subject,
                FromAddress = x.FromAddress,
                FromDistrict = x.FromDistrict,
                FromState = x.FromState,
                FromCountry = x.FromCountry,
                ReplyFor = x.ReplyFor,
                ReplyDetails = x.ReplyDetails,
                BrowseFile = x.BrowseFile,
                Status = x.Status,
                Users = x.Users
            }).ToList();

            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Report");

            ws.Cells["A1"].Value = "Report";
            ws.Cells["B1"].Value = "Inward Register";

            ws.Cells["A2"].Value = "Date";
            ws.Cells["B2"].Value = string.Format("{0:dd MMMM yyyy} at {0:H: mm tt}", DateTimeOffset.Now);

            ws.Cells["A5"].Value = "Inward Number";
            ws.Cells["B5"].Value = "Received On";
            ws.Cells["C5"].Value = "Letter Dated";
            ws.Cells["D5"].Value = "Received From";
            ws.Cells["E5"].Value = "Subject";
            ws.Cells["F5"].Value = "FromAddress";
            ws.Cells["G5"].Value = "FromDistrict";
            ws.Cells["H5"].Value = "FromState";
            ws.Cells["I5"].Value = "FromCountry";
            ws.Cells["J5"].Value = "ReplyFor";
            ws.Cells["K5"].Value = "ReplyDetails";
            ws.Cells["L5"].Value = "BrowseFile";
            ws.Cells["M5"].Value = "Status";
            ws.Cells["N5"].Value = "Users";

            int rowStart = 6;
            foreach (var item in inwardList)
            {
                ws.Cells[string.Format("A{0}", rowStart)].Value = item.InwardNumber;
                ws.Cells[string.Format("B{0}", rowStart)].Value = item.ReceivedOn.ToString();
                ws.Cells[string.Format("C{0}", rowStart)].Value = item.LetterDated.ToString();
                ws.Cells[string.Format("D{0}", rowStart)].Value = item.ReceivedFrom;
                ws.Cells[string.Format("E{0}", rowStart)].Value = item.Subject;
                ws.Cells[string.Format("F{0}", rowStart)].Value = item.FromAddress;
                ws.Cells[string.Format("G{0}", rowStart)].Value = item.FromDistrict;
                ws.Cells[string.Format("H{0}", rowStart)].Value = item.FromState;
                ws.Cells[string.Format("I{0}", rowStart)].Value = item.FromCountry;
                ws.Cells[string.Format("J{0}", rowStart)].Value = item.ReplyFor;
                ws.Cells[string.Format("K{0}", rowStart)].Value = item.ReplyDetails;
                ws.Cells[string.Format("L{0}", rowStart)].Value = item.BrowseFile;
                ws.Cells[string.Format("M{0}", rowStart)].Value = item.Status;
                ws.Cells[string.Format("N{0}", rowStart)].Value = item.Users;
                rowStart++;
            }

            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment: filename" + "ExcelReport.xlsx");
            Response.BinaryWrite(pck.GetAsByteArray());
            Response.End();

        }




        // GET: Inward_Register/Details/5        
        public ActionResult Details(int? id)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                if (Session["data"] != null)
                {
                    db.Database.ExecuteSqlCommand("Update Inward_Register set Users='" + Session["data"].ToString() + "' where Inwardnumber=" + id);
                    ViewBag.udata = Session["data"].ToString();
                }
                if (id == null)
                {
                    return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
                }
                Inward_Register inward_Register = db.Inward_Register.Find(id);
                if (inward_Register == null)
                {
                    return HttpNotFound();
                }
                return View(inward_Register);
            }
        }
        [HttpPost]
        public ActionResult Details(decimal InwardNumber, String Status, String Users)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                Inward_Register obj = db.Inward_Register.ToList().Single(m => m.InwardNumber == InwardNumber);
                obj.Status = Status;
                obj.Users = Users;
                db.Entry(obj).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("GetList");
            }
        }

        // GET: Inward_Register/Create
        public ActionResult Create()
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                Inward_Register inward_Register = new finalProject.Models.Inward_Register();
                
                return View(inward_Register);
            }
        }

         
        // POST: Inward_Register/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "InwardNumber,ReceivedOn,LetterDated,ReceivedFrom,Subject,FromCountry,FromState,FromDistrict,FromAddress,ReplyFor,ReplyDetails,BrowseFile")] Inward_Register inward_Register, HttpPostedFileBase file1)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                if (ModelState.IsValid)
                {
                    try
                    {
                        var allowedExtensions = new[] { ".pdf", ".docx", ".doc" };
                        var checkextension = Path.GetExtension(file1.FileName).ToLower();

                        if (!allowedExtensions.Contains(checkextension))
                        {
                            TempData["msg"] = "Select pdf or doc less than 1Mb";
                        }


                        if (file1 != null && file1.ContentLength > 0)
                        {
                            foreach (var itm in allowedExtensions)
                            {
                                if (itm.Contains(checkextension))
                                {
                                    var extension = Path.GetExtension(file1.FileName);
                                    //  var path = Path.Combine(Server.MapPath("~/Content/AnnFiles/" + "announcement_" + announcement.anak_ID + extension));

                                    file1.SaveAs(Server.MapPath("~/Uploads/" + file1.FileName));
                                    inward_Register.BrowseFile = "~/Uploads/" + file1.FileName;
                                    db.Inward_Register.Add(inward_Register);
                                    db.SaveChanges();


                                    return RedirectToAction("InwardEmail");
                                  //  return RedirectToAction("Index");
                                }
                            }

                        }

                    }
                    catch (System.Data.Entity.Validation.DbEntityValidationException dbEx)
                    {
                        Exception raise = dbEx;
                        foreach (var validationErrors in dbEx.EntityValidationErrors)
                        {
                            foreach (var validationError in validationErrors.ValidationErrors)
                            {
                                string message = string.Format("{0}:{1}",
                                    validationErrors.Entry.Entity.ToString(),
                                    validationError.ErrorMessage);
                                // raise a new exception nesting
                                // the current instance as InnerException
                                raise = new InvalidOperationException(message, raise);
                            }
                        }
                        throw raise;
                    }
                }

                return View(inward_Register);
            }
        }

        // GET: Inward_Register/Edit/5
        public ActionResult Edit(int? id)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                if (id == null)
                {
                    return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
                }
                Inward_Register inward_Register = db.Inward_Register.Find(id);
                if (inward_Register == null)
                {
                    return HttpNotFound();
                }
                return View(inward_Register);
            }
        }

        // POST: Inward_Register/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "InwardNumber,ReceivedOn,LetterDated,ReceivedFrom,Subject,FromAddress,FromDistrict,FromState,FromCountry,ReplyFor,ReplyDetails,BrowseFile")] Inward_Register inward_Register)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                if (ModelState.IsValid)
                {
                    db.Entry(inward_Register).State = EntityState.Modified;
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }
                return View(inward_Register);

            }
        }

        // GET: Inward_Register/Delete/5
        public ActionResult Delete(int? id)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                if (id == null)
                {
                    return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
                }
                Inward_Register inward_Register = db.Inward_Register.Find(id);
                if (inward_Register == null)
                {
                    return HttpNotFound();
                }
                return View(inward_Register);
            }
        }

        // POST: Inward_Register/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int? id)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                Inward_Register inward_Register = db.Inward_Register.Find(id);
                db.Inward_Register.Remove(inward_Register);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
        }

        // GET: Inward_Register

        //this list is for Director's Dashboard to display all inward letter list
        public ActionResult GetList()
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View(db.Inward_Register.OrderByDescending(m => m.InwardNumber).ToList());
            }
        }
        // GET: Inward_Register
        //this list is for receptionalist to display the accepted inward letter
        public ActionResult GetList2()
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View(db.Inward_Register.OrderByDescending(m => m.InwardNumber).ToList().Where(m => m.Status == "Accepted"));

            }
        }
        // GET: Inward_Register
        //this list is for receptionalist to display the rejected inward letter
        public ActionResult GetList3()
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View(db.Inward_Register.OrderByDescending(m => m.InwardNumber).ToList().Where(m => m.Status == "Rejected"));
            }
        }
        /*
        public ActionResult GetList4()
        {

            return View(db.Inward_Register.ToList().Where(m => m.Users == Session["user"]));

        }*/

        public ActionResult BrowseUsers(int Id)
        {
            Session["iid"] = Id;
            return View(db.UserRegistrations.ToList());
        }
        [HttpPost]
        public ActionResult BrowseUsers(FormCollection fc)
        {
            Session["data"] = fc["data"];
            return RedirectToAction("Details", new { Id = Session["iid"].ToString() });
        }

        //To send mail to intended recepients
        public ActionResult SendEmail(int id)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                Inward_Register obj = db.Inward_Register.Single(m => m.InwardNumber == id);
                String users = obj.Users;
                if (users != null && users.Length > 0)
                {
                    users = users.Substring(0, users.Length - 1);
                    String[] s = users.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    MailMessage msg = new MailMessage();
                    msg.Sender = new MailAddress("ritoffice2019@gmail.com");
                    foreach (String em in s)
                    {
                        msg.To.Add(em);
                    }
                    msg.Subject = obj.Subject;
                    msg.Body = "Mail From Office";
                    msg.Attachments.Add(new Attachment(Server.MapPath(obj.BrowseFile)));

                    SmtpClient sc = new SmtpClient();
                    sc.Port = 587;
                    sc.Host = "smtp.gmail.com";
                    sc.EnableSsl = true;
                    sc.DeliveryMethod = SmtpDeliveryMethod.Network;
                    sc.Credentials = new NetworkCredential("ritoffice2019@gmail.com", "Admin@rit");
                    msg.From = new MailAddress("ritoffice2019@gmail.com");

                    sc.Send(msg);
                    obj.Status1 = "true";
                    db.SaveChanges();

                    TempData["msg"] = "<script>alert('Email sent succesfully');</script>";
                }

                return RedirectToAction("GetList2");
            }
        }

        public ActionResult InwardEmail()
        {
            using (MailMessage mm = new MailMessage("ritoffice2019@gmail.com", "poojabandgar1313@gmail.com"))
            {
                mm.Subject = "Inward Letter";
                mm.Body = "New Inward Letter Arrived..";
              
                mm.IsBodyHtml = false;
                using (SmtpClient smtp = new SmtpClient())
                {
                    smtp.Host = "smtp.gmail.com";
                    smtp.EnableSsl = true;                    
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = new NetworkCredential("ritoffice2019@gmail.com", "Admin@rit");
                    smtp.Port = 587;
                    smtp.Send(mm);
                }
            }
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

    }
}
