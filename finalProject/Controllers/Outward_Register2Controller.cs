using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using finalProject.Models;
using System.Net.Mail;
using System.IO;
using OfficeOpenXml;

namespace finalWIORMS.Controllers
{
    public class Outward_Register2Controller : Controller
    {
        private WIORMSDBEntities db = new WIORMSDBEntities();

        // GET: Outward_Register
        /* public ActionResult Index()
         {
             return View(db.Outward_Register2.ToList());
         }*/
        public ActionResult Index(DateTime? fromDate, DateTime? toDate)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                //this code is for to display the letter within given date range
                if (!fromDate.HasValue) fromDate = DateTime.Now.Date;
                if (!toDate.HasValue) toDate = fromDate.GetValueOrDefault(DateTime.Now.Date).Date.AddDays(1);
                if (toDate < fromDate) toDate = fromDate.GetValueOrDefault(DateTime.Now.Date).Date.AddDays(1);
                ViewBag.fromDate = fromDate;
                ViewBag.toDate = toDate;
                var Meeting = db.Outward_Register2.Where(c => c.LetterDated >= fromDate && c.LetterDated < toDate).ToList();

                return View(Meeting);
            }
        }

        //To convert into excel sheet
        public void ExportToExcel()
        {
            List<OutwardViewModel> outwardList = db.Outward_Register2.Select(x => new OutwardViewModel
            {
                OutwardNumber = x.OutwardNumber,
                LetterDated = x.LetterDated,
                SendTo = x.SendTo,
                Subject = x.Subject,
                ToAddress = x.ToAddress,
                ToDistrict = x.ToDistrict,
                ToState = x.ToState,
                ToCountry = x.ToCountry,
                ReplyFor = x.ReplyFor,
                ReplyDetails = x.ReplyDetails,
                BrowseFile = x.BrowseFile,
                SendOn = x.SendOn,
                DispatchType = x.DispatchType,
                ReceiptFile = x.ReceiptFile,
                DispatchDetails = x.DispatchDetails,
                DispatchCharges = x.DispatchCharges,
                Status = x.Status

            }).ToList();

            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Report");

            ws.Cells["A1"].Value = "Report";
            ws.Cells["B1"].Value = "Outward Register";

            ws.Cells["A2"].Value = "Date";
            ws.Cells["B2"].Value = string.Format("{0:dd MMMM yyyy} at {0:H: mm tt}", DateTimeOffset.Now);

            ws.Cells["A5"].Value = "Outward Number";
            ws.Cells["B5"].Value = "Letter Dated";
            ws.Cells["C5"].Value = "Sent To";
            ws.Cells["D5"].Value = "Subject";
            ws.Cells["E5"].Value = "ToAddress";
            ws.Cells["F5"].Value = "ToDistrict";
            ws.Cells["G5"].Value = "ToState";
            ws.Cells["H5"].Value = "ToCountry";
            ws.Cells["I5"].Value = "ReplyFor";
            ws.Cells["J5"].Value = "ReplyDetails";
            ws.Cells["K5"].Value = "BrowseFile";
            ws.Cells["L5"].Value = "SendOn";
            ws.Cells["M5"].Value = "Dispatch Type";
            ws.Cells["N5"].Value = "Receipt File";
            ws.Cells["O5"].Value = "Displatch Details";
            ws.Cells["P5"].Value = "Displatch Charges";
            ws.Cells["Q5"].Value = "Status";

            int rowStart = 6;
            foreach (var item in outwardList)
            {
                ws.Cells[string.Format("A{0}", rowStart)].Value = item.OutwardNumber;
                ws.Cells[string.Format("B{0}", rowStart)].Value = item.LetterDated.ToString();
                ws.Cells[string.Format("C{0}", rowStart)].Value = item.SendTo;
                ws.Cells[string.Format("D{0}", rowStart)].Value = item.Subject;
                ws.Cells[string.Format("E{0}", rowStart)].Value = item.ToAddress;
                ws.Cells[string.Format("F{0}", rowStart)].Value = item.ToDistrict;
                ws.Cells[string.Format("G{0}", rowStart)].Value = item.ToState;
                ws.Cells[string.Format("H{0}", rowStart)].Value = item.ToCountry;
                ws.Cells[string.Format("I{0}", rowStart)].Value = item.ReplyFor;
                ws.Cells[string.Format("J{0}", rowStart)].Value = item.ReplyDetails;
                ws.Cells[string.Format("K{0}", rowStart)].Value = item.BrowseFile;
                ws.Cells[string.Format("L{0}", rowStart)].Value = item.SendOn.ToString();
                ws.Cells[string.Format("M{0}", rowStart)].Value = item.DispatchType;
                ws.Cells[string.Format("N{0}", rowStart)].Value = item.ReceiptFile;
                ws.Cells[string.Format("O{0}", rowStart)].Value = item.DispatchDetails;
                ws.Cells[string.Format("P{0}", rowStart)].Value = item.DispatchCharges;
                ws.Cells[string.Format("Q{0}", rowStart)].Value = item.Status;
                rowStart++;
            }

            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment: filename" + "ExcelReport.xlsx");
            Response.BinaryWrite(pck.GetAsByteArray());
            Response.End();

        }

        // GET: Outward_Register/Details/5
        public ActionResult Details(int? id)
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
                Outward_Register2 outward_Register = db.Outward_Register2.Find(id);
                if (outward_Register == null)
                {
                    return HttpNotFound();
                }
                return View(outward_Register);
            }
        }
        [HttpPost]
        public ActionResult Details(int OutwardNumber, String Status)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                Outward_Register2 obj = db.Outward_Register2.ToList().Single(m => m.OutwardNumber == OutwardNumber);
                obj.Status = Status;
                db.Entry(obj).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("GetList");
            }
        }

        // GET: Outward_Register/Create
        public ActionResult Create()
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                Outward_Register2 outward_Register = new finalProject.Models.Outward_Register2();
                //  outward_Register.OutwardNumber = (decimal)db.PKS.ToList()[0].Id2;            
                //outward_Register.LetterDated = DateTime.Now;
                //outward_Register.SentOn = DateTime.Now;
                return View(outward_Register);
            }
        }

        // POST: Outward_Register/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "OutwardNumber,LetterDated,SendTo,Subject,ToAddress,ToDistrict,ToState,ToCountry,ReplyFor,ReplyDetails,BrowseFile,SendOn,DispatchType,ReceiptFile,DispatchDetails,DispatchCharges")] Outward_Register2 outward_Register, HttpPostedFileBase file1, HttpPostedFileBase file2)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                if (ModelState.IsValid)
                {
                    file1.SaveAs(Server.MapPath("~/Uploads/" + file1.FileName));
                    outward_Register.BrowseFile = "~/Uploads/" + file1.FileName;
                    if (outward_Register.DispatchType == "Post" || outward_Register.DispatchType == "Courier")
                    {
                        file2.SaveAs(Server.MapPath("~/Uploads/" + file2.FileName));
                        outward_Register.ReceiptFile = "~/Uploads/" + file2.FileName;

                    }
                    db.Outward_Register2.Add(outward_Register);

                    db.SaveChanges();

                    //  PK pk = db.PKS.ToList()[0];
                    // pk.Id1 = ((int)pk.Id2) + 1;
                    //db.Entry(pk).State = EntityState.Modified;
                    //db.SaveChanges();
                    return RedirectToAction("OutwardEmail");
                }

                return View(outward_Register);
            }
        }

        // GET: Outward_Register/Edit/5
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
                Outward_Register2 outward_Register = db.Outward_Register2.Find(id);
                if (outward_Register == null)
                {
                    return HttpNotFound();
                }
                return View(outward_Register);

            }
        }

        // POST: Outward_Register/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "OutwardNumber,LetterDated,SendTo,Subject,ToAddress,ToDistrict,ToState,ToCountry,ReplyFor,ReplyDetails,BrowseFile,SendOn,DispatchType,DispatchDetails,DispatchCharges")] Outward_Register2 outward_Register)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                if (ModelState.IsValid)
                {
                    db.Entry(outward_Register).State = EntityState.Modified;
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }
                return View(outward_Register);
            }
        }

        // GET: Outward_Register/Delete/5
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
                Outward_Register2 outward_Register = db.Outward_Register2.Find(id);
                if (outward_Register == null)
                {
                    return HttpNotFound();
                }
                return View(outward_Register);
            }
        }

        // POST: Outward_Register/Delete/5
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
                Outward_Register2 outward_Register = db.Outward_Register2.Find(id);
                db.Outward_Register2.Remove(outward_Register);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
        }
        // GET: Outward_Register
        //this list is for Director's Dashboard to display all outward letter list
        public ActionResult GetList()
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View(db.Outward_Register2.OrderByDescending(m => m.OutwardNumber).ToList());
            }
        }

        // GET: Outward_Register
        //this list is for receptionalist to display the accepted outward letter
        public ActionResult GetList2()
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View(db.Outward_Register2.OrderByDescending(m => m.OutwardNumber).ToList().Where(m => m.Status == "Accepted"));

            }
        }
        // GET: Outward_Register
        //this list is for receptionalist to display the rejected outward letter
        public ActionResult GetList3()
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View(db.Outward_Register2.OrderByDescending(m => m.OutwardNumber).ToList().Where(m => m.Status == "Rejected"));

            }
        }
        public ActionResult OpenEmail(int id)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                Outward_Register2 obj = db.Outward_Register2.Single(m => m.OutwardNumber == id);
                return View(obj);
            }
        }
        [HttpPost]

        public ActionResult SendEmail(int OutwardNumber, String Subject, String bodytext, String sender, String BrowseFile)
        {
            Outward_Register2 obj = db.Outward_Register2.Single(m => m.OutwardNumber == OutwardNumber);
            String users = sender;
            if (users != null && users.Length > 0)
            {
                users = users.Substring(0, users.Length - 1);
                String[] s = users.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                System.Net.Mail.MailMessage msg = new MailMessage();
                msg.Sender = new MailAddress("polnikita13@gmail.com");
                foreach (String em in s)
                {
                    msg.To.Add(em);
                }
                msg.Subject = obj.Subject;
                msg.Body = bodytext;
                msg.Attachments.Add(new Attachment(Server.MapPath(BrowseFile)));

                SmtpClient sc = new SmtpClient();
                sc.Port = 587;
                sc.Host = "smtp.gmail.com";
                sc.EnableSsl = true;
                sc.DeliveryMethod = SmtpDeliveryMethod.Network;
                sc.Credentials = new NetworkCredential("poojabandgar1313@gmail.com", "poojasagarsai");
                msg.From = new MailAddress("poojabandgar1313@gmail.com");

                sc.Send(msg);
            }

            return RedirectToAction("GetList2");
        }


        // GET: Home
        public ActionResult Index1()
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                return View();
            }
        }

        [HttpPost]
        //to send mail
        public ActionResult Index1(EmailMode model)
        {
            if (Session["user"] == null)
            {
                return RedirectToAction("Index", "Home");
            }
            else
            {
                using (MailMessage mm = new MailMessage("ritoffice2019@gmail.com", model.To))
                {
                    mm.Subject = model.Subject;
                    mm.Body = model.Body;
                    if (model.Attachment.ContentLength > 0)
                    {
                        string fileName = Path.GetFileName(model.Attachment.FileName);
                        mm.Attachments.Add(new Attachment(model.Attachment.InputStream, fileName));
                    }
                    mm.IsBodyHtml = false;
                    using (SmtpClient smtp = new SmtpClient())
                    {
                        smtp.Host = "smtp.gmail.com";
                        smtp.EnableSsl = true;
                        // NetworkCredential NetworkCred = new NetworkCredential(model.Email, model.Password);
                        smtp.UseDefaultCredentials = true;
                        // smtp.Credentials = NetworkCred;
                        smtp.Credentials = new NetworkCredential("ritoffice2019@gmail.com", "Admin@rit");
                        smtp.Port = 587;
                        smtp.Send(mm);
                        //ViewBag.Message = "Email sent succesfully.";
                        TempData["msg"] = "<script>alert('Email sent succesfully');</script>";
                    }
                }

                return View();
            }
        }

        public ActionResult OutwardEmail()
        {
            using (MailMessage mm = new MailMessage("ritoffice2019@gmail.com", "poojabandgar1313@gmail.com"))
            {
                mm.Subject = "Outward Letter";
                mm.Body = "New Outward Letter Arrived..";

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
