using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using GFS.Models;
using System.Web.Helpers;
using System.Text;
using System.Net.Mail;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using System.Drawing;
using iTextSharp.text.html.simpleparser;
using Font = iTextSharp.text.Font;
using System.IO;

namespace GFS.Controllers
{
    public class PaymentsController : Controller
    {
        private GFSContext db = new GFSContext();
        // GET: Payments
        //public ActionResult Index()
        //{
        //    return View(db.Payments.ToList());
        //}
        public ActionResult Index(string searchString)
        {
            var pay = from m in db.Payments
                        select m;

            if (!String.IsNullOrEmpty(searchString))
            {
                pay = pay.Where(s => s.policyNo.Contains(searchString));
            }
            return View(pay);
        }

        public FileStreamResult PrintList()
        {
            //Set up the document and the MS to write it to and create the PDF writer instance
            MemoryStream memStream = new MemoryStream();
            // Create a Document object
            var document = new Document(PageSize.A4, 0, 0, 0, 0);

            // Create a new PdfWriter object, specifying the output stream
            var output = new MemoryStream();
            var writer = PdfWriter.GetInstance(document, output);

            // Open the Document for writing
            document.Open();

            //Set up fonts used in the document
            Font fontHeading3 = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, Font.BOLD + Font.UNDERLINE, BaseColor.BLACK);
            Font subHeaderFont = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 10, Font.BOLDITALIC, BaseColor.BLACK);
            Font PolicyDetailFont = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, Font.BOLD, BaseColor.BLACK);
            Font fontBody = FontFactory.GetFont(FontFactory.TIMES_ITALIC, 10, Font.BOLD, BaseColor.BLACK);
            Font thFont = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, Font.BOLD, BaseColor.BLACK);
            Font redFont = FontFactory.GetFont(FontFactory.COURIER_BOLDOBLIQUE, 18, Font.BOLD, BaseColor.RED);
            Font fontData = FontFactory.GetFont(FontFactory.COURIER_OBLIQUE, 12, BaseColor.BLACK);


            //Open the PDF document
            document.Open();

            PdfPTable tblHeader = new PdfPTable(3)
            {
                SpacingBefore = 50f,
                SpacingAfter = 50f
            };

            PdfPCell hCell6 = new PdfPCell(new Phrase("\nGumbi Financial Services\n", fontHeading3))
            {
                HorizontalAlignment = 1,
                Colspan = 3,
                Border = 0
            };

            PdfPCell hCell7 = new PdfPCell(new Phrase("Gumbi Financial Services, Durban, South Africa", PolicyDetailFont))
            {
                Border = 0,
                Colspan = 3,
                HorizontalAlignment = 1
            };

            PdfPCell hCell8 = new PdfPCell(new Phrase("Tel: 031-459-7500 / Fax: 031-459-7600", PolicyDetailFont))
            {
                Border = 0,
                Colspan = 3,
                HorizontalAlignment = 1
            };
            PdfPCell hCell9 = new PdfPCell(new Phrase("Email :Info@gumbi.com / Enquiry@gumbi.com", PolicyDetailFont))
            {
                Colspan = 3,
                Border = 0,
                HorizontalAlignment = 1
            };



            tblHeader.AddCell(hCell6);
            tblHeader.AddCell(hCell7);
            tblHeader.AddCell(hCell8);
            tblHeader.AddCell(hCell9);

            document.Add(tblHeader);

            //Get individual Payment Details for a Member
            var obj = db.Payments.ToList().FindLast(x => x.policyNo == Session["det"].ToString());

            var orderInfoTable = new PdfPTable(1);
            orderInfoTable.HorizontalAlignment = Element.ALIGN_CENTER;
            orderInfoTable.SpacingBefore = 40;
            orderInfoTable.SpacingAfter = 50;
            orderInfoTable.DefaultCell.Border = 0;
            orderInfoTable.WidthPercentage = 80;



            PdfPCell tc1 = new PdfPCell(new Phrase("Policy No: " + obj.policyNo, fontData)) { Indent = 5, Border = 0, Colspan = 2, };
            orderInfoTable.AddCell(tc1);
            PdfPCell tc2 = new PdfPCell(new Phrase("Customer Name:" + obj.CustomerName, fontData)) { Indent = 5, Border = 0 };
            orderInfoTable.AddCell(tc2);
            PdfPCell tc3 = new PdfPCell(new Phrase("Policy Plan:" + obj.plan, fontData)) { Indent = 5, Border = 0, Colspan = 2, };
            orderInfoTable.AddCell(tc3);
            PdfPCell tc4 = new PdfPCell(new Phrase("Amount Due:" + Convert.ToString(obj.dueAmount), fontData)) { Indent = 5, Border = 0, Colspan = 2, };
            orderInfoTable.AddCell(tc4);
            PdfPCell tc5 = new PdfPCell(new Phrase("Amount Paid:" + Convert.ToString(obj.amount), fontData)) { Indent = 5, Border = 0, Colspan = 2, };
            orderInfoTable.AddCell(tc5);
            PdfPCell tc6 = new PdfPCell(new Phrase("Out Standing Amount:" + Convert.ToString(obj.outstandingAmount), fontData)) { Indent = 5, Border = 0, Colspan = 2, };
            orderInfoTable.AddCell(tc6);
            PdfPCell tc7 = new PdfPCell(new Phrase("Date :" + DateTime.Now.ToString(), fontData)) { Indent = 5, Border = 0, Colspan = 2, };
            orderInfoTable.AddCell(tc7);
            PdfPCell tc8 = new PdfPCell(new Phrase("Captured By:" + obj.cashierName, fontData)) { Indent = 5, Border = 0, Colspan = 2, };
            orderInfoTable.AddCell(tc8);
            PdfPCell tc9 = new PdfPCell(new Phrase("Branch: " + obj.branch, fontData)) { Indent = 5, Border = 0, Colspan = 2, };
            orderInfoTable.AddCell(tc9);



            var paragraph = new Paragraph("Payment Details Receipt", redFont);
            paragraph.Alignment = Element.ALIGN_CENTER;
            document.Add(paragraph);
            document.Add(orderInfoTable);

            Paragraph gap = new Paragraph("\n\n");

            document.Add(gap);

            PdfPTable tblOfficeDetail = new PdfPTable(2);
            PdfPCell signCell = new PdfPCell(new Phrase("Signature: .....................................", fontBody)) { Colspan = 2, HorizontalAlignment = 3, Border = 0 };
            tblOfficeDetail.AddCell(signCell);
            PdfPCell dateCell = new PdfPCell(new Phrase("Date Issued:............/............/..........", fontBody)) { HorizontalAlignment = 3, Border = 0 };
            tblOfficeDetail.AddCell(dateCell);
            PdfPCell stampCell = new PdfPCell(new Phrase("Stamp  :.......................................", fontBody)) { Rowspan = 2, HorizontalAlignment = 3, Border = 0 };
            tblOfficeDetail.AddCell(stampCell);

            document.Add(tblOfficeDetail);
            var logo = iTextSharp.text.Image.GetInstance(Server.MapPath("~/Images/Logo.png"));
            logo.Alignment = Element.ALIGN_CENTER; // Absolute position
            document.Add(logo);
            document.Close();

            Response.ContentType = "application/pdf";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename=Receipt-{0}.pdf", obj.policyNo));
            Response.BinaryWrite(output.ToArray());
            return new FileStreamResult(memStream, "application/pdf");
        }


        // GET: Payments/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Payment payment = db.Payments.Find(id);
            if (payment == null)
            {
                return HttpNotFound();
            }
            return View(payment);
        }

        public ActionResult Search()
        {
            Session["polNo"] = null;
            Session["fullname"] = null;
            Session["plan"] = null;
            return View();
        }

        [HttpPost]
        public ActionResult Search(string searchStr)
        {
            var payment = from m in db.NewMembers
                          select m;

            if (!String.IsNullOrEmpty(searchStr))
            {
                payment = payment.Where(s => s.policyNo.Contains(searchStr));

                var d = db.NewMembers.ToList().Find(r => r.policyNo == searchStr);
                var du = db.Payers.ToList().Find(r => r.policyNo == searchStr);
                var stand = db.Payments.ToList().FindLast(r => r.policyNo == searchStr);
                if(d!=null)
                {
                    Session["polNo"] = d.policyNo;
                    Session["fullname"] = d.fName + " " + d.lName;
                    Session["plan"] = d.Policyplan;
                }
                else if(d==null)
                {
                    Session["responce3"] = "Sorry, Member you searched for does not exist in the database! please add the Member first.";
                    return View("Search");
                }
                if(stand!=null)
                {
                    Session["iniPrem"] = du.initialPremium + stand.outstandingAmount;
                }
                else if(stand==null)
                {
                    Session["iniPrem"] = du.initialPremium;
                }
                
                return RedirectToAction("Create");
            }
            return View(payment);
        }

        // GET: Payments/Create
        public ActionResult Create()
        {
            var planList = new List<SelectListItem>();
            var PlanQuery = from e in db.PolicyPlans select e;
            foreach (var m in PlanQuery)
            {
                planList.Add(new SelectListItem { Value = m.policyType, Text = m.policyType });
            }
            ViewBag.plnlist = planList;
            return View();
        }

        // POST: Payments/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "referenceNo,policyNo,CustomerName,plan,dueAmount,amount,outstandingAmount,datePayed,cashierName,branch,emailSlip")] Payment payment)
        {
            //if (ModelState.IsValid)
            //{
            
                if (Session["polNo"]!=null)
                {
                    payment.policyNo = Session["polNo"].ToString();
                }
                if (Session["fullname"]!=null)
                {
                    payment.CustomerName = Session["fullname"].ToString();
                }
                if (Session["plan"]!=null)
                {
                    payment.plan = Session["plan"].ToString();
                }
                if (Session["iniPrem"] != null)
                {
                    payment.dueAmount = Convert.ToDouble(Session["iniPrem"]);
                }
                else if (Session["iniPrem"] == null)
                {
                    payment.dueAmount = 0;
                }
                
                double outst = payment.dueAmount - payment.amount;
                payment.outstandingAmount = outst;
                payment.datePayed = DateTime.Now;

                db.Payments.Add(payment);
                db.SaveChanges();
                if(payment.emailSlip==true)
                {
                    try
                    {
                        var emailA = db.NewMembers.ToList().Find(p => p.policyNo == payment.policyNo);
                        var boddy = new StringBuilder();

                        boddy.Append("Dear " + payment.CustomerName + "<br/>" +
                                     "Thank You For Being G.F.S Customer" + "<br/>" +
                                     "You Just made a payment with the following details" +
                                      "Policy Number: " + payment.policyNo +
                                      "Policy Plan: " + payment.plan +
                                      "Amount That Was Due: R" + payment.dueAmount +
                                      "Amount Paid: R" + payment.amount +
                                      "Outstanding Amount: R" + payment.outstandingAmount +
                                      "Date Paid: " + payment.datePayed +
                                      "Your Cashier Was: " + payment.cashierName +
                                      "Branch: " + payment.branch + "<br/>" +
                                     "your satisfaction with our service is our priority" + "<br/>" +
                                     "==========================================" + "<br/>");

                        string body_for = boddy.ToString();
                        string to_for = "";
                        if(emailA!=null)
                        {
                            to_for = emailA.CustEmail;
                        }                       
                        string subject_for = "G.F.S Payment for "+DateTime.Now.Month.ToString();

                        WebMail.SmtpServer = "pod51014.outlook.com";
                        WebMail.SmtpPort = 587;

                        WebMail.UserName = "21353863@dut4life.ac.za";
                        WebMail.Password = "Dut930717";

                        WebMail.From = "21353863@dut4life.ac.za";
                        WebMail.EnableSsl = true;
                        WebMail.Send(to: to_for, subject: subject_for, body: body_for);
                    }
                    catch (Exception)
                    {
                        
                    }
                } 
            //}
                return RedirectToAction("Details", new { id = payment.referenceNo });
        }

        // GET: Payments/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Payment payment = db.Payments.Find(id);
            if (payment == null)
            {
                return HttpNotFound();
            }
            return View(payment);
        }

        // POST: Payments/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "referenceNo,policyNo,CustomerName,plan,dueAmount,amount,outstandingAmount,datePayed,cashierName,branch,emailSlip")] Payment payment)
        {
            if (ModelState.IsValid)
            {
                db.Entry(payment).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(payment);
        }

        // GET: Payments/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Payment payment = db.Payments.Find(id);
            if (payment == null)
            {
                return HttpNotFound();
            }
            return View(payment);
        }

        // POST: Payments/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Payment payment = db.Payments.Find(id);
            db.Payments.Remove(payment);
            db.SaveChanges();
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
