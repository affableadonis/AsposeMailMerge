﻿using Aspose.Cells;
using Aspose.Email;
using Aspose.Email.Mail;
using Aspose.Words;
using MailMerge.Libs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web.UI.WebControls;

namespace MailMerge
{
    public partial class VerifySend : System.Web.UI.Page
    {
        static string filePath;
        static List<Model.Employee> lst = null;
        protected void Page_Load(object sender, EventArgs e)
        {
            if(!IsPostBack)
            {
                lblMessage.Text = "";
                if (Session["FileName"] != null)
                {
                    filePath = Session["FileName"].ToString();
                    ReadAndPopulateData(filePath);
                }
                else
                {
                    Response.Redirect("~/Upload.aspx", false);
                }
            }
        }

        private void ReadAndPopulateData(string filename)
        {
            try
            {
                Workbook employeeData = new Workbook(filename);
                Worksheet sheet = employeeData.Worksheets[0];
                int count = sheet.Cells.Rows.Count - 1;
                DataTable dt = sheet.Cells.ExportDataTable(1, 0, count, 4);
                dt.Columns[0].ColumnName = "FullName";
                dt.Columns[1].ColumnName = "Email";
                dt.Columns[2].ColumnName = "Address";
                dt.Columns[3].ColumnName = "Salary";

                rptEmployeeData.DataSource = dt;
                rptEmployeeData.DataBind();

                lst = new List<Model.Employee>();
                lst = DataTableToObject(dt);
            }
            catch (Exception ex)
            {
                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Text = "Oooppss... There is something went wrong! Check Log.";
                ExceptionManager.LogError(ex);
            }
        }

        private List<Model.Employee> DataTableToObject(DataTable dt)
        {
            try
            {
                var convertedList = (from rw in dt.AsEnumerable()
                                     select new Model.Employee()
                                     {
                                         FullName = rw["FullName"].ToString(),
                                         Address = rw["Address"].ToString(),
                                         Email = rw["Email"].ToString(),
                                         Salary = Convert.ToInt32(rw["Salary"])
                                     }).ToList();

                return convertedList;
            }
            catch (Exception ex)
            {
                ExceptionManager.LogError(ex);
                return null;
            }
        }

        private void sendEmail(Model.Employee emp)
        {
            try
            {
                //Declare msg as MailMessage instance
                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();

                //use MailMessage properties like specify sender, recipient and message
                msg.To.Add(new System.Net.Mail.MailAddress(emp.Email));

                msg.From = new System.Net.Mail.MailAddress(ConfigurationManager.AppSettings["senderemail"]);

                msg.Attachments.Add(new System.Net.Mail.Attachment(Server.MapPath("~/Data/" + emp.FullName + ".docx")));

                var client = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587)
                {
                    Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["senderemail"], ConfigurationManager.AppSettings["senderpass"]),
                    EnableSsl = true
                };
                client.Send(msg);
            }
            catch (Exception ex)
            {
                ExceptionManager.LogError(ex);
            }
        }

        private void SendEmail(Model.Employee emp)
        {
            //Declare msg as MailMessage instance
            MailMessage msg = new MailMessage();

            //use MailMessage properties like specify sender, recipient and message
            msg.To.Add(new MailAddress(emp.Email));

            msg.From = new MailAddress(ConfigurationManager.AppSettings["senderemail"]);

            msg.Attachments.Add(new Attachment(Server.MapPath("~/Data/" + emp.FullName + ".docx")));

            //Create an instance of SmtpClient class
            SmtpClient client = new SmtpClient();

            //Specify your mailing host server
            client.Host = "smtp.gmail.com";

            //Specify your mail user name
            client.Username = ConfigurationManager.AppSettings["senderemail"];

            //Specify your mail password
            client.Password = ConfigurationManager.AppSettings["senderpass"];

            //Specify your Port #
            client.Port = 587;

            //Specify security option
            client.SecurityOptions = SecurityOptions.SSLExplicit;

            try
            {
                //Client.Send will send this message
                client.Send(msg);
            }

            catch (Exception ex)
            {
                ExceptionManager.LogError(ex);
            }
        }

        protected void btnSend_Click(object sender, EventArgs e)
        {
            try
            {
                foreach(Model.Employee emp in lst)
                {
                    Document doc = new Document(Server.MapPath("~/Data/MasterLetter.docx"));

                    string text = doc.GetText();

                    text = text.Replace("<date>", DateTime.Now.ToString("dd/MMM/yyyy"));
                    text = text.Replace("<name>", emp.FullName);
                    text = text.Replace("<salary>", emp.Salary.ToString());

                    // Initialize a Document.
                    Document newdoc = new Document();

                    // Use a document builder to add content to the document.
                    DocumentBuilder builder = new DocumentBuilder(newdoc);
                    builder.Writeln(text);

                    string dataDir = Server.MapPath("~/Data/" + emp.FullName + ".docx");
                    // Save the document to disk.
                    newdoc.Save(dataDir);

                    SendEmail(emp);
                }
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Increment Letters Have Been Sent Successfully!";
            }
            catch(Exception ex)
            {
                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Text = "Oooppss... There is something went wrong! Check Log.";
                ExceptionManager.LogError(ex);
            }
        }
    }
}