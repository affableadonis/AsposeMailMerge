using MailMerge.Libs;
using System;
using System.Collections.Generic;
using System.IO;

namespace MailMerge
{
    public partial class UploadFile : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if(!IsPostBack)
            {
                lblMessage.Text = "";
            }
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> allowedExtensions = new List<string>();
                allowedExtensions.Add(".xls");
                allowedExtensions.Add(".xlsx");

                if (fuExcelUpload.HasFile)
                {
                    string fileExtension = Path.GetExtension(fuExcelUpload.FileName).ToLower();

                    if(allowedExtensions.IndexOf(fileExtension) > -1)
                    {
                        // Proceed Reading File
                        string dataDir = Server.MapPath(@"~/Data/");

                        string filePath = dataDir + fuExcelUpload.FileName;

                        fuExcelUpload.SaveAs(filePath);

                        Session.Add("FileName", filePath);

                        fuExcelUpload.Attributes.Clear();

                        Response.Redirect("~/VerifySend.aspx", false);
                    }
                    else
                    {
                        // Prompt Some Error if Extensions Doesn't Support
                        lblMessage.ForeColor = System.Drawing.Color.Red;
                        lblMessage.Text = "File extension is not supported.";
                    }
                }
            }
            catch (Exception ex)
            {
                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Text = "Something went wrong. Check Log.";
                ExceptionManager.LogError(ex);
            }
        }

        protected void hlnkDownload_Click(object sender, EventArgs e)
        {
            try
            {
                System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                response.ClearContent();
                response.Clear();
                response.ContentType = "text/plain";
                response.AddHeader("Content-Disposition",
                                   "attachment; filename=MasterEmployeeData.xlsx;");
                response.TransmitFile(Server.MapPath("~/Data/MasterEmployeeData.xlsx"));
                response.Flush();
                response.End();
            }
            catch (Exception ex)
            {
                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Text = "Something went wrong. Check Log.";
                ExceptionManager.LogError(ex);
            }
        }
    }
}