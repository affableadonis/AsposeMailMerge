<%@ Page Title="" Language="C#" MasterPageFile="~/Main.Master" AutoEventWireup="true" CodeBehind="UploadFile.aspx.cs" Inherits="MailMerge.UploadFile" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="row">
        <div class="col-md-6 col-md-offset-3">
            <h1>
                Employee Salary Increments
            </h1>
            <p>
                Use this page to upload Excel File and then generate increment letter and send them to the employees.
                <br />
                <br />
                You can download sample Excel File from <asp:LinkButton runat="server" ID="hlnkDownload" OnClick="hlnkDownload_Click" Text="here"></asp:LinkButton>
            </p>
            <asp:FileUpload ID="fuExcelUpload" runat="server" />
            <br />
            <asp:Button ID="btnUpload" runat="server" Text="Upload" CssClass="btn btn-primary" OnClick="btnUpload_Click" />
            <br />
            <asp:Label ID ="lblMessage" Text ="" runat ="server"></asp:Label>
        </div>
    </div>
</asp:Content>
