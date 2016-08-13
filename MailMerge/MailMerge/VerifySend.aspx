<%@ Page Title="" Language="C#" MasterPageFile="~/Main.Master" AutoEventWireup="true" CodeBehind="VerifySend.aspx.cs" Inherits="MailMerge.VerifySend" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="row">
        <div class="col-md-6 col-md-offset-3">
            <h1>
                Employee Salary Increments
            </h1>
            <p>
                Please verify the list of Employees before sending them increment letters.
            </p>
            <br />
            <asp:Repeater ID="rptEmployeeData" runat="server">
                <HeaderTemplate>
                    <div class="table-responsive">
                    <table class="table table-bordered">
                        <thead>
                            <tr>
                                <th>
                                    Full Name
                                </th>
                                <th>
                                    Email
                                </th>
                                <th>
                                    Address
                                </th>
                                <th>
                                    Salary
                                </th>
                            </tr>
                        </thead>
                        <tbody>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td>
                            <%# Eval("FullName") %>
                        </td>
                        <td>
                            <%# Eval("Email") %>
                        </td>
                        <td>
                            <%# Eval("Address") %>
                        </td>
                        <td>
                            <%# Eval("Salary") %>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </tbody>
                    </table>
                    </div>
                </FooterTemplate>
            </asp:Repeater>
            <br />
            <asp:Button ID="btnSend" runat="server" Text="Send Increment Letters" CssClass="btn btn-success" OnClick="btnSend_Click" />
            <br />
            <asp:Label ID ="lblMessage" Text ="" runat ="server"></asp:Label>
        </div>
    </div>
</asp:Content>
