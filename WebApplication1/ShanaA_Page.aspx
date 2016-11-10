<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="ShanaA_Page.aspx.cs" Inherits="WebApplication1.ShanaA_Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">

    <p dir="rtl">
        ברוכים הבאים לאתר ש&quot;ב של &quot;תכנות א&quot; - תשנז (עונת 2016-2017 כאילו)</p>
    <p>
        Please enter your id (9 digits)<br />
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Get my HWs" />
        <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ErrorMessage="Must be 9 digits number" ControlToValidate="TextBox1" ValidationExpression="\d{9}"></asp:RegularExpressionValidator>
    </p>
    <p>
    </p>

</asp:Content>
