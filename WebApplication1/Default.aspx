<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebApplication1._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <p dir="rtl">
        ברוכים הבאים לאתר ש&quot;ב של&nbsp; &quot;מטלות קיץ&quot; (לקראת תכנות מבוסס ארועים תשנז)</p>
    <p>
        Please enter your id (9 digits)<br />
        <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Get my HWs" />
        <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ErrorMessage="Must be 9 digits number" ControlToValidate="TextBox1" ValidationExpression="\d{9}"></asp:RegularExpressionValidator>
    </p>
    <p>
    </p>

</asp:Content>
