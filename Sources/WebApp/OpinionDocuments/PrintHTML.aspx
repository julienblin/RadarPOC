<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PrintHTML.aspx.cs" Inherits="Russell.RADAR.POC.WebApp.OpinionDocuments.PrintHTML" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h2>Discussion</h2>
        <asp:Literal ID="literalDiscussion" runat="server" />
    </div>
    <div>
        <h2>Investment Staff</h2>
        <asp:Literal ID="literalInvestmentStaff" runat="server" />
    </div>
    </form>
</body>
</html>
