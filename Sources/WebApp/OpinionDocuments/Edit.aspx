<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Edit.aspx.cs" Inherits="Russell.RADAR.POC.WebApp.OpinionDocuments.Edit" %>

<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager" runat="server" />
    <div>
        <h2>Discussion</h2>
        <telerik:RadEditor ID="editorDiscussion" runat="server">
        </telerik:RadEditor>
    </div>
    <div>
        <h2>Investment Staff</h2>
        <telerik:RadEditor ID="editorInvestmentStaff" runat="server">
        </telerik:RadEditor>
    </div>
    <div>
        <asp:Button ID="buttonSave" Text="Save" runat="server" />
        <asp:Button ID="buttonCancel" Text="Cancel" runat="server" />
    </div>
    </form>
</body>
</html>
