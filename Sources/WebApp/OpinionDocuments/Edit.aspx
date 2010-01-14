<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Edit.aspx.cs" Inherits="Russell.RADAR.POC.WebApp.OpinionDocuments.Edit" ValidateRequest="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>

    <script src="../javascripts/ckeditor/ckeditor.js" type="text/javascript" />

</head>
<body>
    <form id="form1" runat="server">
    <p>
        <h2>
            Discussion</h2>
        <textarea id="textAreaDiscussion" class="ckeditor" name="textAreaDiscussion" cols="80" rows="20" runat="server"></textarea>
    </p>
    <p>
        <h2>
            Investment Staff</h2>
        <textarea id="textAreaInvestmentStaff" class="ckeditor" name="textAreaInvestmentStaff" cols="80" rows="20" runat="server"></textarea>
    </p>
    <p>
        <asp:Button ID="buttonSave" Text="Save" runat="server" />
        <asp:Button ID="buttonCancel" Text="Cancel" runat="server" />
    </p>
    </form>

</body>
</html>
