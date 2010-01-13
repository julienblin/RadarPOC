<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Edit.aspx.cs" Inherits="Russell.RADAR.POC.WebApp.OpinionDocuments.Edit" ValidateRequest="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>

    <script src="../javascripts/ckeditor/ckeditor.js" type="text/javascript" />

</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h2>
            Discussion</h2>
        <textarea id="textAreaDiscussion" name="textAreaDiscussion" cols="1" runat="server"></textarea>
    </div>
    <div>
        <h2>
            Investment Staff</h2>
        <textarea id="textAreaInvestmentStaff" name="textAreaInvestmentStaff" cols="1" runat="server"></textarea>
    </div>
    <div>
        <asp:Button ID="buttonSave" Text="Save" runat="server" />
        <asp:Button ID="buttonCancel" Text="Cancel" runat="server" />
    </div>
    </form>

    <script type="text/javascript">
	    CKEDITOR.replace('textAreaDiscussion');
	    CKEDITOR.replace('textAreaInvestmentStaff');
    </script>

</body>
</html>
