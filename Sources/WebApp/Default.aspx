<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebApp._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link rel="Stylesheet" href="stylesheets/hubble.css" />
</head>
<body>
    <form id="form1" runat="server">
    <div id="container">
        <div id="header" style="height: 100%">
            <table id="menuNavigation" style="margin: 0px; height: 50px" cellspacing="0" cellpadding="0">
                <tbody>
                    <tr align="left">
                        <td valign="top">
                            <input id="PageTemplate_header_imgLogo" tabindex="-1" type="image" alt="" src="images/logo_russell.gif"
                                border="0" name="PageTemplate:header:imgLogo">
                        </td>
                        <td align="right">
                            <a class="suivant" id="PageTemplate_header_lnkHome">Back to Search</a> | <a class="suivant"
                                id="PageTemplate_header_lnkInterAction" tabindex="-1">Go to InterAction</a>
                            | <a class="suivant" id="PageTemplate_header_lnkMyPreferences" tabindex="-1">My Preferences</a>
                            | <a class="suivant" id="PageTemplate_header_lnkMyContacts" tabindex="-1">My Contacts</a>
                            | <a class="suivant" id="PageTemplate_header_lnkHelp" tabindex="-1">Help</a>&nbsp;&nbsp;
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        <div id="content">
            <div id="ProductNav">
                <table style="width: 100%" cellspacing="2" cellpadding="2">
                    <tbody>
                        <tr id="buttonRow">
                            <td valign="top">
                                <div id="PageTemplate_ucContentResearch_panTitle">
                                    <h1>
                                        Document list</h1>
                                </div>
                            </td>
                            <td>
                                <div align="right">
                                    <asp:ImageButton ID="linkNewOpinionDocument" runat="server" ImageUrl="~/images/btn_create_content_dd.gif" />
                                </div>
                            </td>
                        </tr>
                    </tbody>
                </table>
                <asp:Repeater ID="repeaterDocuments" runat="server">
                    <HeaderTemplate>
                        <table class="current2">
                            <thead>
                                <tr>
                                    <td>
                                        <b>Id</b>
                                    </td>
                                    <td>
                                        <b>Type</b>
                                    </td>
                                    <td width="100%">
                                        <b>State</b>
                                    </td>
                                    <td colspan="4">
                                        <b>Actions</b>
                                    </td>
                                </tr>
                            </thead>
                            <tbody>
                    </HeaderTemplate>
                    <ItemTemplate>
                        <tr class="trResult">
                            <td>
                                <%# DataBinder.Eval(Container.DataItem, "Id") %>
                            </td>
                            <td>
                                <%# DataBinder.Eval(Container.DataItem, "DocumentType")%>
                            </td>
                            <td>
                                <%# DataBinder.Eval(Container.DataItem, "State")%>
                            </td>
                            <td>
                                <asp:HyperLink runat="server" Text='Edit' NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/Edit.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                            </td>
                            <td>
                                <asp:HyperLink runat="server" Text="Print-HTML" NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/PrintHTML.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                            </td>
                            <td>
                                <asp:HyperLink ID="HyperLink1" runat="server" Text="Print-Word" NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/PrintPDF.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                            </td>
                            <td>
                                <asp:LinkButton runat="server" Text="Delete" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "Id")%>'
                                    OnClick="Delete_Click" />
                            </td>
                        </tr>
                    </ItemTemplate>
                    <AlternatingItemTemplate>
                        <tr class="trResultWhite">
                            <td>
                                <%# DataBinder.Eval(Container.DataItem, "Id") %>
                            </td>
                            <td>
                                <%# DataBinder.Eval(Container.DataItem, "DocumentType")%>
                            </td>
                            <td>
                                <%# DataBinder.Eval(Container.DataItem, "State")%>
                            </td>
                            <td>
                                <asp:HyperLink runat="server" Text='Edit' NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/Edit.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                            </td>
                            <td>
                                <asp:HyperLink runat="server" Text="Print-HTML" NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/PrintHTML.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                            </td>
                            <td>
                                <asp:HyperLink runat="server" Text="Print-Word" NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/PrintPDF.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                            </td>
                            <td>
                                <asp:LinkButton runat="server" Text="Delete" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "Id")%>'
                                    OnClick="Delete_Click" />
                            </td>
                        </tr>
                    </AlternatingItemTemplate>
                    <FooterTemplate>
                        </tbody> </table>
                    </FooterTemplate>
                </asp:Repeater>
            </div>
        </div>
    </form>
</body>
</html>
