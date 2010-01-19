<%@ Page Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true"
    CodeBehind="Default.aspx.cs" Inherits="Russell.RADAR.POC.WebApp.Default" Title="RadarPOC" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContentPlaceHolder" runat="server">
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
        <div id="content">
            <asp:Repeater ID="repeaterDocuments" runat="server">
                <HeaderTemplate>
                    <table class="current2">
                        <thead>
                            <tr>
                                <td>
                                    <b>Id</b>
                                </td>
                                <td width="100%">
                                    <b>Type</b>
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
                            <asp:HyperLink ID="HyperLink1" runat="server" Text='Edit' NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/EditDocument.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                        </td>
                        <td>
                            <asp:HyperLink ID="HyperLink2" runat="server" Text="Print-HTML" NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/PrintHTML.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                        </td>
                        <td>
                            <asp:HyperLink ID="HyperLink3" runat="server" Text="Print-Word" NavigateUrl='<%# "~/PrintWord.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                        </td>
                        <td>
                            <asp:LinkButton ID="LinkButton1" runat="server" Text="Delete" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "Id")%>'
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
                            <asp:HyperLink ID="HyperLink4" runat="server" Text='Edit' NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/EditDocument.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                        </td>
                        <td>
                            <asp:HyperLink ID="HyperLink5" runat="server" Text="Print-HTML" NavigateUrl='<%# "~/" + DataBinder.Eval(Container.DataItem, "DocumentType") + "Documents/PrintHTML.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                        </td>
                        <td>
                            <asp:HyperLink ID="HyperLink6" runat="server" Text="Print-Word" NavigateUrl='<%# "~/PrintWord.aspx?id=" + DataBinder.Eval(Container.DataItem, "Id")%>' />
                        </td>
                        <td>
                            <asp:LinkButton ID="LinkButton2" runat="server" Text="Delete" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "Id")%>'
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
</asp:Content>
