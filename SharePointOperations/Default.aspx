<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="SharePointOperations.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style1 {
            width: 163px;
        }

        TD {
            font-family: Arial;
            font-size: 10pt;
        }

        .WaterMarkedTextBox {
            color: gray;
            font-size: 9pt;
            text-align: left;
        }

        .button {
            background-color: #666699; /* Green */
            color: white;
            text-align: center;
            text-decoration: none;
            font-size: 16px;
            cursor: pointer;
            transition-duration: 0.4s;
            border: none;
        }

            .button:hover {
                box-shadow: 0 4px 6px 0 rgba(0,0,0,0.17),0 6px 8px 0 rgba(0,0,0,0.18);
                background-color: #9999FF;
            }
    </style>
   


</head>
<body>
    <form id="form1" runat="server">
        <div>
            <div align="center">
                <asp:Label ID="lblRes" runat="server" Text=""></asp:Label>
            </div>

            <table align="center" valign="top">

                <tr>
                    <td class="auto-style1">Enter Sp url</td>
                    <td>
                        <asp:TextBox ID="txtUrl" runat="server" Width="350px" CssClass="WaterMarkedTextBox"></asp:TextBox></td>
                    <br />

                </tr>
                <tr>
                    <td colspan="2" align="center">
                        <asp:Button ID="btnDisplayGroups" runat="server" Text="Display list items" OnClick="btnDisplayGroups_Click" class="button" /></td>

                </tr>
                <tr runat="server" id="rwchkMessage" visible="false">
                    <td align="left" class="auto-style1" valign="top">
                        <asp:Label ID="lblchkmessage" runat="server" Text="Select list items"></asp:Label>

                    </td>
                    <td colspan="2">
                        
                        <asp:DropDownList ID="chksharePointGroups" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr runat="server" id="rwselformatte" visible="false">
                    <td class="auto-style1">Select Formatte</td>
                    <td align="left">
                        <asp:RadioButtonList ID="DropDownList1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged">
                            <asp:ListItem>Excel</asp:ListItem>
                            <asp:ListItem>SqlDataBase</asp:ListItem>

                        </asp:RadioButtonList>
                        </td>
                        
                </tr>
                 <tr runat="server" id="rwAuthentication" visible="false">
                    <td class="auto-style1">Select authentication</td>
                    <td align="left">
                        <asp:DropDownList ID="ddlSelectAuth" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlSelectAuth_SelectedIndexChanged">
                            <asp:ListItem>Windows</asp:ListItem>
                            <asp:ListItem>Sql Server</asp:ListItem>

                        </asp:DropDownList>
                        </td>
                        
                </tr>
                <tr runat="server" id="rwDbdatasource" visible="false">
                    <td>Please enter DataSource</td>
                    <td>
                        <asp:TextBox ID="txtDataSource" runat="server"></asp:TextBox></td>
                </tr>
                <tr runat="server" id="rwDataBase" visible="false">
                    <td>Please enter DataBase</td>
                    <td>
                        <asp:TextBox ID="txtDataBase" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr runat="server" id="trUserName" visible="false">
                    <td>Please enter username</td>
                    <td>
                        <asp:TextBox ID="txtUserName" runat="server"></asp:TextBox>
                    </td>
                </tr>
                 <tr runat="server" id="trpassword" visible="false">
                    <td>Please enter password</td>
                    <td>
                        <asp:TextBox ID="txtPassword" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr runat="server" id="rwExportToDb" visible="false">
                    <td colspan="2" align="center">
                        <asp:Button ID="brnExport" runat="server" Text="Export" OnClick="brnExport_Click" CssClass="button" /></td>

                </tr>

            </table>

        </div>
    </form>
</body>
</html>
