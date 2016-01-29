<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="BulkFileUpload.aspx.cs" Inherits="excelImportAsp.BulkFileUpload" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h3>Code Scratcher</h3>
        <p>

            <asp:Button ID="btnUpload" runat="server" Text="SQL Bulk Upload" OnClick="btnUpload_Click" />

        </p>


    </div>
        <asp:Label ID="lbl_Message" runat="server" Visible="false"></asp:Label>
    </form>
</body>
</html>
