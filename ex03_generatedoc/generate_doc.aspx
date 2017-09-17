<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="generate_doc.aspx.vb" Inherits="ex03_generatedoc.generate_doc" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button runat="server" ID="btnGen" Text="產生DOC" />
        <asp:Button runat="server" ID="btnValidate" Text="驗證" />
        <asp:Label runat="server" ID="lblMesg" />
    </div>
    </form>
</body>
</html>
