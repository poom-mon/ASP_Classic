<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<title>Untitled Document</title>
<!--#include virtual="/silkspan_ssl/inc/conn.asp"-->
<%

 
  'if( Request("TabType") = "excel") then
    tb = "<table>" & Request("objhtml") & "</table>"
    tb = Replace(tb,"<input type=""text"" class=""form-control"" placeholder=""","")
    tb = Replace(tb,""" disabled="""">","") 
  'end if

    xlsName = Request("xlsName")
    Response.Write(tb)
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition","attachment;filename="& xlsName &".ods"
      
%>
</head>

<body>
   
</body>
</html> 
<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->