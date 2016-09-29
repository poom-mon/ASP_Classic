<%
'	response.Clear

'	response.write("<br><br><br><br><br><br><br><br><br><br><br><br>")
'	response.write("<div align='center'>")
'	response.write("<font color='#FF0000'>Sorry....website is on maintainance</font>")
'	response.write("<div>")
'	response.end

'if instr(lcase(request("typedealer")), "yimyim") <> 0 then
'	response.end
'end if

ConnStr="Provider=SQLOLEDB.1;Persist Security Info=False;User ID=webapp;Password=1nt3l1n5id3;Initial Catalog=silkspan;Data Source=192.168.10.103"
ConnStr3="Provider=SQLOLEDB.1;Persist Security Info=False;User ID=webapp;Password=1nt3l1n5id3;Initial Catalog=silkspan;Data Source=192.168.10.103"


'ConnStr = "Driver={SQL Server};Server=SERVER02;Database=silkspan;UID=sa;PWD=password"
Set Conn = Server.createObject("ADODB.Connection")
Conn.Open ConnStr
Conn.CursorLocation=3
'ConnStr = "Driver={SQL Server};Server=SERVER02;Database=silkspan;UID=sa;PWD=password"
Set Conn2 = Server.createObject("ADODB.Connection")
Conn2.Open ConnStr
Conn2.CursorLocation=3

'ConnStr3 = "Driver={SQL Server};Server=SERVER02;Database=silkspan;UID=sa;PWD=password"
Set Conn3 = Server.createObject("ADODB.Connection")
Conn3.Open ConnStr3
Conn3.CursorLocation=3
	
'If InStr(Request.ServerVariables("SCRIPT_NAME"), "showdispatch_excel.asp") <> 0 Or InStr(Request.ServerVariables("SCRIPT_NAME"), "showinfo.asp") <> 0 Or InStr(Request.ServerVariables("SCRIPT_NAME"), "cusfile") <> 0 Or InStr(Request.ServerVariables("SCRIPT_NAME"), "printpdf") <> 0 then
'	log_conn_date = year(date) & Right("0"& month(date), 2) & Right("0"& day(date), 2)

'	sql = " insert into log_connection_"& log_conn_date &" (new_id, date, ipaddress, url, url_referrer, query_string, remote_host, user_agent) values (newid(), getdate(), '"& Left(Request.ServerVariables("REMOTE_ADDR"), 500) &"', '"& Left(Request.ServerVariables("SCRIPT_NAME"), 500) &"', '"& Left(Request.ServerVariables("HTTP_REFERER"), 500) &"', '"& Left(Request.ServerVariables("QUERY_STRING"), 500) &"', '"& Left(Request.ServerVariables("REMOTE_HOST"), 100) &"', '"& Left(Request.ServerVariables("HTTP_USER_AGENT"), 500) &"') "
'	conn.execute sql
'End if
%>
<!--#include virtual="/inc/nop/dug_script.asp"-->