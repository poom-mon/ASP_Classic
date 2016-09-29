<!--#include virtual="/silkspan_ssl/inc/conn.asp"-->
<%
conn.CommandTimeout = 1000
server.scripttimeout = 500
 

Function Dmy(val,vtype)
    if vtype = "ymd" then
		If isdate(val) Then
		 Dmy =  Year(val)&"-"&right("0"&month(val),2)&"-"&right("0"&day(val),2)
		 else
		 Dmy = Year(now)&"-"&right("0"&month(now),2)&"-"&right("0"&day(now),2)
		end if
	else
		If isdate(val) Then
		 Dmy = right("0"&day(val),2)&"-"&right("0"&month(val),2)&"-"& Year(val)
		 else
		 Dmy = right("0"&day(now),2)&"-"&right("0"&month(now),2)&"-"& Year(now)
		end if
	end if
End Function

Function FormatDate(val)
	If val <> "" Then
     FormatDate = right(val,4)&"-"&mid(val,4,2)&"-"&left(val,2)
	Else
	 FormatDate = year(now)&"-"&right("0"&month(now),2)&"-"&right("0"&day(now),2)
	End If
End Function
  
CALL SENDMAIL()  

SUB SENDMAIL()
        
	    strmail = strmail & "<style type='text/css'>"
		strmail = strmail & ".Title"
		strmail = strmail & "{"
		strmail = strmail & "	font-family: 'Microsoft Sans Serif';"
		strmail = strmail & "	font-size: 10pt;"
		strmail = strmail & "	color: #232323;"
		strmail = strmail & "}"
		strmail = strmail & ".Detail"
		strmail = strmail & "{"
		strmail = strmail & "	font-family: 'Microsoft Sans Serif';"
		strmail = strmail & "	font-size: 9pt;"
		strmail = strmail & "	color: #232323;"
		strmail = strmail & "	height: 25;"
		strmail = strmail & "}"
		strmail = strmail & "</style>"
        strmail = strmail & "<br>"
        strmail = strmail & "<table width='740' border='0' cellpadding='2' cellspacing='1' bgcolor='#003366' align='center'>"
		strmail = strmail & "<tr bgcolor='#cccccc' align='left'>"
		strmail = strmail & "<td class='Title' colspan='7'>DETAIL DESTROYING</td>"
		strmail = strmail & "</tr>"
		strmail = strmail & "<tr bgcolor='#ffffff' height='2'>"
		strmail = strmail & "<td colspan='7'></td>"
		strmail = strmail & "</tr>"
		strmail = strmail & "<tr bgcolor='#cccccc' align='center'>"
		strmail = strmail & "<td class='Title' width='50'>NO.</td>"
		strmail = strmail & "<td class='Title' width='60'>TID</td>"
		strmail = strmail & "<td class='Title' width='160'>NAME - LASTNAME</td>"
		strmail = strmail & "<td class='Title' width='120'>PRODUCT</td>"
		strmail = strmail & "<td class='Title' width='120'>APPLY DATE</td>"
		strmail = strmail & "<td class='Title' width='150'>CANCEL DATE</td>"
		strmail = strmail & "<td class='Title' width='80'>STATUS</td>"
		strmail = strmail & "</tr>"
		
		'วันจันทร์ ต้องถอยไป 2 วัน
		num = 0
		sql = "select * from ("&strsql&") as a where retdoc is null "&_
		"and ( ((t_status = 'c' or t_status = 'r') "&_
		"and year(t_editdate) >= 2011 "&_
		"and appreturn_msg is not null "&_
		"and t_editdate is not null and datediff(d,convert(varchar,year(dateadd(m,-1,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-1,getdate()))),2)+'-01',t_editdate) >= 0 "&_
		"and datediff(d,getdate() - 9,dateadd(d,-2,t_editdate)) < 0) or (t_status is null "&_
		"and appreturn is null and incomplete is not null "&_
		"and datediff(d,convert(varchar,year(dateadd(m,-3,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-3,getdate()))),2)+'-01',incomplete) >= 0 "&_
		"and datediff(d,getdate() - 62,dateadd(d,-2,incomplete)) < 0 and datediff(d,'2010-11-04',incomplete) >= 0) ) "&_
		"order by name asc,applicationdate asc" 'and appreturn is null 
		'response.write sql
		'response.flush
        'Response.Write(sql)
		'rs.open sql,conn,0
		'if not (rs.bof and rs.eof) then
		'	do while not rs.eof
		'		 num = num + 1
		'		if num mod 2 = 1 then
		'		ResetColor = color2
		'		else
		'		ResetColor = color3
		'		end if
		'	
		'		strmail = strmail & "<tr bgcolor='#CEE7FF' align='center'>"
		'		strmail = strmail & "<td class='Detail'>"&num&".</td>"
		'		strmail = strmail & "<td class='Detail'>"&trim(rs("tid"))&"</td>"
		'		strmail = strmail & "<td class='Detail' align='left'>&nbsp;&nbsp;"&trim(rs("name"))&"</td>"
		'		strmail = strmail & "<td class='Detail'>"&FncProduct(trim(rs("tbname")))&"</td>"
		'		strmail = strmail & "<td class='Detail'>"&Dmy(rs("applicationdate"),"")&"</td>"
		'		if isnull(rs("t_status")) then
		'		strmail = strmail & "<td class='Detail'>"&Dmy(rs("t_editdate"),"")&"</td>"
		'		else
		'		strmail = strmail & "<td class='Detail'>"&Dmy(rs("incomplete"),"")&"</td>"
		'		end if
		'		if trim(rs("t_status")) = "c" then
		'		strmail = strmail & "<td class='Detail'>Cancel</td>"
		'		elseif trim(rs("t_status")) = "r" then
		'		strmail = strmail & "<td class='Detail'>Reject</td>"
		'		else
		'		strmail = strmail & "<td class='Detail'>Incomplete</td>"
		'		end if
		'		strmail = strmail & "</tr>"
		' 
		'		 rs.movenext
		'	loop
        '
		'else
			
			strmail = strmail & "<tr bgcolor='"&color3&"' align='center'>"
			strmail = strmail & "<td class='Title' colspan='7'><font color='#FF0000'>Empty record !!!</font></td>"
			strmail = strmail & "</tr>"
			
		'end if
		'rs.close

 
        strmail = strmail & "</table>"
		response.write strmail


        
 'if mail = "y" and num <> 0 then
           
		 strmail = replace(strmail, "<", ".[.")
         strmail = replace(strmail, ">", ".].")

		 Recipient = "sittiporn@silkspan.local"  
	 %>

     	<iframe name="fmmail" width="660" height="50" frameborder="0" align="center" style="display:none"></iframe>
		<form name='Fm' method='post'>
		<input type='Hidden' name='mailForm' value="internal@silkspan.local">
		<input type='Hidden' name='mailType' value="internal"> 
		<input type='Hidden' name='mailSend' value="<%=Recipient%>"> 
		<input type='Hidden' name='subject' value="เอกสารที่ยังไม่ถูกทำลาย <%=Dmy(now,"")%>">
		<input type='Hidden' name='mailbody' value="<%=strmail%>">
		<script language="javascript">
		    document.Fm.target = "_blank";// "fmmail";
		    document.Fm.method = "post";
		    document.Fm.action = "http://www.silkspan.com/v2/test/SendMail.aspx";
		    document.Fm.submit();  

	    </script>
	    </form> 

	<!--<iframe name="fmmail" width="660" height="50" frameborder="0" align="center" style="display:none"></iframe>
		<form name='Fm' method='post'>
		<input type='Hidden' name='Sender' value="internal@silkspan.local">
		<input type='Hidden' name='MailServerUserName' value="internal@silkspan.local">
		<input type='Hidden' name='MailServerPassword' value="internalss">
		<input type='Hidden' name='Recipient' value="<%=Recipient%>">
		<input type='Hidden' name='RecipientCC' value="">
		<input type='Hidden' name='RecipientBCC' value="">
		<input type='Hidden' name='Subject' value="เอกสารที่ยังไม่ถูกทำลาย <%=Dmy(now,"")%>">
		<input type='Hidden' name='Body' value="<%=strmail%>">
		<script language="javascript">
		    document.Fm.target = "_blank"; //'"fmmail";
		    document.Fm.method = "post";
		    document.Fm.action = "http://www.silkspan.com/sendmail/sender.aspx"; // "http://192.168.0.33/sendmail/sender_local.aspx";
		    document.Fm.submit();  

	    </script>
	    </form>-->
    <%
'end if

set rs = nothing 

END SUB

 

%> 

<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->
 