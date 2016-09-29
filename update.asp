<!--#include virtual="/silkspan_ssl/inc/conn.asp"-->
<html>
<title>Update Data</title>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">	
<style type="text/css">
	Body
	{
		margin-left: 5;
		margin-top: 0;
		margin-right: 5;
		margin-bottom: 5;
	}
	.Title
	{
		font-family: "Microsoft Sans Serif";
		font-size: 10pt;
		color: #FFFFFF;
	}
	.Titlemail
	{
		font-family: "Microsoft Sans Serif";
		font-size: 10pt;
		color: #232323;
	}
	.TitleTB{
		font-family: "Microsoft Sans Serif";
		font-size: 10pt;
		color: #232323;
		height: 25;
	}
	.Titlep{
		font-family: "CordiaUPC";
		font-size: 16px;
		font-weight:normal
		color: #232323;
	}
	.Detail
	{
		font-family: "Microsoft Sans Serif";
		font-size: 9pt;
		color: #232323;
		height: 25;
	}
	.ButtonBlue
	{
		background-color: #DDEEFF;
		border: 1 solid #000000;
		color: #003366;
		font-family: "Microsoft Sans Serif";
		font-size: 8pt;
		height: 19;
		cursor: pointer;
	}
	.ButtonViolet
	{
		background-color: #E1E1FF;
		border: 1 solid #000000;
		color: #333366;
		font-family: "Microsoft Sans Serif";
		font-size: 8pt;
		height: 19;
		cursor: pointer;
	}
	.ButtonBrown
	{
		background-color: #FFE1DB;
		border: 1 solid #606060;
		color: #330000;
		font-family: "Microsoft Sans Serif";
		font-size: 8pt;
		height: 19;
		cursor: pointer;
	}
	.ButtonOrange
	{
		background-color: #FFE2C6;
		border: 1 solid #993300;
		color: #993300;
		font-family: "Microsoft Sans Serif";
		font-size: 8pt;
		height: 19;
		cursor: pointer;
	}
	.ButtonRed
	{
		background-color: #FFCCCC;
		border: 1 solid #660000;
		color: #003366;
		font-family: "Microsoft Sans Serif";
		font-size: 8pt;
		height: 19;
		cursor: pointer;
	}
	.ButtonGreen
	{
		background-color: #ECFFEC;
		border: 1 solid #003300;
		color: #003300;
		font-family: "Microsoft Sans Serif";
		font-size: 8pt;
		height: 19;
		cursor: pointer;
	}
	.Box
	{
		border: #7F9DB9 1px solid;
		font-family: "Microsoft Sans Serif";
		font-size: 8pt; 
		color: #333333; 
		height: 19;
		text-align: left;
	}
</style>
</head>
<body>
<%
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

set rs = server.createobject("adodb.recordset") 
tid = trim(request("tid"))
tabType = trim(Request("TabType")) 
  Response.Write("<script>parent.fncAlert('" & tid & "');</script>")
   
  strsql ="" 
    sql = "select tablename from cc_condition_data "&_
    "where inbound_status = 'on' and tablename not like '%mg_%' "&_
    "and tablename not like '%car_%' and tablename not like '%citi%' order by bank asc"
    rs.open sql,conn,0
    if not(rs.bof and rs.eof) then
       while not rs.eof
         strsql = strsql & "(select name+'  '+lastname as name,tid,applicationdate"&_
	     ",incomplete,appreturn,appreturn_msg,t_status,t_editdate,activeby,retdoc"&_
	     ",'"&trim(rs("tablename"))&"' as tbname from "&trim(rs("tablename"))&") union "
       rs.movenext
       wend
    end if
    rs.close

    if len(strsql) <> 0 then
    strsql = left(strsql,len(strsql) - 7)
    end if

    str = "" 
     'INCOMPLETE
     sql = "select * from ("&strsql&") as a where retdoc is null and t_status is null "&_
	 "and incomplete is not null and appreturn is null "&_
	 "and datediff(d,convert(varchar,year(dateadd(m,-3,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-3,getdate()))),2)+'-01',incomplete) >= 0 "&_
	 "and datediff(d,getdate() - 60,incomplete) < 0 "&_
	 "and datediff(d,'2010-11-04',incomplete) >= 0 "&_
     " and year(incomplete) = year(getdate())  "&_
	 "order by name asc,incomplete asc"
      rs.open sql,conn,0
      str =  str & " INCOMPLETE : \n"
      if not (rs.bof and rs.eof) then
	       do while not rs.eof
             str = str  & "\n APPLY DATE  : " & Dmy(rs("applicationdate"),"") &"  INCOMPLETE: " & Dmy(rs("incomplete"),"") & " : " & trim(rs("name"))
            rs.movenext
		   loop 
      else
	        str = str &  "  Empty record !!!"  
	 end if
	 rs.close
    
    '-REJECT 
     sql = "select * from ("&strsql&") as a where retdoc is null and t_status = 'r' "&_
	 "and appreturn_msg is not null "&_
	 "and t_editdate is not null and datediff(d,convert(varchar,year(dateadd(m,-1,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-1,getdate()))),2)+'-01',t_editdate) >= 0 "&_
	 "and datediff(d,getdate() - 7,t_editdate) < 0 "&_
	 "and year(t_editdate) >= 2011 "&_
     "and year(t_editdate) = year(getdate()) "&_  
	 "order by name asc,t_editdate asc" 'and appreturn is null 
      rs.open sql,conn,0
      str =  str & " REJECT : \n"
		if not (rs.bof and rs.eof) then
			do until rs.eof
                  str = str  & "\n APPLY DATE  : " & Dmy(rs("applicationdate"),"") &"  INCOMPLETE: " & Dmy(rs("t_editdate"),"") & " : " & trim(rs("name"))
               rs.movenext
			loop 
		else
          str = str &  "  Empty record !!!"  
        end if
		rs.close 
      
     ''CANCEL
     sql = "select * from ("&strsql&") as a where retdoc is null and t_status = 'c' "&_
	 "and appreturn_msg is not null "&_
	 "and t_editdate is not null and datediff(d,convert(varchar,year(dateadd(m,-1,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-1,getdate()))),2)+'-01',t_editdate) >= 0 "&_
	 "and datediff(d,getdate() - 7,t_editdate) < 0 "&_
	 "and year(t_editdate) >= 2011 "&_
      " and year(t_editdate) = year(getdate()) "&_
	 "order by name asc,t_editdate asc" 'and appreturn is null 
       rs.open sql,conn,0
        str =  str & " CANCEL : \n"
	 if not (rs.bof and rs.eof) then
	 	do until rs.eof
                str = str  & "\n APPLY DATE  : " & Dmy(rs("applicationdate"),"") &"  INCOMPLETE: " & Dmy(rs("t_editdate"),"") & " : " & trim(rs("name"))
             rs.movenext
	 	loop 
	 else
        str = str &  "  Empty record !!!"  
     end if
	 rs.close 
     Response.Write("<script>parent.fncResponeListDestroy('" & str & "');</script>")



'tid = trim(request("tid"))
'
'if tid <> "" then
'		sql = "select tablename from trans as a inner join cc_condition_data as b "&_
'		"on a.pid = b.pid where a.tid = "& tid
'		rs.open sql,conn,0
'		if not(rs.bof and rs.eof) then
'           tbname = trim(rs("tablename"))
'        else
'          response.write("<script language='javascript'>")
'		  response.write("parent.document.getElementById('tr_loading').style.display = 'none';")
'		  response.write("parent.document.getElementById('txt_tid').style.background='#ffcccc';")
'		  response.write("parent.document.getElementById('td_show').innerHTML='&nbsp;&nbsp;<font color=#ff0000>Not Found Data In Trans !!!</font>';")
'	      response.write("</script>")
'		end if
'		rs.close
'        
'
'        if tbname <> "" then
'
'			sql = "select name,lastname,t_status,applicationdate,incomplete,t_editdate,"&_
'			"retdoc,appreturn,appreturn_msg from "&tbname&" "&_
'			"where retdoc is null "&_
'			"and ((t_status is null and incomplete is not null) or "&_
'			"(t_status is not null and t_editdate is not null and appreturn_msg is not null)) "&_
'			"and tid = "& tid  ' and appreturn is null
'			rs.open sql,conn,0
'			if not(rs.bof and rs.eof) then
'			   name = trim(rs("name")) & " " & trim(rs("lastname"))
'			   apply = trim(rs("applicationdate"))
'			   if lcase(trim(rs("t_status"))) = "c" then
'				status = "Cancel"
'				statusdate = trim(rs("t_editdate"))
'			   elseif lcase(trim(rs("t_status"))) = "r" then
'				status = "Reject"
'				statusdate = trim(rs("t_editdate"))
'			   else
'				status = "Incomplete"
'				statusdate = trim(rs("incomplete"))
'			   end if
'               
'			  
'					sql = "select top 1 * from log_destroy where tid = "& tid
'					rslog.open sql,conn,0
'					if (rslog.bof and rslog.eof) then
'						rslog.close  
'						sql = "select top 1 * from log_destroy order by id desc"
'						rslog.open sql,conn,1,2
'						if not(rslog.bof and rslog.eof) then
'							id = rslog("id") + 1
'						else
'							id = 1
'						end if
'						rslog.addnew
'						rslog("id") = id
'						rslog("tid") = tid
'						rslog("name") = name
'						rslog("tbname") = tbname
'						rslog("status") = status
'						rslog("applydate") = cdate(apply)
'						rslog("statusdate") = cdate(statusdate)
'						rslog("logdate") = now
'						rslog.update
'						rslog.close
'					else
'					   rslog.close
'					end if
'
'					conn.execute "update "&tbname&" set retdoc = 'Destroy' "&_
'					"where (retdoc is null or retdoc = '') and tid = "& tid
'
'
'					response.write("<script language='javascript'>")
'					response.write("parent.document.getElementById('tr_loading').style.display = 'none';")
'					 response.write("parent.document.getElementById('txt_tid').style.background='#ffffff';")
'					 response.write("parent.document.getElementById('txt_tid').value = '';")
'					 response.write("parent.document.getElementById('td_show').innerHTML= parent.document.getElementById('"&status&"num"&tid&"').innerHTML + parent.document.getElementById('"&status&"name"&tid&"').innerHTML + '&nbsp;&nbsp;&nbsp;<font color=#003399>("&status&")</font>';")
'					 response.write("parent.document.getElementById('"&status&tid&"').innerHTML = '<font color=#003399>Complete</font>';")
'					 response.write("</script>")
'			
'			else 'not(rs.bof and rs.eof)
'
'				response.write("<script language='javascript'>")
'				response.write("parent.document.getElementById('tr_loading').style.display = 'none';")
'				response.write("parent.document.getElementById('txt_tid').style.background='#ffcccc';")
'				response.write("parent.document.getElementById('td_show').innerHTML='&nbsp;&nbsp;<font color=#ff0000>Not Found Data In Product !!!</font>';")
'				response.write("</script>")
'
'			end if 'not(rs.bof and rs.eof)
'			rs.close
'	   
'	   end if 'tbname <> ""
'
'	   set rs = nothing
'       set rslog = nothing
'
'end if 'Tid <> ""
%>
</body>
</html>
<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->