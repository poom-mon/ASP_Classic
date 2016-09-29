<!--#include virtual="/silkspan_ssl/inc/conn.asp"-->
<html>
<title>Update Data</title>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">	
 
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
Function fncMonth(month)

		   select case month
           case "1" : vProduct = "มกราคม"
		   case "2" : vProduct = "กุมภาพันธ์"
		   case "3" : vProduct = "มีนาคม"
		   case "4" : vProduct = "เมษายน"
		   case "5" : vProduct = "พฤษภาคม"
		   case "6" : vProduct = "มิถุนายน"
		   case "7" : vProduct = "กรกฎาคม"
		   case "8" : vProduct = "สิงหาคม"
		   case "9" : vProduct = "กันยายน"
		   case "10" : vProduct = "ตุลาคม"
		   case "11" : vProduct = "พฤศจิกายน"
           case "12" : vProduct = "ธันวาคม"
		   end select 

		   fncMonth = vProduct
End  Function


set rs = server.createobject("adodb.recordset") 
set rslog  = server.createobject("adodb.recordset") 
tid = trim(request("tid")) 
TabType = trim(request("TabType"))  

'Response.Write("<script>alert('" & TabType & "');</script>")
   'Response.Write("<script>parent.fncResponeListDestroy('" & TabType & "');</script>")
    
    FUNCTION fncQueryTableAllByCondition()
        strsql ="" 
        '' ดึง table ทั้งหมดที่เกียวข้องตาม tid ที่ค้นหา
         sql = "select  tablename,name  from cc_condition_data "&_
         "where inbound_status = 'on' and tablename not like '%mg_%' "&_
         "and tablename not like '%car_%' and tablename not like '%citi%' order by bank asc"
         rs.open sql,conn,0
         if not(rs.bof and rs.eof) then
            while not rs.eof
              strsql = strsql & "(select name+'  '+lastname as name,tid,applicationdate"&_
	          ",incomplete,appreturn,appreturn_msg,t_status,t_editdate,activeby,retdoc,qcby"&_
	          ",'"&trim(rs("tablename"))&"' as tbname ,'"&trim(rs("name"))&"' as  crdname  from "&trim(rs("tablename"))&" where tid = " & tid & ") union "
            rs.movenext
            wend
         end if
         rs.close 
         if len(strsql) <> 0 then
              strsql = left(strsql,len(strsql) - 7)
         end if
         fncQueryTableAllByCondition = strsql 
     end function
      
      FUNCTION fncQueryTableAll()
        strsql ="" 
        '' ดึง table ทั้งหมดที่เกียวข้อง
         sql = "select tablename from cc_condition_data "&_
         "where inbound_status = 'on' and tablename not like '%mg_%' "&_
         "and tablename not like '%car_%' and tablename not like '%citi%' order by bank asc"
         rs.open sql,conn,0
         if not(rs.bof and rs.eof) then
            while not rs.eof
              strsql = strsql & "(select top 100  name+'  '+lastname as name,tid,applicationdate"&_
	          ",incomplete,appreturn,appreturn_msg,t_status,t_editdate,activeby,retdoc"&_
	          ",'"&trim(rs("tablename"))&"' as tbname from "&trim(rs("tablename"))&") union "
            rs.movenext
            wend
         end if
         rs.close 
         if len(strsql) <> 0 then
              strsql = left(strsql,len(strsql) - 7)
         end if
         fncQueryTableAll = strsql 
     end function

     FUNCTION fncRenderRowScan(rs,no,Status,StatusDestroy) 
      '' render html row scan 
        strtd = ""
        strtd =strtd &  "<div class=""td"">" & no & "</div>"
        strtd =strtd &  "<div class=""td"">" & rs("tid") & "</div>"
        strtd =strtd &  "<div class=""td"">" & rs("name") & "   " & rs("lastname") & "</div>"
        strtd =strtd &  "<div class=""td"">" &  rs("crdname")  & "</div>"
        strtd =strtd &  "<div class=""td"">" & Dmy(rs("applicationdate"),"") & "</div>"
        strtd =strtd &  "<div class=""td"">" & Status & "</div>"
        strtd =strtd &  "<div class=""td"">" & Dmy(rs("t_editdate"),"")   & "</div>"
        strtd =strtd &  "<div class=""td"">"& StatusDestroy&"</div>"  
        strtd =strtd &  "<div class=""td"">" & rs("qcby") & "</div>"
        fncRenderRowScan =  strtd 
     END FUNCTION

      FUNCTION fncRenderRowScanEmp(rs,no,Status,StatusDestroy) 
      '' render html row scan 
        strtdem = ""
        strtdem =strtdem &  "<div class=""td"">" & no & "</div>"
        strtdem =strtdem &  "<div class=""td"">" & rs("tid") & "</div>"
        strtdem =strtdem &  "<div class=""td"">" & rs("name") &  "</div>"
        strtdem =strtdem &  "<div class=""td"">" &  rs("crdname")  & "</div>"
        strtdem =strtdem &  "<div class=""td"">" & Dmy(rs("create_date"),"") & "</div>"
        strtdem =strtdem &  "<div class=""td"">" & Status & "</div>"
        strtdem =strtdem &  "<div class=""td"">" & Dmy(rs("statusdate"),"")   & "</div>"
        strtdem =strtdem &  "<div class=""td"">"& StatusDestroy&"</div>"  
        strtdem =strtdem &  "<div class=""td"">" & rs("qcby") & "</div>"
        fncRenderRowScanEmp =  strtdem 
     END FUNCTION

      FUNCTION fncRenderRowDestroy(rs,no,Status) 
       '' render html row destroy 
        strtdd = ""
        strtdd =strtdd &  "<td>" & no & "</td>"
        strtdd =strtdd &  "<td>" & rs("tid") & "</td>"
        strtdd =strtdd &  "<td>" & rs("name") &"  "& rs("lastname")   & "</td>"
        strtdd =strtdd &  "<td>" & rs("crdname")  & "</td>" 
        strtdd =strtdd &  "<td>" & Dmy(rs("applicationdate"),"") & "</td>"
        strtdd =strtdd &  "<td>" & Status & "</td>"
        strtdd =strtdd &  "<td>" & Dmy(rs("t_editdate"),"")   & "</td>" 
           strtdd =strtdd &  "<td>" & rs("qcby") & "</td>"
        fncRenderRowDestroy =  strtdd 
      END FUNCTION
      
      FUNCTION fncRenderRowAllDestroy(rs,no,Status) 
       '' render html row destroy 
        strAllDes = ""
        strAllDes =strAllDes &  "<td>" & no & "</td>"
        strAllDes =strAllDes &  "<td>" & rs("tid") & "</td>"
        strAllDes =strAllDes &  "<td>" & rs("name") & "</td>"
        strAllDes =strAllDes &  "<td>" & rs("crdname")  & "</td>"  
        strAllDes =strAllDes &  "<td>" & Status & "</td>"
        strAllDes =strAllDes &  "<td>" & Dmy(rs("create_date"),"")   & "</td>" 
        strAllDes =strAllDes &  "<td>" & rs("qcby") & "</td>"
        fncRenderRowAllDestroy =  strAllDes 
      END FUNCTION


      FUNCTION fncRenderRowLogDestroy(rs,no,Status)   
        strdCdes = ""
        strdCdes =strdCdes &  "<td>" & no & "</td>"
        strdCdes =strdCdes &  "<td>" & rs("tid") & "</td>"
        strdCdes =strdCdes &  "<td>" & rs("name") & "</td>"
        strdCdes =strdCdes &  "<td>" & rs("crdname")  & "</td>"  
        strdCdes =strdCdes &  "<td>" & Status & "</td>"
        strdCdes =strdCdes &  "<td>" & Dmy(rs("create_date"),"")   & "</td>" 
        strdCdes =strdCdes &  "<td>" & rs("qcby") & "</td>" 
        fncRenderRowLogDestroy =  strdCdes 
      END FUNCTION

       
   FUNCTION fncRenderRowEmpDestroy(rs,no)   
        strRowEmp = ""
        strRowEmp =strRowEmp &  "<td>" & no & "</td>"
        strRowEmp =strRowEmp &  "<td>" & rs("tid") & "</td>"
        strRowEmp =strRowEmp &  "<td>" & rs("name") & "</td>"
        strRowEmp =strRowEmp &  "<td>" & rs("crdname")  & "</td>"  
        strRowEmp =strRowEmp &  "<td>" & rs("status") & "</td>"
        strRowEmp =strRowEmp &  "<td>" & Dmy(rs("statusdate"),"")   & "</td>" 
        strRowEmp =strRowEmp &  "<td>" & rs("qcby") & "</td>" 
         strRowEmp =strRowEmp &  "<td><input class=""form-control"" name=""chkDetroy"" data-tid=""" & rs("tid") & """  data-status=""" & rs("status") & """  data-name=""" & rs("name") & """  data-tbname=""" & rs("tbname") & """  type=""checkbox"" /></td>" 
        fncRenderRowEmpDestroy =  strRowEmp 
      END FUNCTION

      FUNCTION fncRenderRowsDetailPenalty(year)   
       
            'sql =" select  des.tid ,des.name,convert(varchar(30),des.create_date,103) as create_date ,cc.name crdName   "
            'sql =sql &" ,case WHEN   datediff(d,b.create_date,getdate()) > 0  THEN 210   "
            'sql =sql &"       ELSE 200       "
            'sql =sql &"  END  as Penalty      "
            'sql =sql &" from log_Approve_destroy des inner join cc_condition_data cc on(cc.tablename = des.tbname)   "
            'sql =sql &" left join log_destroy_UploadDoc b on(des.tid = b.tid)  "
            'sql =sql &" where  des.status_update= 'not-destroy'  and des.create_user='"& Session("usrname")  &"'   and des.Step = 2     "
            'sql =sql &"  and year(des.create_date) = year(getdate())   and  month(des.create_date) = "&year
        

          sql ="    select  a.name, a.tid,200 Penalty  ,convert(varchar(30),a.create_date,103) as create_date , 'ส่งเอกสารให้บัญชีทำลาย' as 'remark'    "
	      sql =sql &"    from log_Approve_destroy a   "
	      sql =sql &"    where  a.status_update= 'not-destroy'    "
	      sql =sql &"    and a.Step = 2  and a.create_user='"& Session("usrname")  &"'     "
	      sql =sql &"    and year(a.create_date) = year(getdate())   and  month(a.create_date) = "&year 
          sql =sql &"  union all  "
	      sql =sql &"   select a.name, b.tid,"
	      sql =sql &"     10   as Penalty "  
	      sql =sql &"    ,convert(varchar(30),a.create_date,103) as create_date   "
	      sql =sql &"   , 'บัญชีส่งเอกสารคืน SUP แต่ SUP  ยังไม่รับเข้าระบบ  ' as 'remark'     "
	      sql =sql &"    from log_Approve_destroy a  inner join log_destroy_UploadDoc b on(a.tid = b.tid)  "
	      sql =sql &"    where  a.status_update= 'not-destroy'    "
	      sql =sql &"    and a.Step = 2   and a.create_user='"& Session("usrname")  &"'    and b.update_date is not null  "
	      sql =sql &"    and year(a.create_date) = year(getdate()) "
	      sql =sql &"    and ((a.doc_Returns_date is null and datediff(d,b.update_date,getdate()) > 0  ) or datediff(d,b.update_date,a.doc_Returns_date) > 0)"
          sql =sql &"    and  month(a.create_date) = "&year
         ' sql =sql &"  union all  "
	     ' sql =sql &"   select a.name, a.tid,"
	     ' sql =sql &"    10   as Penalty   "
	     ' sql =sql &"    ,convert(varchar(30),a.create_date,103) as create_date  "
	     ' sql =sql &"   , ' SUP ส่งเอกสารคืน QC ยังไม่ได้รับเข้าระบบ ' as 'remark'     "
	     ' sql =sql &"    from log_Approve_destroy a   "
	     ' sql =sql &"    where  a.status= 'not-destroy'    "
	     ' sql =sql &"    and a.Step = 1   and a.create_user='"& Session("usrname")  &"' "
	     ' sql =sql &"    and year(a.create_date) = year(getdate()) "
         ' sql =sql &"    and  month(a.create_date) = "&year
          
          stRows=""
          countRw=0
          stRows = stRows &"<table class=""table"" style=""background-color:#F9F9F9;"" ><thead><tr> <th>Tid</th> <th>Name</th> <th>remark</th> <th>date</th> <th>Penalty</th> </tr></thead><tbody>"
            rslog.open sql,conn,0 
	        if not (rslog.bof and rslog.eof) then 
	 	       do until rslog.eof    
                   stRows = stRows &"<tr> <td>"& rslog("tid") &"</td><td>"& rslog("name") &"</td><td>"& rslog("remark") &"</td><td>"& rslog("create_date") &"</td><td>-"& rslog("Penalty") &"</td> </tr>" 
                   countRw = countRw + cint(rslog("Penalty"))
                    rslog.movenext
	 	       loop  
                stRows  =  stRows &"<tr> <td colspan=""4"" style=""text-align:center;""><b>รวม</b></td><td><b> - "&countRw&"</b></td> </tr>"
                countRw=0
            end if 
             stRows = stRows &"</tbody></table>" 
	        rslog.close 
            fncRenderRowsDetailPenalty  = stRows
      END FUNCTION 

     FUNCTION fncRenderRowsDetailPenaltyEmp(year)   
       
        sql =  " select   a.tid,b.name, convert(varchar(30),b.empUpdatedate,103)  as create_date ,cc.name crdName "
        sql =sql & " from log_destroy_UploadDoc  a  inner join log_Approve_destroy b on(a.tid = b.tid)"
        sql =sql & " inner join cc_condition_data cc on(cc.tablename = b.tbname)"
        sql =sql &" where  a.create_user='"& Session("usrname")  &"'     "
        sql =sql & " and year(b.empUpdatedate) = year(getdate()) and datediff(d,b.empUpdatedate,a.update_date) > 0  and  month(b.empUpdatedate) = "&year  

        'Response.Write(sql)
          stRows=""
          stRows = stRows &"<table class=""table"" style=""background-color:#F9F9F9;"" ><thead><tr> <th>Tid</th> <th>Name</th> <th>Product</th> <th>date</th> <th>Penalty</th> </tr></thead><tbody>"
            rslog.open sql,conn,0 
             countRw=0
	        if not (rslog.bof and rslog.eof) then 
	 	       do until rslog.eof    
                   stRows = stRows &"<tr> <td>"& rslog("tid") &"</td><td>"& rslog("name") &"</td><td>"& rslog("crdName") &"</td><td>"& rslog("create_date") &"</td><td>-50</td> </tr>" 
                   countRw = countRw + 50
                    rslog.movenext
	 	       loop   
                stRows  =  stRows &"<tr> <td colspan=""4"" style=""text-align:center;""><b>รวม</b></td><td><b> - "&countRw&"</b></td> </tr>"
                countRw=0
            end if 
             stRows = stRows &"</tbody></table>" 
	        rslog.close 
            fncRenderRowsDetailPenaltyEmp  = stRows
      END FUNCTION 
      
      FUNCTION fncIncompleteSQl()  
        sqlIncompleteSql = "select  top 100 * from  VW_DestroyList  where retdoc is null and t_status is null "&_
	     "and incomplete is not null and appreturn is null "&_
	     "and datediff(d,convert(varchar,year(dateadd(m,-3,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-3,getdate()))),2)+'-01',incomplete) >= 0 "&_
	     "and datediff(d,getdate() - 60,incomplete) < 0 "&_
	     "and datediff(d,'2010-11-04',incomplete) >= 0 "&_
         " and (datediff(m,applicationdate,getdate()) between 0 and 3)  and tid not  in (select tid from log_approve_destroy)  "&_
	     "order by name asc,incomplete asc"
         
        fncIncompleteSQl = sqlIncompleteSql
      END FUNCTION 

      FUNCTION fncRejectSQl()
        '' render sql reject sql 
         sqlrejectSql = "select  top 100  * from VW_DestroyList   where retdoc is null and t_status = 'r' "&_
	     "and appreturn_msg is not null "&_
	     "and t_editdate is not null and datediff(d,convert(varchar,year(dateadd(m,-1,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-1,getdate()))),2)+'-01',t_editdate) >= 0 "&_
	     "and datediff(d,getdate() - 7,t_editdate) < 0 "&_
	     "and year(t_editdate) >= 2011 "&_
         "and (datediff(m,applicationdate,getdate()) between 0 and 3)  and tid not  in (select tid from log_approve_destroy) "&_  
	     "order by name asc,t_editdate asc"  
         fncRejectSQl =sqlrejectSql 
     END FUNCTION 

      FUNCTION fncCancelSQl()
          '' render sql cancel sql 
          sqlcancelSql ="select  top 100  * from  VW_DestroyList  where retdoc is null and t_status = 'c' "&_
	     "and appreturn_msg is not null "&_
	     "and t_editdate is not null and datediff(d,convert(varchar,year(dateadd(m,-1,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-1,getdate()))),2)+'-01',t_editdate) >= 0 "&_
	     "and datediff(d,getdate() - 7,t_editdate) < 0 "&_
	     "and year(t_editdate) >= 2011 "&_
          " and (datediff(m,applicationdate,getdate()) between 0 and 3) and tid not  in (select tid from log_approve_destroy) "&_
	     "order by name asc,t_editdate asc" 'and appreturn is null 
          fncCancelSQl =sqlcancelSql 
     END FUNCTION 

     Sub fncUpdateLogAppDestroy(tid,name,tbname,status,apply,statusdate,typeEmp,usrname,qc)   
         if typeEmp = "emp" then 
               if(status ="not-destroy") then
                  conn.execute "update log_Approve_destroy set empUpdate_user = '"& usrname &"' ,empUpdatedate  =getdate() ,status_update='not-destroy',qcby = '" & qc &"',doc_returns ='N' , step=2  where  tid = "& tid   
                else
                   conn.execute "update log_Approve_destroy set empUpdate_user = '"& usrname &"' ,empUpdatedate  =getdate() ,status_update='Destroy',qcby = '" & qc &"',doc_returns ='N'  , step=2 where  tid = "& tid  
               	   conn.execute " update "&tbname&" set retdoc = 'Destroy'  where (retdoc is null or retdoc = '') and tid = "& tid 
               end if  
         else 
        	sql = "select top 1 * from log_Approve_destroy where tid = "& tid & " and  doc_returns_date is null"
        	rslog.open sql,conn,0
        	if (rslog.bof and rslog.eof) then
        		rslog.close  
        		sql = "select top 1 * from log_Approve_destroy  order by id desc"
        		rslog.open sql,conn,1,2
        		if not(rslog.bof and rslog.eof) then
        			id = rslog("id") + 1
        		else
        			id = 1
        		end if
        		rslog.addnew
        		rslog("id") = id
        		rslog("tid") = tid
        		rslog("name") = name
        		rslog("tbname") = tbname
        		rslog("status") = status
        		rslog("applydate") = cdate(apply)
                  
                if isnull(statusdate) or statusdate ="" then
                    statusdate =  now  
                end if   
        		rslog("statusdate") = cdate(statusdate)
        		rslog("create_user") = usrname
                rslog("create_date") = now
                rslog("qcby") = qc  
                rslog("step") = 1 
                ''if(status ="not-destroy") then
                    rslog("doc_returns") = "N"
               '' end if 

        		rslog.update
        		rslog.close
        	else   
                 
                ''if(status ="not-destroy") then
                   conn.execute "update log_Approve_destroy set update_user = '"& usrname &"' ,update_date ='" & now & "',step=1,doc_returns='N',status='"&status&"' where  tid = "& tid  
               '' else
               ''   conn.execute "update log_Approve_destroy set update_user = '"& usrname &"' ,update_date ='" & now & "',step=1 ,status='"&status&"' where  tid = "& tid  
               '' end if 
        	     rslog.close
        	end if 

        end if 

     end sub
     
     Sub fncUpdateLogApprove(tid,name,tbname,status,apply,statusdate) 
        	sql = "select top 1 * from log_destroy where tid = "& tid
        	rslog.open sql,conn,0
        	if (rslog.bof and rslog.eof) then
        		rslog.close  
        		sql = "select top 1 * from log_destroy order by id desc"
        		rslog.open sql,conn,1,2
        		if not(rslog.bof and rslog.eof) then
        			id = rslog("id") + 1
        		else
        			id = 1
        		end if
        		rslog.addnew
        		rslog("id") = id
        		rslog("tid") = tid
        		rslog("name") = name
        		rslog("tbname") = tbname
        		rslog("status") = status
        		rslog("applydate") = cdate(apply)
        		rslog("statusdate") = cdate(statusdate)
        		rslog("logdate") = now
        		rslog.update
        		rslog.close
        	else
        	   rslog.close
        	end if
        
        	conn.execute "update "&tbname&" set retdoc = 'Destroy' "&_
        	"where (retdoc is null or retdoc = '') and tid = "& tid 
     end sub

     SUB fncUpdateStatus(rs)
       tid = rs("tid")
       name=""
         tbname = trim(rs("tbname"))
        typeEmp = Session("typeEmp")
        usrname =Session("usrname") 
        status=""
        apply=""
         statusdate = ""   
         qc = ""
          if typeEmp = "emp" then
             qc = rs("qcby") 
             status = rs("status")
          else 
             name = trim(rs("name")) & "  " & trim(rs("lastname"))
             apply = trim(rs("applicationdate")) 
             qc = rs("qcby")  

             if lcase(trim(rs("statusDestroy"))) = "cancel" then
                status = "Cancel"
                statusdate = trim(rs("t_editdate"))
             elseif lcase(trim(rs("statusDestroy"))) = "reject" then
                status = "Reject"
                statusdate = trim(rs("t_editdate"))
             elseif lcase(trim(rs("statusDestroy"))) = "incomplete" then
                status = "Incomplete" 
                statusdate =  trim(rs("incomplete")) 
             elseif lcase(trim(rs("statusDestroy"))) = "complete" then
                status = "Complete" 
                statusdate =  trim(rs("incomplete")) 
            else 
                 status = "not-destroy"
                 statusdate = now
             end if 
       end if 
        
        Call  fncUpdateLogAppDestroy(tid,name,tbname,status,apply,statusdate,typeEmp,usrname,qc)  

    END SUB

     

    Select Case TabType
          Case "Scan"
           '*****************case scan
                no = 1   
                if(IsNumeric(tid)) then
                    strsql = fncQueryTableAllByCondition()   
                     sql = " SELECT * from tmp_destroy_product where TID = '" & tid & "'  and  statusDestroy is not null "
                      sql = sql & " and tid not in (select tid from log_Approve_destroy where tid = " & tid & ")"
                           rs.open sql,conn,0 
	                     if not (rs.bof and rs.eof) then 
	 	                    do until rs.eof 
                                    str = str & " <div class=""tr"">"
                                     str = str & fncRenderRowScan(rs,no,rs("statusDestroy"),"OK")
                                    str = str & "</div>"
                                no = no + 1
                               Call fncUpdateStatus(rs)

                                 rs.movenext
	 	                    loop   
                          end if 
	                     rs.close 
                         ccout =  1
                         if( len(str) <=  0 ) then 
                                 'sql = " SELECT * from tmp_destroy_product where TID = " & tid & "  and  statusDestroy is  null " 
                                 sql = "SELECT * from tmp_destroy_product where TID = '" & tid & "'  and  statusDestroy is  null "
                                 sql = sql & " and tid not in (select tid from log_Approve_destroy where tid = " & tid & ")"

                                   rs.open sql,conn,0  
                                   if not (rs.bof and rs.eof) then  
	                                    do while not rs.eof 
                                          str = str & " <div class=""tr"">"
                                           str = str & fncRenderRowScan(rs,no,"-","ทำลายไม่ได้")  
                                           str = str & " </div>"
                                            str = str & "<div style=""position:absolute;""><div class=""blink_second"" style=""visibility: visible; "">**เอกสารทำลายไม่ได้ส่งกลับด่วน</div></div>"
                                         no = no + 1 
                                        Call fncUpdateStatus(rs)
                                        ccout = 0
                                         rs.movenext
		                                loop 
                                        rs.close 
	                             else       
                                    str = str &  "<div class=""tr""><div class=""td"">ไม่พบข้อมูล</div></div>"  
                                 end if
                         end if 
                   else
                     str = str &  "<div class=""tr""><div class=""td"">ไม่พบข้อมูล</div></div>"  
                  end if
                Response.Write("<script>parent.fncResponeListScan('" & str & "');parent.fncRenderPopupconfirmDestroy("&ccout&");</script>")  
           
            Case "ScanEmp"
           '*****************case scan
                 no = 1   
                if(IsNumeric(tid)) then
                        strsql = fncQueryTableAllByCondition()   
                       sql = " select a.*,b.name as crdname from log_approve_destroy a inner join cc_condition_data b  on a.tbname = b.tablename  where a.tid= '" & tid &"'   and  a.status not like '%not-destroy%'   and empUpdatedate is null  "
                        Response.Write(sql)
                        rs.open sql,conn,0 
	                     if not (rs.bof and rs.eof) then  
	 	                    do until rs.eof 
                                    str = str & " <div class=""tr"">"
                                     str = str & fncRenderRowScanEmp(rs,no,rs("status"),"OK")
                                    str = str & "</div>"
                                no = no + 1
                               Call fncUpdateStatus(rs)

                                 rs.movenext
	 	                    loop   
                          end if 
	                     rs.close 

                         if( len(str) <=  0 ) then   
                                  sql = " select a.*,b.name as crdname from log_approve_destroy a inner join cc_condition_data b  on a.tbname = b.tablename  where a.tid= '" & tid &"'  and  a.status  like '%not-destroy%'    and empUpdatedate is   null  "
                                 'Response.Write(sql)
                                   rs.open sql,conn,0  
                                   if not (rs.bof and rs.eof) then  
	                                    do while not rs.eof 
                                          str = str & " <div class=""tr"">"
                                           str = str & fncRenderRowScanEmp(rs,no,rs("status"),"ทำลายไม่ได้")  ' fncRenderRowScan(rs,no,"-","ทำลายไม่ได้")  
                                           str = str & " </div>"
                                           str = str & "<div style=""position:absolute;""><div class=""blink_second"" style=""visibility: visible; "">**เอกสารทำลายไม่ได้ส่งกลับด่วน</div></div>"
                                         no = no + 1
                                        Call fncUpdateStatus(rs)
                                         ccout = 0
                                         rs.movenext
		                                loop 
                                        rs.close 
	                             else       
                                    str = str &  "<div class=""tr""><div class=""td"">ไม่พบข้อมูล</div></div>"  
                                 end if
                         end if
                  else
                     str = str &  "<div class=""tr""><div class=""td"">ไม่พบข้อมูล</div></div>"  
                  end if

                 Response.Write("<script>parent.fncResponeListScan('" & str & "');parent.fncRenderPopupconfirmDestroy("&ccout&");</script>")  
           


           '*****************end case scan 
          Case "DocDestroy"
           '*****************  case DocDestroy
                      no = 1     
                     sql = " SELECT TOP 200 * from tmp_destroy_product where    statusDestroy is not null and tid not in (select tid from log_Approve_destroy) "
                     rs.open sql,conn,0 
	                 if not (rs.bof and rs.eof) then 
	 	                do until rs.eof 
                                 str = str & " <tr>"
                                  str = str & fncRenderRowDestroy(rs,no,rs("statusDestroy")) 
                                  str = str & "</tr>"
                            no = no + 1
                             rs.movenext
	 	                loop   
                     end if

         
	                 rs.close 
                     if( len(str) <=  0 ) then    
                        str = str &  "<tr><td colspan=""7"">ไม่พบข้อมูล</td></tr>"   
                     end if  
                      Response.Write("<script>parent.fncResponeAllDocDesctroy('" & str & "');</script>")  
          '*****************end case DocDestroy
          Case "DocDesCon"
          '*****************  case DocDesCon 
                      sql = "select a.*,b.name as crdname from log_Approve_destroy a inner join "
                      sql = sql & " cc_condition_data b on(a.tbname = b.tablename) where  a.create_user='"& Session("usrname")  &"' and   a.status= 'not-destroy'  and a.doc_Returns_date is null "
                     ' sql = sql & " cc_condition_data b on(a.tbname = b.tablename) where  a.create_user='"& Session("usrname")  &"' and   a.status= 'not-destroy'  and a.doc_Returns_date is null and year(a.statusdate) = year(getdate()) "
                        
                       no =1
                     rs.open sql,conn,0 
	                 if not (rs.bof and rs.eof) then 
	 	                do until rs.eof 
                                 str = str & " <tr>"
                                  str = str & fncRenderRowLogDestroy(rs,no,"เอกสารทำลายไม่ได้") 
                                  str = str & "</tr>"
                            no = no + 1
                             rs.movenext
	 	                loop   
                     end if 
	                 rs.close 
                     if( len(str) <=  0 ) then    
                        str = str &  "<tr><td colspan=""7"">ไม่พบข้อมูล</td></tr>"   
                     end if  
                      Response.Write("<script>parent.fncResponeAllDocCon('" & str & "');</script>")  
            '*****************end case DocDesCon
          Case "AllDocDesCon"
            '*****************  case AllDocDesCon
             
                tid = Request("tid")
                nameFilter = Request("nameFilter")
                productFilter = Request("productFilter") 
                createDateFilter = Request("createDateFilter")
                qcFilter = Request("qcFilter")  

                conWhere = ""
                 if tid <> "" then
                    conWhere  = conWhere & " and a.tid like '%" & tid & "%'"
                end if 

                if nameFilter <> "" then
                    conWhere  = conWhere & " and a.name like '%" & nameFilter & "%'"
                end if 

                 if productFilter <> "" then
                    conWhere  = conWhere & " and b.name like '%" & productFilter & "%'"
                end if  
                 if createDateFilter <> "" then
                    conWhere  = conWhere & " and convert(varchar(20),a.create_date,105) like '" & createDateFilter & "%'"
                end if  
                 if qcFilter <> "" then
                    conWhere  = conWhere & " and a.qcby like '%" & qcFilter & "%'"
                end if 
                 
                 if(conWhere = "") then
                    sql = "select  top 100 a.*,b.name as crdname from log_Approve_destroy a inner join cc_condition_data b on(a.tbname = b.tablename) where   a.create_user='"& Session("usrname")  &"' and   a.status not like '%destroy%' "
                 else 
                   sql = "select a.*,b.name as crdname from log_Approve_destroy a inner join cc_condition_data b on(a.tbname = b.tablename) where   a.create_user='"& Session("usrname")  &"' and   a.status not like '%destroy%' "
                   sql = sql & conWhere
                 end if
                   Response.Write(sql)

                 '    sql = "select a.*,b.name as crdname from log_Approve_destroy a inner join cc_condition_data b on(a.tbname = b.tablename) where   a.create_user='"& Session("usrname")  &"' and   a.status not like '%destroy%' and year(a.statusdate) = year(getdate())"
                  no =1
                 rs.open sql,conn,0 
	             if not (rs.bof and rs.eof) then 
	 	            do until rs.eof 
                             str = str & " <tr>"
                              str = str & fncRenderRowLogDestroy(rs,no,"เอกสารที่รอการทำลาย") 
                              str = str & "</tr>"
                        no = no + 1
                         rs.movenext
	 	            loop   
                 end if 
	             rs.close 
                 if( len(str) <=  0 ) then    
                    str = str &  "<tr><td colspan=""7"">ไม่พบข้อมูล</td></tr>"   
                 end if  
                  Response.Write("<script>parent.fncResponeAllDocDesCon('" & str & "');</script>")  
             '*****************end case AllDocDesCon
               Case "AllDocDesConEmp"
            '*****************  case AllDocDesCon

                tid = Request("tid")
                nameFilter = Request("nameFilter")
                productFilter = Request("productFilter") 
                createDateFilter = Request("createDateFilter")
                qcFilter = Request("qcFilter")  

                conWhere = ""
                 if tid <> "" then
                    conWhere  = conWhere & " and a.tid like '%" & tid & "%'"
                end if 

                if nameFilter <> "" then
                    conWhere  = conWhere & " and a.name like '%" & nameFilter & "%'"
                end if 

                 if productFilter <> "" then
                    conWhere  = conWhere & " and b.name like '%" & productFilter & "%'"
                end if  
                 if createDateFilter <> "" then
                    conWhere  = conWhere & " and convert(varchar(20),a.create_date,105) like '" & createDateFilter & "%'"
                end if  
                 if qcFilter <> "" then
                    conWhere  = conWhere & " and a.qcby like '%" & qcFilter & "%'"
                end if 
                 
                 if(conWhere = "") then
                     sql = "select top 100 a.*,b.name as crdname from log_Approve_destroy a inner join cc_condition_data b on(a.tbname = b.tablename) where  a.empUpdate_user='"& Session("usrname")  &"'and a.empUpdate_user is not null and a.status_update = 'destroy'   and step=2"
                 else 
                    sql = "select a.*,b.name as crdname from log_Approve_destroy a inner join cc_condition_data b on(a.tbname = b.tablename) where  a.empUpdate_user='"& Session("usrname")  &"'and a.empUpdate_user is not null and a.status_update = 'destroy'   and step=2 "
                   sql = sql & conWhere
                 end if

               
                  ' sql = "select a.*,b.name as crdname from log_Approve_destroy a inner join cc_condition_data b on(a.tbname = b.tablename) where  a.empUpdate_user='"& Session("usrname")  &"'and a.empUpdate_user is not null and a.status_update = 'destroy' and year(a.statusdate) = year(getdate()) and step=2" 
                  no =1
                 rs.open sql,conn,0 
	             if not (rs.bof and rs.eof) then 
	 	            do until rs.eof 
                             str = str & " <tr>"
                              str = str & fncRenderRowLogDestroy(rs,no,"เอกสารที่รอการทำลาย") 
                              str = str & "</tr>"
                        no = no + 1
                         rs.movenext
	 	            loop   
                 end if 
	             rs.close 
                 if( len(str) <=  0 ) then    
                    str = str &  "<tr><td colspan=""7"">ไม่พบข้อมูล</td></tr>"   
                 end if  
                  Response.Write("<script>parent.fncResponeAllDocDesCon('" & str & "');</script>")  
             '*****************end case AllDocDesCon

                  Case "penalty"
            '*****************  case penalty   
                 'sql =  "   select sum(Penalty) as sumPenalty , mm from (   "
	             'sql = sql & "       select   "
		         'sql = sql & "         CASE   "
			     'sql = sql & "              WHEN  a.doc_Returns_date is null or datediff(d,a.doc_Returns_date,getdate()) > 0  THEN 210 "
			     'sql = sql & "              ELSE 200 "
		         'sql = sql & "           END  as Penalty  "
		         'sql = sql & "         , convert(varchar(10),month(a.create_date)) as mm    "
		         'sql = sql & "          from log_Approve_destroy a  left join log_destroy_UploadDoc b on(a.tid = b.tid) "
		         'sql = sql & "          where  a.status_update= 'not-destroy'   "
		         'sql = sql & "          and a.Step = 2   and a.create_user='"& Session("usrname")  &"'    "
		         'sql = sql & "          and year(a.create_date) = year(getdate())  " 
                 'sql = sql & "  ) a   "
                 'sql = sql & "  group by  mm   "

                  
                sql =  "     select sum(Penalty) as sumPenalty , mm from ( "
                sql = sql & "       select   a.tid,200 Penalty , convert(varchar(10),month(a.create_date)) as mm , 'ส่งเอกสารให้บัญชีทำลาย' as 'remark'   " 
                sql = sql & "	  from log_Approve_destroy a   "
                sql = sql & "	  where  a.status_update= 'not-destroy'    "
                sql = sql & "	  and a.Step = 2  and a.create_user='"& Session("usrname")  &"'     "
                sql = sql & "	  and year(a.create_date) = year(getdate()) "
                sql = sql & "union all  "
                sql = sql & "	 select b.tid,"
                sql = sql & "	   10   as Penalty   "
                sql = sql & "	 , convert(varchar(10),month(a.create_date)) as mm  "
                sql = sql & "	 , 'บัญชีส่งเอกสารคืน SUP แต่ SUP  ยังไม่รับเข้าระบบ  ' as 'remark'     "
                sql = sql & "	  from log_Approve_destroy a  inner join log_destroy_UploadDoc b on(a.tid = b.tid)  "
                sql = sql & "	  where  a.status_update= 'not-destroy'    "
                sql = sql & "	  and a.Step = 2   and a.create_user='"& Session("usrname")  &"' and b.update_date is not null  "
                sql = sql & "	  and year(a.create_date) = year(getdate()) "
                sql = sql & "	  and ((a.doc_Returns_date is null and datediff(d,b.update_date,getdate()) > 0 ) or datediff(d ,b.update_date,a.doc_Returns_date) > 0)"
               ' sql = sql & " union all  "
               ' sql = sql & "	 select a.tid,"
               ' sql = sql & "	  10   as Penalty   "
               ' sql = sql & "	 , convert(varchar(10),month(a.create_date)) as mm  "
               ' sql = sql & "	 , ' SUP ส่งเอกสารคืน QC ยังไม่ได้รับเข้าระบบ ' as 'remark'     "
               ' sql = sql & "	  from log_Approve_destroy a   "
               ' sql = sql & "	  where  a.status= 'not-destroy'    "
               ' sql = sql & "	  and a.Step = 1   and a.create_user='"& Session("usrname")  &"' "    
               ' sql = sql & "	  and year(a.create_date) = year(getdate())  "
                sql = sql & "  ) a    "
                sql = sql & "  group by  mm   " 
                   
                      no =1
                      str=""
                        x="mac"
                     rs.open sql,conn,0 
	                 if not (rs.bof and rs.eof) then 
	 	                do until rs.eof 
                            str = str & " <tr>" 
                            str = str &  " <td><button type=""button"" class=""btn btn-link btnExpand""  data-rowClass=""trExpand"& no &""" ><span class=""glyphicon glyphicon-plus-sign"" aria-hidden=""true""></span><span  style=""display:none;"" class=""glyphicon glyphicon-minus-sign"" aria-hidden=""true""></span></button></td>" 
                            str =str &  "<td>" & no & "</td>"
                            str =str &  "<td>" & fncMonth(rs("mm"))  & "</td>"
                            str =str &  "<td style=""text-align:center;"">-" & rs("sumPenalty")  & "</td>" 
                            str = str &  " </tr>"  
                            
                            strRow=""
                            strRow = fncRenderRowsDetailPenalty(rs("mm"))
                            str = str & " <tr id=""trExpand" & no & """ style=""display:none;""><td colspan=""4"" style=""background-color:#EFEFEF;"">"& strRow  &"</td></tr>" 
                            no = no + 1
                             rs.movenext
	 	                loop   
                     end if 
	                 rs.close 
                     if( len(str) <=  0 ) then    
                        str = str &  "<tr><td colspan=""7"">ไม่พบข้อมูล</td></tr>"   
                     end if  
                      Response.Write("<script>parent.fncResponePenalty('" & str & "');</script>")  

            '*****************end case penalty

                   Case "penaltyEmp"
            '*****************  case penalty   
                     
                    sql =" select sum(Penalty) as sumPenalty , mm from ( "
                     sql = sql & "	select   50 as Penalty  , convert(varchar(10),month(b.empUpdatedate)) as mm    "
                     sql = sql & "	from log_destroy_UploadDoc a inner join log_Approve_destroy b on(a.tid = b.tid)"
                     sql = sql & "	  where   a.create_user='"& Session("usrname")  &"'  "
                     sql = sql & "	and year(b.empUpdatedate) = year(getdate()) and datediff(d,b.empUpdatedate,a.update_date) > 0   "
                     sql = sql & ") a "
                     sql = sql & " group by Penalty , mm   "

                      
                      no =1
                      str=""
                        x="mac"
                     rs.open sql,conn,0 
	                 if not (rs.bof and rs.eof) then 
	 	                do until rs.eof 
                            str = str & " <tr>" 
                            str = str &  " <td><button type=""button"" class=""btn btn-link btnExpand""  data-rowClass=""trExpand"& no &""" ><span class=""glyphicon glyphicon-plus-sign"" aria-hidden=""true""></span><span  style=""display:none;"" class=""glyphicon glyphicon-minus-sign"" aria-hidden=""true""></span></button></td>" 
                            str =str &  "<td>" & no & "</td>"
                            str =str &  "<td>" & fncMonth(rs("mm"))  & "</td>"
                            str =str &  "<td style=""text-align:center;"">-" & rs("sumPenalty")  & "</td>" 
                            str = str &  " </tr>"  
                            
                            strRow=""
                            strRow = fncRenderRowsDetailPenaltyEmp(rs("mm"))
                            str = str & " <tr id=""trExpand" & no & """ style=""display:none;""><td colspan=""4"" style=""background-color:#EFEFEF;"">"& strRow  &"</td></tr>" 
                            no = no + 1
                             rs.movenext
	 	                loop   
                     end if 
	                 rs.close 
                     if( len(str) <=  0 ) then    
                        str = str &  "<tr><td colspan=""7"">ไม่พบข้อมูล</td></tr>"   
                     end if  
                      Response.Write("<script>parent.fncResponePenalty('" & str & "');</script>")  

            '*****************end case penalty
             
              Case "excelPenaltyEmp"
            '*****************start case upload document 
                usrname =Session("usrname")  
                sql = " select    a.tid,month(b.empUpdatedate) as mm,b.name,  "
                sql = sql &" convert(varchar(30),b.empUpdatedate,103)  as create_date ,cc.name as  crdName"
                sql = sql &" from log_destroy_UploadDoc  a  inner join log_Approve_destroy b"
                sql = sql &" on(a.tid = b.tid) inner join cc_condition_data cc on(cc.tablename = b.tbname) "
                sql = sql &" where  a.create_user='"&usrname&"' "
                sql = sql &"  and year(b.empUpdatedate) = year(getdate())  and datediff(d,b.empUpdatedate,a.update_date) > 0   order by mm asc"
                  
                no = 1 
                rs.open sql,conn,0
                strRows = ""
                if not(rs.bof and rs.eof) then 
                    strTh ="<thead><th>no.</th> <th>tid</th> <th>month</th> <th>name</th> <th>create date</th> <th>Credit Name</th><th>Penalty</th></thead>"  
                    rows=""
                      do until rs.eof 
                          rows = rows & "<tr> <td>" & no & "</td> <td>" &rs("tid")& "</td> <td>" &fncMonth(rs("mm"))& "</td> <td>" &rs("name")& "</td> <td>" &rs("create_date")& "</td> <td>" &rs("crdName")& "</td><td>-50</td> </tr>"
                          no = no + 1
                          rs.movenext
                      loop 
                    strRows = strTh & "<tbody>" & rows & "</tbody>" 

                end if 
                Response.Write("<script>parent.fnbcExportExcel(""" &strRows&""");</script>")
            '******************end case upload document


          Case "excelPenalty"
            '*****************start case upload document 
             usrname =Session("usrname")    
              ' sql =" select  des.tid, "
              ' sql =sql &" case  "
	          ' sql =sql &"     when  des.doc_Returns_date is null or datediff(d,des.doc_Returns_date,getdate()) > 0 then "
		      ' sql =sql &"         210 "
	          ' sql =sql &"     else "
		      ' sql =sql &"         200 "
	          ' sql =sql &"     end as Penalty   "
              ' sql =sql &"  ,month(des.create_date) as mm,des.name,convert(varchar(30),des.create_date,103) as create_date ,cc.name crdName   "
              ' sql =sql &" from log_Approve_destroy des inner join cc_condition_data cc on(cc.tablename = des.tbname)   "
              ' sql =sql &" left join log_destroy_UploadDoc b on(des.tid =b.tid ) "
              ' sql =sql &" where  des.status_update= 'not-destroy'  and des.create_user='" & usrname & "'   and des.Step = 2     "
              ' sql =sql &" and year(des.create_date) = year(getdate()) order by mm asc   "

          sql ="  select * from (  select  a.name, a.tid,200 Penalty   ,month(a.create_date) as mm ,convert(varchar(30),a.create_date,103) as create_date , 'ส่งเอกสารให้บัญชีทำลาย' as 'remark'    "
	      sql =sql &"    from log_Approve_destroy a   "
	      sql =sql &"    where  a.status_update= 'not-destroy'    "
	      sql =sql &"    and a.Step = 2  and a.create_user='" & usrname & "'     "
	      sql =sql &"    and year(a.create_date) = year(getdate())  "
          sql =sql &"  union all  "
	      sql =sql &"   select a.name, b.tid,"
	      sql =sql &"     10   as Penalty ,month(a.create_date) as mm"  
	      sql =sql &"    ,convert(varchar(30),a.create_date,103) as create_date   "
	      sql =sql &"   , 'บัญชีส่งเอกสารคืน SUP แต่ SUP  ยังไม่รับเข้าระบบ ' as 'remark'     "
	      sql =sql &"    from log_Approve_destroy a  inner join log_destroy_UploadDoc b on(a.tid = b.tid)  "
	      sql =sql &"    where  a.status_update= 'not-destroy'    "
	      sql =sql &"    and a.Step = 2   and a.create_user='" & usrname & "'  and b.update_date is not null  "
	      sql =sql &"    and year(a.create_date) = year(getdate()) "
	      sql =sql &"    and ((a.doc_Returns_date is null  and datediff(d,b.update_date,getdate()) > 0  ) or datediff(d,b.update_date,a.doc_Returns_date) > 0)" 
         'sql =sql &"  union all  "
	     'sql =sql &"   select a.name, a.tid,"
	     'sql =sql &"    10   as Penalty  ,month(a.create_date) as mm "
	     'sql =sql &"    ,convert(varchar(30),a.create_date,103) as create_date  "
	     'sql =sql &"   , ' SUP ส่งเอกสารคืน QC ยังไม่ได้รับเข้าระบบ ' as 'remark'     "
	     'sql =sql &"    from log_Approve_destroy a   "
	     'sql =sql &"    where  a.status= 'not-destroy'    "
	     'sql =sql &"    and a.Step = 1   and a.create_user='" & usrname & "'     "
	     'sql =sql &"    and year(a.create_date) = year(getdate())  "
          sql = sql & " ) a order by mm asc "
           


                no = 1 
                rs.open sql,conn,0
                strRows = ""
                if not(rs.bof and rs.eof) then 
                    strTh ="<thead><th>no.</th> <th>tid</th> <th>month</th> <th>name</th> <th>create date</th> <th>Penalty</th><th> remark</th></thead>"  
                    rows=""
                      do until rs.eof 
                          rows = rows & "<tr> <td>" & no & "</td> <td>" &rs("tid")& "</td> <td>" &fncMonth(rs("mm"))& "</td> <td>" &rs("name")& "</td> <td>" &rs("create_date")& "</td> <td>-" & rs("Penalty") &"</td><td>"  &rs("remark")& "</td> </tr>"
                          no = no + 1
                          rs.movenext
                      loop 
                    strRows = strTh & "<tbody>" & rows & "</tbody>" 

                end if 
                Response.Write("<script>parent.fnbcExportExcel(""" &strRows&""");</script>")
            '******************end case upload document

          Case "AllEmpDesCon"
             '*****************  case AllEmpDesCon 
                   sql = "select a.* ,a.statusdate t_editdate,a.statusdate applicationdate, b.name crdname  from dbo.log_Approve_destroy  a inner join cc_condition_data b "
                   sql = sql & "on (a.tbname = b.tablename)"
                   sql = sql &"where a.status not like '%destroy%' and a.status_update is null and a.step = 1 "
                  ' sql = sql &"and year(a.statusdate) = year(getdate())"

                  no =1
                 rs.open sql,conn,0 
	             if not (rs.bof and rs.eof) then 
	 	            do until rs.eof 
                             str = str & " <tr>"
                              'str = str &  fncRenderRowEmpDestroy(rs,no)   'fncRenderRowLogDestroy(rs,no,"เอกสารที่รอการทำลาย") 
                                str = str & fncRenderRowAllDestroy(rs,no,rs("status")) 
                              str = str & "</tr>"
                        no = no + 1
                         rs.movenext
	 	            loop   
                 end if 
	             rs.close 
                 if( len(str) <=  0 ) then    
                    str = str &  "<tr><td colspan=""7"">ไม่พบข้อมูล</td></tr>"   
                 end if  
                Response.Write("<script>parent.fncResponeAllEmpDes('" & str & "');</script>")  

         Case "NotDestroyEmp"
             '*****************  case AllEmpDesCon 
                   sql = " select a.* ,a.statusdate t_editdate,a.statusdate applicationdate, b.name crdname  from dbo.log_Approve_destroy  a inner join cc_condition_data b "
                   sql = sql & " on (a.tbname = b.tablename)"
                   sql = sql &" where a.empUpdate_user='"& Session("usrname")  &"' and  a.status_update  like '%not-destroy%'  and a.empUpdate_user is not null  and a.step = 2 "
                  ' sql = sql &" and year(a.statusdate) = year(getdate())"
                   sql = sql & " and a.tid not in (select tid from log_destroy_UploadDoc )" 
                  no =1
                 rs.open sql,conn,0 
	             if not (rs.bof and rs.eof) then 
	 	            do until rs.eof 
                             str = str & " <tr>"
                                 str = str & fncRenderRowAllDestroy(rs,no,rs("status")) 
                              str = str & "</tr>"
                        no = no + 1
                         rs.movenext
	 	            loop   
                 end if 
	             rs.close 
                 if( len(str) <=  0 ) then    
                    str = str &  "<tr><td colspan=""7"">ไม่พบข้อมูล</td></tr>"   
                 end if  
                Response.Write("<script>parent.fncResponeNotDestroyEmp('" & str & "');</script>")  


             '*****************end case AllEmpDesCon
          Case "empDestroy"
           '*****************  case empDestroy
                   sql = "select top 1 * from log_destroy where tid = "& tid
                    tbname = Request("tbname")
                    status =Request("status")
                    name = Request("name")
 			        rslog.open sql,conn,0
 			        if (rslog.bof and rslog.eof) then
 				        rslog.close  
 				        sql = "select top 1 * from log_destroy order by id desc"
 				        rslog.open sql,conn,1,2
 				        if not(rslog.bof and rslog.eof) then
 					        id = rslog("id") + 1
 				        else
 					        id = 1
 				        end if
 				        rslog.addnew
 				        rslog("id") = id
 				        rslog("tid") = tid
 				        rslog("name") = name
 				        rslog("tbname") =tbname
 				        rslog("status") = status
 				        rslog("applydate") = cdate(apply)
 				        rslog("statusdate") = cdate(statusdate)
 				        rslog("logdate") = now
 				        rslog.update
 				        rslog.close
 			        else
 				        rslog.close
 			        end if    
                    conn.execute " update log_Approve_destroy set status_update = 'Destroy',empUpdatedate = getdate(),empUpdate_user='"& Session("usrname") &"'   where  tid = "& tid 
 			        conn.execute " update "&tbname&" set retdoc = 'Destroy'  where (retdoc is null or retdoc = '') and tid = "& tid
        
                    Response.Write("<script>parent.fncCallDocDestro();</script>")
           '*****************end case empDestroy
            Case "login"
                    '*****************  case login
                    sql = "select * from user_login where username = '"& Request("usr_login") &"' and password = '" & Request("pass_login") & "' and system_name = 'doc_destroy' and department = 'doc'" 
 			        rslog.open sql,conn,0 
                    if (rslog.bof and rslog.eof) then
  			            rslog.close
                      Response.Write("<script>parent.fncLoginFail();</script>")

                    else 
 				        Session("usrname") =  rslog("USERNAME")
                        Session("typeEmp") =  rslog("USER_ROLE")
                        Session("name") = rslog("NAME")  & " " &  rslog("LNAME")  
                        url = ""  
                        if rslog("USER_ROLE") ="sup"  then
                           url = "main.asp"
                        elseif  rslog("USER_ROLE") ="admin" then
                             url = "manage_user.asp"
                        else
                           url = "emp.asp"
                        end if
                        Response.Write("<script>parent.fncLoginSuccess('"& url &"');</script>")
                         rslog.close
 			        end if   
                  '*****************end case login
            Case "supApprove"
                        '*****************  case supApprove
                        sql = "select top 1 * from log_Approve_destroy where tid = "& tid
 			            rslog.open sql,conn,0
 			            if (rslog.bof and rslog.eof) then
 				            rslog.close  
 				            sql = "select top 1 * from log_Approve_destroy order by id desc"
 				            rslog.open sql,conn,1,2
 				            if not(rslog.bof and rslog.eof) then
 					            id = rslog("id") + 1
 				            else
 					            id = 1
 				            end if
 				            rslog.addnew
 				            rslog("id") = id
 				            rslog("tid") = tid
 				            rslog("name") = name
 				            rslog("tbname") = tbname
 				            rslog("status") = status
 				            rslog("applydate") = cdate(apply)
 				            rslog("statusdate") = cdate(statusdate)
 				            rslog("dateApprove") = now
                            rslog("userApprove") = Session("usrname") 
 				            rslog.update
 				            rslog.close
 			            else
 				            rslog.close
 			            end if 
                     '*****************end case supApprove
            Case "loadUserLogin"
              '*****************  case loadUserLogin
                        sql = "select * from user_login where system_name = 'doc_destroy' and USER_ROLE <> 'admin' and department = 'doc'" 
 			            rs.open sql,conn,0 
                         str =""
                         no=1
                        if not (rs.bof and rs.eof) then  
	 	                    do until rs.eof 
                              name = rs("NAME")  
                              lastname = rs("LNAME") 
                              email = rs("EMAIL")
                              userRole = rs("USER_ROLE")
                              username =  rs("USERNAME")
                              password = rs("PASSWORD") 
                              id = rs("id")
                              str = str & " <tr>" 
                               str = str & "<td>"& no &"</td>"
                              str = str & "<td>"& username&"</td>"
                              str = str & "<td>"& password&"</td>"
                              str = str & "<td>"& name & " " & lastname &"</td>"
                              str = str & "<td>"& email  &"</td>"
                              str = str & "<td>"& userRole &"</td>"
                              str = str & "<td>"
                              str = str & " <button type=""button""  onclick=""fncEdit(this);""  data-name=""" & name & """ data-lname="""&lastname&""" data-email="""&email&"""  data-userRole=""" & userRole&""" data-userName="""& username &""" data-password="""& password & """ data-id="""&id&""" class=""btn btn-link btnEdit"" aria-label=""Left Align"">"
                               str = str & " <span class=""glyphicon glyphicon-edit"" aria-hidden=""true""> edit</span>"
                              str = str & " </button>"

                              str = str & " <button type=""button"" onclick=""fncDelete(this);""   data-name=""" & name & """ data-lname="""&lastname&""" data-email="""&email&"""  data-userRole=""" & userRole&""" data-userName="""& username &""" data-password="""& password & """ data-id="""&id&"""    class=""btn btn-link btnDelete"" aria-label=""Left Align"">"
                              str = str & " <span class=""glyphicon  glyphicon-trash"" aria-hidden=""true""> delete</span>"
                              str = str & " </button>"
                   
                              str = str & "</td> "
                              str = str & "</tr>"
                  
                                no = no + 1
                                 rs.movenext
	 	                    loop   
                         end if 
                         rs.close
                        Response.Write("<script>parent.fncRenderTbody('"& str &"');</script>")
                '*****************end case loadUserLogin
            Case "EditUsrLogin"
             '*****************  case EditUsrLogin
                 name=Request("Name")
                lname= Request("Lname")
                username= Request("UserName")
                password = Request("Password")
                UsrRole  =Request("UsrRole") 
                email = Request("Email")
                id  =Request("id")   
          	    sql = "select top 1 * from user_login where id = " & id
 			    rs.open sql,conn,1,2  
 			    rs("NAME") =  name
 			    rs("LNAME") = lname
 			    rs("USERNAME") =  username
 			    rs("PASSWORD") =password
                rs("USER_ROLE") = UsrRole
                rs("Email") = email
 			    rs.update
 			    rs.close   
               Response.Write("<script>parent.fncEditSuccess();</script>")
              '*****************end case EditUsrLogin
            Case "AddUserLogin"
            '*****************end case AddUserLogin
                sql = "select top 1 * from user_login order by id desc"
 				rs.open sql,conn,1,2 
 				rs.addnew 
 				rs("NAME") =  Request("Name")
 				rs("LNAME") =  Request("Lname")
 				rs("USERNAME") =  Request("UserName")
 				rs("PASSWORD") = Request("Password") 
                rs("USER_ROLE") = Request("UsrRole") 
 				rs("CREATE_USER") = Session("usrname")
 				rs("CREATE_DATE") = now
                rs("SYSTEM_NAME") = "doc_destroy"
                rs("DEPARTMENT") = "doc"  
 				rs.update
 				rs.close 
                Response.Write("<script>parent.fncAddSuccess();</script>")
            '*****************end case AddUserLogin
          Case "deleteUserLogin"
                id  =Request("id") 
                sql =" delete from  user_login   where  id = "& id  
                conn.execute sql 
               Response.Write("<script>parent.fncDeleteSuccess();</script>")

          Case "UploadDoc"
            '*****************start case upload document  
             folderName = Year(now()) & "" & Month(now())& "" & Day(now()) & "" & HOUR(now()) & "" & MINUTE(now())& "" & SECOND(now())
                For Each Item In Request.Form
                fieldName = Item 
                     IF fieldName = "TabType" OR fieldName = "hdfFilename" THEN
                           x=1 ''data not save
                     else 
                        fieldValue = Request.Form(Item)    
                        sql = "select top 1 * from log_destroy_UploadDoc"
 				        rs.open sql,conn,1,2 
 				        rs.addnew 
 				        rs("TID") =  fieldValue
 				        rs("Upload_Date") =  now()
 				        rs("Upload_User") = Session("usrname") 
                        rs("File_Name") = Request("filename")
                        rs("Folder_Name") = folderName
 				        rs.update
 				        rs.close   
                    END IF
                Next 
                Response.Write("<script>parent.fncUploadFile('" & folderName &"');</script>")
            '******************end case upload document
             Case "UpdateWaitSup"
            '*****************start case upload document  
             folderName = Year(now()) & "" & Month(now())& "" & Day(now()) & "" & HOUR(now()) & "" & MINUTE(now())& "" & SECOND(now())
                For Each Item In Request.Form
                fieldName = Item 
                     IF fieldName = "TabType" OR fieldName = "hdfFilename" THEN
                           x=1 ''data not save
                     else 
                        fieldValue = Request.Form(Item)    
                        sql = "select top 1 * from log_destroy_UploadDoc"
 				        rs.open sql,conn,1,2 
 				        rs.addnew 
 				        rs("TID") =  fieldValue
 				        rs("Create_Date") =  now()
 				        rs("Create_User") = Session("usrname")  
                        rs("Folder_Name") = folderName
 				        rs.update
 				        rs.close   
                    END IF
                Next 
                Response.Write("<script>parent.fncUploadFile('" & folderName &"');</script>")
            '******************end case upload document

          Case "UpdateDataUpload"
            '*****************start case upload document 
                usrname =Session("usrname") 
                sql = "select top 1 Folder_Name  from  log_destroy_UploadDoc where create_user = '" &  usrname&"' order by create_date desc  " 
                rs.open sql,conn,0
                if not(rs.bof and rs.eof) then
                   foldername = rs("Folder_Name")    
                   rs.Close
                     sql =" update log_destroy_UploadDoc set Update_Date=getdate() , Update_User='"&usrname&"' , File_Name='" & Request("filename") & "'  where create_user = '" &  usrname&"' and Folder_Name = '"&foldername&"'"
                      conn.execute sql 
                end if 
                Response.Write("<script>parent.fncUploadFile('" & folderName &"');</script>")
            '******************end case upload document

                         
          Case "ShowUploadDoc"
            '*****************Show list Upload Doc
                usrname =Session("usrname") 
                sql = "  select distinct folder_name, convert(varchar(50),update_date,103) update_date ,"
                sql = sql & "('/destroying/doc/excel/'+folder_name+'/'+file_name) as path   from dbo.log_destroy_UploadDoc "
                sql = sql & "where update_user = '" & usrname &"' "  
                 no = 1 
                rs.open sql,conn,0
                rows = "" 
                if not(rs.bof and rs.eof) then    
                      do until rs.eof 
                       urlGetDoc="/destroying/doc/getDocument.asp?folderName=" & rs("folder_name")
                       urlSaveDoc="/destroying/doc/SaveDocument.asp?folderName=" & rs("folder_name")
                          rows = rows & "<tr> <td>" & no & "</td> <td>" &rs("update_date")& "</td> <td><a class=""aPreview"" data-href="""& urlGetDoc &""" target=""_blank"" ><i class=""glyphicon glyphicon-search""></i></a></td> <td>  <a class=""aDowload""   href=""#""  data-href="""& urlSaveDoc &"""><i class=""glyphicon glyphicon-floppy-save""></i></a></td></tr>"
                          no = no + 1
                          rs.movenext
                      loop   
                end if 
                 
               Response.Write("<script>parent.fncLoadShowListEmpUpload('" & rows &"');</script>")
            '******************end case Show list Upload Doc

           Case "tmpScan" 
                sql = "select top 1 * from TMP_DESTROY_SCAN where tid = "& tid  
        	    rslog.open sql,conn,0
        	    if (rslog.bof and rslog.eof) then
        		    rslog.close  
        		    sql = "select top 1 * from TMP_DESTROY_SCAN  order by id desc"
        		    rslog.open sql,conn,1,2 
        		    rslog.addnew 
        		    rslog("TID") = tid
        		    rslog("CREATE_DATE") = now 
        		    rslog.update
        		    rslog.close
        	    else    
                       conn.execute "update TMP_DESTROY_SCAN set UPDATE_DATE = getdate() where  tid = "& tid   
        	         rslog.close
        	    end if 
               Response.Write("<script>parent.RenderResult('" & tid &"');</script>")




          Case else
            response.write("out of query ja !")
    End Select
         
          
%>
</body>
</html>
<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->