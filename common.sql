

/********** script gen view  VW_DestroyAll ************************/ 
select  ' (select   name+''   ''+lastname as name,tid,applicationdate  
,incomplete,appreturn,appreturn_msg,t_status,t_editdate,activeby,retdoc ,qcby
 ,'''+tablename+''' as tbname, '''+name+'''  as crdname from '+tablename+' where datediff(m,applicationdate,getdate()) between 1 and 3   AND (appreturn_msg IS NOT NULL)) union '   from cc_condition_data  
 where inbound_status = 'on' and tablename not like '%mg_%'  
and tablename not like '%car_%' and tablename not like '%citi%'  
order by bank asc
/********************************************************/  

/**** VW_DestroyList *******/
 
select  *,'INCOMPLETE' AS statusTb from  VW_DestroyAll  where retdoc is null and t_status is null 
	     and incomplete is not null and appreturn is null 
	     and datediff(d,convert(varchar,year(dateadd(m,-3,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-3,getdate()))),2)+'-01',incomplete) >= 0 
	     and datediff(d,getdate() - 60,incomplete) < 0 
	     and datediff(d,'2010-11-04',incomplete) >= 0 
          and (datediff(m,applicationdate,getdate()) between 0 and 3)  and tid not  in (select tid from log_approve_destroy)  
	 

UNION
select  *,'REJECT' AS statusTb  from VW_DestroyAll   where retdoc is null and t_status = 'r' 
	     and appreturn_msg is not null
	     and t_editdate is not null and datediff(d,convert(varchar,year(dateadd(m,-1,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-1,getdate()))),2)+'-01',t_editdate) >= 0 
	     and datediff(d,getdate() - 7,t_editdate) < 0 
	     and year(t_editdate) >= 2011 
         and (datediff(m,applicationdate,getdate()) between 0 and 3)  and tid not  in (select tid from log_approve_destroy) 
	    

UNION
select   *  ,   'CANCEL' AS statusTb from  VW_DestroyAll  where retdoc is null and t_status = 'c' 
	     and appreturn_msg is not null 
	     and t_editdate is not null and datediff(d,convert(varchar,year(dateadd(m,-1,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-1,getdate()))),2)+'-01',t_editdate) >= 0 
	     and datediff(d,getdate() - 7,t_editdate) < 0 
	     and year(t_editdate) >= 2011 
         and (datediff(m,applicationdate,getdate()) between 0 and 3) and tid not  in (select tid from log_approve_destroy) 
	     
select  *,'INCOMPLETE' AS statusTb from  VW_DestroyAll  where retdoc is null and t_status is null 
	     and incomplete is not null and appreturn is null 
	     and datediff(d,convert(varchar,year(dateadd(m,-3,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-3,getdate()))),2)+'-01',incomplete) >= 0 
	     and datediff(d,getdate() - 60,incomplete) < 0 
	     and datediff(d,'2010-11-04',incomplete) >= 0 
          and (datediff(m,applicationdate,getdate()) between 0 and 3)  and tid not  in (select tid from log_approve_destroy)  
	 

UNION
select  *,'REJECT' AS statusTb  from VW_DestroyAll   where retdoc is null and t_status = 'r' 
	     and appreturn_msg is not null
	     and t_editdate is not null and datediff(d,convert(varchar,year(dateadd(m,-1,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-1,getdate()))),2)+'-01',t_editdate) >= 0 
	     and datediff(d,getdate() - 7,t_editdate) < 0 
	     and year(t_editdate) >= 2011 
         and (datediff(m,applicationdate,getdate()) between 0 and 3)  and tid not  in (select tid from log_approve_destroy) 
	    

UNION
select   *  ,   'CANCEL' AS statusTb from  VW_DestroyAll  where retdoc is null and t_status = 'c' 
	     and appreturn_msg is not null 
	     and t_editdate is not null and datediff(d,convert(varchar,year(dateadd(m,-1,getdate())))+'-'+right('0'+convert(varchar,month(dateadd(m,-1,getdate()))),2)+'-01',t_editdate) >= 0 
	     and datediff(d,getdate() - 7,t_editdate) < 0 
	     and year(t_editdate) >= 2011 
         and (datediff(m,applicationdate,getdate()) between 0 and 3) and tid not  in (select tid from log_approve_destroy) 
	     



--
--update log_Approve_destroy set status_update=null,empUpdatedate=null,empUpdate_user=null where tid= 3264429 
--update  pl_bay_type1
--set retdoc = null
--  where tid= 3264429
--delete from log_destroy where tid= 3264429 
 

--
--select status_update,empUpdatedate,empUpdate_user ,* from log_Approve_destroy  where tid= 3264429  
--select retdoc,* from  pl_bay_type1  where tid= 3264429 
--select   * from log_destroy where tid= 3264429 











-- ค่าปรับ emp
 select sum(Penalty) as sumPenalty , mm from (  
	select   50 as Penalty  , convert(varchar(10),month(b.empUpdatedate)) as mm ,b.empUpdatedate,a.update_date
	from log_destroy_UploadDoc a inner join log_Approve_destroy b on(a.tid = b.tid) 
	  where   a.create_user='emp'  
	and year(b.empUpdatedate) = year(getdate()) and datediff(d,b.empUpdatedate,a.update_date) > 0    
) a  
 group by Penalty , mm    
  

--  ค่าปรับ  sup
select sum(Penalty) as sumPenalty , mm from (  
       select   a.tid,200 Penalty , convert(varchar(10),month(a.create_date)) as mm , 'ส่งเอกสารให้บัญชีทำลาย' as 'remark'     
	  from log_Approve_destroy a   
	  where  a.status_update= 'not-destroy'     
	  and a.Step = 2  and a.create_user='emp'      
	  and year(a.create_date) = year(getdate())  
union all  
	 select b.tid, 
	   10   as Penalty   
	 , convert(varchar(10),month(a.create_date)) as mm  
	 , 'บัญชีส่งเอกสารคืน SUP แต่ SUP  ยังไม่รับเข้าระบบ  ' as 'remark'     
	  from log_Approve_destroy a  inner join log_destroy_UploadDoc b on(a.tid = b.tid)   
	  where  a.status_update= 'not-destroy'    
	  and a.Step = 2   and a.create_user='emp' and b.update_date is not null   
	  and year(a.create_date) = year(getdate())  
	  and ((a.doc_Returns_date is null and datediff(d,b.update_date,getdate()) > 0 ) or datediff(d ,b.update_date,a.doc_Returns_date) > 0) 
  ) a     
  group by  mm     




 














/*******************/
/** ค่าปรับ ***********/ 
-- emp
 select sum(Penalty) as sumPenalty , mm ,empUpdate_user,(b.name+' '+b.lname) fullName from (  
	select   50 as Penalty  , convert(varchar(10),month(b.empUpdatedate)) as mm ,b.empUpdate_user
	from log_destroy_UploadDoc a inner join log_Approve_destroy b on(a.tid = b.tid) 
     --  inner join user_login c on(b.empUpdate_user = c.username)
	where  year(b.empUpdatedate) = year(getdate()) and month(b.empUpdatedate) = month(getdate()) 
	and datediff(d,b.empUpdatedate,a.update_date) > 0    
) a  inner join user_login b on(a.empUpdate_user = b.username)
 group by Penalty , mm,empUpdate_user ,name,lname
 

--  sup
select sum(Penalty) as sumPenalty , mm ,remark,a.create_user,(b.name+' '+b.lname) fullName from (  
       select   a.tid,200 Penalty , convert(varchar(10),month(a.create_date)) as mm , 'ส่งเอกสารให้บัญชีทำลาย' as 'remark' 
      ,a.create_user    
	  from log_Approve_destroy a   
	  where  a.status_update= 'not-destroy'     
	  and a.Step = 2      
	  and year(a.create_date) = year(getdate())   and month(a.create_date) = month(getdate())  
union all  
	 select b.tid, 
	   10   as Penalty   
	 , convert(varchar(10),month(a.create_date)) as mm  
	 , 'บัญชีส่งเอกสารคืน SUP แต่ SUP  ยังไม่รับเข้าระบบ  ' as 'remark'     
	,a.create_user
	  from log_Approve_destroy a  inner join log_destroy_UploadDoc b on(a.tid = b.tid)   
	  where  a.status_update= 'not-destroy'    
	  and a.Step = 2    and b.update_date is not null   
	  and year(a.create_date) = year(getdate())  and month(a.create_date) = month(getdate())  
	  and ((a.doc_Returns_date is null and datediff(d,b.update_date,getdate()) > 0 ) or datediff(d ,b.update_date,a.doc_Returns_date) > 0) 

  ) a   inner join user_login b on(a.create_user = b.username)
  group by  mm  ,remark ,a.create_user,name,lname

























 