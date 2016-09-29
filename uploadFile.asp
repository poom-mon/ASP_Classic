 
<%@  language="VBScript" %>      
<!--#include file="aspupload.asp" -->
<%      
 
  For Each Item In Request.Form
    fieldName = Item
    fieldValue = Request.Form(Item) 
         Response.Write(""& fieldName &" : "& fieldValue)       
    Next 

    'Response.Write(Request("filename"))

'folder  =  Year(now()) & "-" & Month(now())& "-" & Day(now()) & "_" & HOUR(now()) & "-" & MINUTE(now())& "-" & SECOND(now())
folder = Request("filename")
uploadsDirVar = Server.MapPath("/destroying/doc/excel/") 
Dim ArrFile(12, 2),tid 

call Uploader()
sub Uploader
		Dim Upload, fileName, fileSize, ks, i, fileKey
				
		Set Upload = New ASPUpload
		call Upload.GetData() 
		call Upload.Save(uploadsDirVar, folder) 
		If Err.Number<>0 then Exit sub

	    ks = Upload.UploadedFiles.keys

	    if (UBound(ks) <> -1) then
			i = 1
	        for each fileKey in Upload.UploadedFiles.keys
				ArrFile(i, 1) = Upload.UploadedFiles(fileKey).FileName
				ArrFile(i, 2) = Upload.UploadedFiles(fileKey).Length
				i = i + 1
	        next
	    end if

end sub
%>