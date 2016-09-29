<%
 
  
Dim objFSO, objFile, objFolder
'20157793235

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
folderName = Request("folderName")
path = "/destroying/doc/excel/"&folderName&"/" 
VisualPath = Server.MapPath("/destroying/doc/excel/" & folderName & "/")
Set objFolder = objFSO.GetFolder(VisualPath)

For Each objFile in objFolder.Files  
    Response.Write("<div><b>"& objFile.Name &"</b></div><div><iframe frameBorder=""0"" width=""600px""   height=""600px""  src=""" &  path &  objFile.Name & """></iframe></div>") 

Next
Set objFolder = Nothing
Set objFSO = Nothing


 %>