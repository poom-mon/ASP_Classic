
<%
 
  
Dim objFSO, objFile, objFolder
'20157793235

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
folderName = Request("folderName")
path = "/destroying/doc/excel/"&folderName&"/" 
VisualPath = Server.MapPath("/destroying/doc/excel/" & folderName & "/")
Set objFolder = objFSO.GetFolder(VisualPath)

For Each objFile in objFolder.Files 
    Response.Write("<div><b>"& objFile.Name &"</b><br><input name=""btnSave"" data-urlDoc=""" &  path &  objFile.Name &""" type=""button"" value=""บันทึกเอกสาร""  /> </div><div><iframe frameBorder=""0"" width=""600px""   height=""600px""  src=""" &  path &  objFile.Name & """></iframe></div>") 

Next 

Set objFolder = Nothing
Set objFSO = Nothing 
 %>
<script src="js/jquery-1.11.3.min.js" type="text/javascript"></script>
<script type="text/javascript">
    $("input[name=btnSave]").on("click", function () { 
        saveFile($(this).attr("data-urlDoc"));
    });
    function saveFile(url) {
        var filename = url.substring(url.lastIndexOf("/") + 1).split("?")[0];
        var xhr = new XMLHttpRequest();
        xhr.responseType = 'blob';
        xhr.onload = function () {
            var a = document.createElement('a');
            a.href = window.URL.createObjectURL(xhr.response);
            a.download = filename;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            delete a;
        };
        xhr.open('GET', url);
        xhr.send();
    }
</script>
