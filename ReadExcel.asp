<!--#include virtual="/silkspan_ssl/inc/conn.asp"-->
<%
conn.CommandTimeout = 1000
server.scripttimeout = 500
%>
<html>
<head>
    <title>DESTROYING DOCUMENTS REPORT</title>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-874">
</head>
<body>
    <form name="f" enctype="multipart/form-data">
    <!--include virtual="/destroying/doc/html/head.html" -->
    <input type="hidden" name="tid" id="hdfTid">
    <input type="hidden" name="TabType" id="hdfTabType">
    <input type="hidden" name="objhtml" id="hdfObjhtml">
   <br />
   <div class="well">
        <div class="row">
      <div class="col-xs-6 col-md-4"> อัพโหลดไฟล์ excel ที่ใช้</div>
          <div class="col-xs-6 col-md-4"> <input type="FILE" id="fileInputCSV" size="40" name="fUload"><br></div>
          <div class="col-xs-6 col-md-4"> </div>
        </div>

        <div class="row">
      <div class="col-xs-6 col-md-4">เอกสารใบประหน้า</div>
          <div class="col-xs-6 col-md-4"> <input type="FILE"   size="40" name="fupload2"><br></div>
          <div class="col-xs-6 col-md-4"> </div>
        </div>

        <div class="row">
      <div class="col-xs-6 col-md-4"></div>
          <div class="col-xs-6 col-md-4"><input class="btn btn-success" type="button" value="บันทึกข้อมูลเข้าสู่ระบบ" id="btnSave" /></div>
          <div class="col-xs-6 col-md-4"> </div>
        </div>

    </div> 
    <table style="display: none;" id="tbResult">
    </table>
    <!--include virtual="/destroying/doc/html/footer.html" -->
    </form>
    <link href="css/bootstrap.css" rel="stylesheet" type="text/css" />
    <script src="js/jquery-1.11.3.min.js" type="text/javascript"></script>
    <script src="js/simple-excel.js" type="text/javascript"></script>
    <script type="text/javascript">

        var fileInputCSV = document.getElementById('fileInputCSV');
        fileInputCSV.addEventListener('change', function (e) {
            var file = e.target.files[0];
            var csvParser = new SimpleExcel.Parser.CSV();
            csvParser.setDelimiter(',');
            csvParser.loadFile(file, function () {
                var sheet = csvParser.getSheet();
                var table = document.getElementById('tbResult');
                table.innerHTML = "";
                sheet.forEach(function (el, i) {
                    var row = document.createElement('tr');
                    el.forEach(function (el, i) {
                        var cell = document.createElement('td');
                        cell.innerHTML = el.value;
                        row.appendChild(cell);
                    });
                    table.appendChild(row);
                });
            });
        });
        $("#btnSave").on("click", function () {
            var str = "";
            $('#tbResult tbody tr td:nth-child(2)').each(function () {
                var tid = $(this).text();
                console.log(tid);
                str = str + "<input type='hidden' name='hdf" + tid + "' value='" + tid + "'> ";
            });
            $("#dvDataform").html(str);
            fncUpdateData();
        });
        function fncUploadFile(folodername) {
            document.f.target = "iF_Status"
            document.f.method = "post"
            document.f.action = "uploadFile.asp?filename=" + folodername;
            document.f.submit();
        }
        function fncUpdateData() {
            var str = document.getElementById("fileInputCSV").value;
            var pieces = str.split('\\');
            var filename = pieces[pieces.length - 1]

           
            document.Fm.target = "iF_Status"
            document.Fm.method = "post"
            document.Fm.action = "Query.asp?filename=" + filename;
            document.Fm.submit();
        }
        </script>
    <iframe name="iF_Status" width="800" height="100" align="center" frameborder="0">
    </iframe>
    <%set rs = nothing%>
</body>
</html>
<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->
<form name="Fm" method="post">
<input type='hidden' name="TabType" id="hdfTabType" value="UploadDoc">
<!--   <input type='hidden' name="hdfFilename" id="hdfFilename"  >-->
<div style="display: none;" id="dvDataform">
</div>
<!--  <INPUT TYPE="FILE" SIZE="40" NAME="fUload"><BR>  -->
</form>