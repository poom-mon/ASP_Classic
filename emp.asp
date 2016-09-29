<!--#include virtual="/silkspan_ssl/inc/conn.asp"-->
<%
conn.CommandTimeout = 1000
server.scripttimeout = 500
%>
<html>
<head>
    <title>DESTROYING DOCUMENTS REPORT</title>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-874">
    <!--#include virtual="/destroying/doc/html/head.html" -->
    <script language="javascript">
        function num_key(evt) {
            var iKeyCode;
            var IsValid = false;

            if (window.event) { // IE
                iKeyCode = evt.keyCode
            }
            else if (evt.which) { // Netscape/Firefox/Opera
                iKeyCode = evt.which
            }
            //alert(iKeyCode)
            if (iKeyCode == 13) {
                return true;
            }
            else {
                return false;
            }
        }

        setInterval(function () {
          var user= "<%=Session("usrname") %>";
          var url = "/destroying/doc/login.asp?logout=true"
            if(user =="")
              window.location.href = url;

        }, 3000);

        
 var statusBarcode = true;

          function fncGetRowVisible(idTb, idTbody) {
                        var strtb = "";
                        var strTh = "";
                        $("#" + idTb + ">thead>tr>th").each(function () { // ดึง th
                            if ($(this).find("input").html() == "") {  // เช็คประเภท th ดึงชือใน Textbox
                                $(this).find("input").each(function () {
                                    strTh = strTh + "<th>" + $(this).attr("placeholder") + "</th>";
                                });
                            }
                            else { // ดึงธรรมดา
                                strTh = strTh + "<th>" + $(this).html() + "</th>";
                            }
                        });
                        strtb = "<thead>" + strTh + "</thead>";
                        var rows = "";
                        $("#" + idTbody + ">tr:visible").each(function (i, tr) { // ดึง Rows ที่เปิดใช้
                            rows = rows + "<tr>" + $(this).html() + "</tr>";
                        });
                        rows = "<tbody>" + rows + "</tbody>";
                        strtb = strtb + rows
                        return strtb;
            }

        $(function () {
            fncCallDocDestro();
            $("#btnDestroy").on("click", function () {
                $('input[name=chkDetroy]:checked').each(function () {
                    var _tid = $(this).attr("data-tid");
                    var _tbname = $(this).attr("data-tbname");
                    var _status = $(this).attr("data-status");
                    var _name = $(this).attr("data-name");

                    fncDestroy(_tid, _tbname, _status, _name);
                });
            }); 
            function fncSaveExportWaitSup(idTb, idTbody) {
                var str = "";
                $("#" + idTbody + ">tr:visible td:nth-child(2)").each(function (i, tr) { // ดึง Rows ที่เปิดใช้
                    var tid = $(this).html();
                    str = str + "<input type='hidden' name='hdf" + tid + "' value='" + tid + "'> ";
                });
                $("#dvHideTid").html(str)
                document.FmSubUpload.target = "iF_Status"
                document.FmSubUpload.method = "post"
                document.FmSubUpload.action = "Query.asp"
                document.FmSubUpload.submit();
            }

            $("#btnExport").on("click", function () {
                document.getElementById("hdfTabType").value = "excel";
                document.getElementById("hdfObjhtml").value = fncGetRowVisible("tbname", "tbodyAllDocDelete");
                document.f.target = "_blank"
                document.f.method = "post"
                document.f.action = "excel.asp?xlsName=doc_destroyEmp"
                document.f.submit();
            });
            $("#btnExportExceltab2").on("click", function () {
                document.getElementById("hdfTabType").value = "excel";
                document.getElementById("hdfObjhtml").value = fncGetRowVisible("tbTabNotDesEmp", "tbodyNotDesEmp");
                document.f.target = "_blank"
                document.f.method = "post"
                document.f.action = "excel.asp?xlsName=doc_notdestroy_emp"
                document.f.submit();

                fncSaveExportWaitSup("tbTabNotDesEmp", "tbodyNotDesEmp");
            });
            $("#btnExportExcelDes").on("click", function () { 
                $("#hdfExportExcelFilter").val("export")  
                fncCallDocDesCondition()   

            });
            $("#btnUpload").on("click", function () {
                fncUpdateDataUpload();
            });
            var i =1;
            $("#btnAddDocUpload").on("click", function () {
              $("#dvAddRows").append("<div class='alert alert-info' id='dvUp"+i+"' style='width:600px;margin-bottom: 5px;height: 40px;'> <div class='col-xs-8 col-md-4'> <input type='FILE'  size='40' name='fupload" + i + "'></div><div class='col-xs-4 col-md-4'  style='text-align:right;'><i class='glyphicon glyphicon-trash' onclick='delUpload("+i+");'></i></div></div>  ");
              i=i+1;
            });
            $("#btnExportPenalty").on("click", function () {
                //document.getElementById("hdfObjhtml").value = fncGetRowVisible("tbPenalty", "tbodyPenalty");
                fncExportXlsPenatyEmp();
            });

             $("#txt_tid").on("keyup",function(){
              toUpdate($(this).val());
            });

            $("#btnKeyManual").on("click",function(){
                $(this).css("display","none");
                $("#btnScan").css("display","");
                $("#tbIdManual").css("display","");
                 $("#txt_tid").css("display","none");
                 statusBarcode = false;
            });
              $("#btnScan").on("click",function(){
                $(this).css("display","none");
                $("#btnKeyManual").css("display","");
                $("#tbIdManual").css("display","none");
                 $("#txt_tid").css("display","");
                 statusBarcode = true;
            });

        });
        function delUpload(va){
            $("#dvUp"+va).remove();
        }
        function fnbcExportExcel(str) {
            document.getElementById("hdfObjhtml").value=str;
            document.getElementById("hdfTabType").value = "excel";
            document.f.target = "_blank"
            document.f.method = "post"
            document.f.action = "excel.asp?xlsName=doc_penaltyEmp"
            document.f.submit();
        }
        function bindPageing() {
            $("#myPagerTab1,#myPagerTab2,#myPagerDocUpload").html("") ;//,#myPagerNotDesEmp
           // $('#tbodyDocDes').pageMe({ pagerSelector: '#myPagerTab2', showPrevNext: true, hidePageNumbers: false, perPage: 20 });
            $('#tbodyAllDocDelete').pageMe({ pagerSelector: '#myPagerTab1', showPrevNext: true, hidePageNumbers: false, perPage: 20 });
              $('#tbodyDocUpload').pageMe({ pagerSelector: '#myPagerTab1', showPrevNext: true, hidePageNumbers: false, perPage: 20 });
          //  $('#tbodyNotDesEmp').pageMe({ pagerSelector: '#myPagerNotDesEmp', showPrevNext: true, hidePageNumbers: false, perPage: 20 });
        }
        function fncDestroy(tid, tbname, status, name) {
            document.getElementById("hdfTid").value = tid
            document.getElementById("hdfTabType").value = "empDestroy";
            document.getElementById("hdfStatus").value = status;
            document.getElementById("hdfName").value = name;
            document.getElementById("hdfTbname").value = tbname;
           document.f.target = "iF_Status"
           document.f.method = "post"
           document.f.action = "Query.asp"
           document.f.submit();
            //fncCallDocDestro();
       }

       function fncCallDocDesCondition() {
           $("#processing-modal").modal("show");
           //document.getElementById("hdfTid").value = ""
           document.getElementById("hdfTabType").value = "AllDocDesConEmp";
           document.f.target = "iF_Status"
           document.f.method = "post"
           document.f.action = "Query.asp"
           document.f.submit();

       }

       function fncFilter(data,$$$this){ 
            console.log($($$$this).val());
            var _data = $($$$this).val();
              switch (data) { 
               case "tid" : $("#hdfTid").val(_data);  break; 
               case "name":$("#hdfNameFilter").val(_data);  break;
               case "product":$("#hdfProduct").val(_data);  break;
               case "date":$("#hdfCreateDate").val(_data);  break;
               case "qc":$("#hdfQcBy").val(_data);  break; 
            }
            fncCallDocDesCondition();
        }
        function fncCallDocDestro() {
            $("#processing-modal").modal("show")
            document.getElementById("hdfTid").value = ""
            document.getElementById("hdfTabType").value = "AllEmpDesCon";
            document.f.target = "iF_Status"
            document.f.method = "post"
            document.f.action = "Query.asp"
            document.f.submit();
        }
        function fncNotDestroyEmp() {
            $("#processing-modal").modal("show")
            document.getElementById("hdfTid").value = ""
            document.getElementById("hdfTabType").value = "NotDestroyEmp";
            document.f.target = "iF_Status"
            document.f.method = "post"
            document.f.action = "Query.asp"
            document.f.submit();
        }
        function fncShowUploadDoc(){
            $("#processing-modal").modal("show")
            document.getElementById("hdfTid").value = ""
            document.getElementById("hdfTabType").value = "ShowUploadDoc";
            document.f.target = "iF_Status"
            document.f.method = "post"
            document.f.action = "Query.asp"
            document.f.submit();
        }
        function fncLoadShowListEmpUpload(str) {
             $("#processing-modal").modal("hide")
             $("#tbodyDocUpload").html(str);
             bindPageing();

             $(function(){
                 $(".aDowload").on("click",function(){
                        // saveFile($(this).data("href"));
                          popup($(this).data("href"),'',600,600);
                 });
                  $(".aPreview").on("click",function(){
                     popup($(this).data("href"),'',600,600);
                 });
             });
         }

         function popup(url,name,windowWidth,windowHeight){
            myleft=(screen.width)?(screen.width-windowWidth)/2:100;
            mytop=(screen.height)?(screen.height-windowHeight)/2:100;
            properties = "width="+windowWidth+",height="+windowHeight;
            properties +=",scrollbars=yes, top="+mytop+",left="+myleft;
            window.open(url,name,properties);
        }

         function saveFile(url) {
              var filename = url.substring(url.lastIndexOf("/") + 1).split("?")[0];
              var xhr = new XMLHttpRequest();
              xhr.responseType = 'blob';
              xhr.onload = function() {
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
         function fncResponeAllEmpDes(str) {
             $("#processing-modal").modal("hide")
             $("#tbodyAllDocDelete").html(str);
             bindPageing();
         }

        function fncResponeNotDestroyEmp(str) {
            $("#processing-modal").modal("hide")
            $("#tbodyNotDesEmp").html(str);
            bindPageing();
        }

        function fncResponeAllDocDesCon(str) {  
             $("#processing-modal").modal("hide")
                $("#tbodyDocDes").html(str);
                 console.log("fncResponeAllDocDesCon export emp : ",$("#hdfExportExcelFilter").val());
                if($("#hdfExportExcelFilter").val() == "export"){ 
                   $("#hdfExportExcelFilter").val("") 
                    document.getElementById("hdfTabType").value = "excel";
                    document.getElementById("hdfObjhtml").value = fncGetRowVisible("tbTab2", "tbodyDocDes");
                    document.f.target = "_blank"
                    document.f.method = "post"
                    document.f.action = "excel.asp?xlsName=doc_destroy_emp"
                    document.f.submit();    
                } 
                $("#myPagerTab2").html("")
                $("#tbodyDocDes").pageMe({ pagerSelector: "#myPagerTab2", showPrevNext: true, hidePageNumbers: false, perPage: 20 }); 
        }
        function toUpdate(_tid) {
            $("#processing-modal").modal("show")
            document.getElementById("hdfTid").value = _tid
            document.getElementById("hdfTabType").value = "ScanEmp";
            document.f.target = "iF_Status"
            document.f.method = "post"
            document.f.action = "Query.asp"
            document.f.submit();

        }
        function fncResponeListScan(str) {
            $("#processing-modal").modal("hide");
            $("#trDocScan").html(str);

           if(statusBarcode==true) 
                $("#txt_tid").val("");
            else
              $("#tbIdManual").val("");
               
        }
        function fncRenderPopupconfirmDestroy(count) {
            $("#modalScan").modal("show");
            if (count == 0) {
                $("#contactLabel").text("เอกสารส่งคืน");
                $("#ModalHeader").css("background-color", "D9534F");
               // $('.blink_second').blink({  delay: 800 });
            }
            else {
                $("#contactLabel").text("ตรวจสอบเอกสาร");
                $("#ModalHeader").css("background-color", "337AB7");
            }

        }
        function fncConfirm() {
            $("#modalScan").modal("hide");
            $(".modal-backdrop").remove();
            fncCallDocDestro();

            if(statusBarcode == true) 
                $("#txt_tid").focus();
             else
                $("#tbIdManual").focus();
        }

        function fncUploadFile(folodername) {
            document.Fm.target = "iF_Status"
            document.Fm.method = "post"
            document.Fm.action = "uploadFile.asp?filename=" + folodername;
            document.Fm.submit();

        }
        function fncUpdateDataUpload(){
            //var str = document.getElementById("fileInputCSV").value;
            //var pieces = str.split('\\');
            var filename = ""; //pieces[pieces.length - 1]
            $("#hdfTabType").val("UpdateDataUpload")

            document.f.target =   "iF_Status"
            document.f.method = "post"
            document.f.action = "Query.asp?filename=" + filename;
            document.f.submit();
            alert("upload complete !");
            location.reload();

        }
        function fncReportPenalty() {
            $("#processing-modal").modal("show");
            document.getElementById("hdfTid").value = ""
            document.getElementById("hdfTabType").value = "penaltyEmp";
            document.f.target = "iF_Status"
            document.f.method = "post"
            document.f.action = "Query.asp"
            document.f.submit();
        }
        function fncResponePenalty(str) {
            $("#processing-modal").modal("hide")
            $("#tbodyPenalty").html(str);
            bindPageing();
            loadExpandFnc();
        }
        function loadExpandFnc() {
            $(".btnExpand").on("click", function (e) {
                var id = $(this).attr("data-rowClass");
                $("#" + id).toggle();
                $(this).find("span").toggle();
            });
        }

        function fncExportXlsPenatyEmp() {
            document.getElementById("hdfTabType").value = "excelPenaltyEmp";
            document.f.target = "iF_Status"
            document.f.method = "post"
            document.f.action = "Query.asp"
            document.f.submit();
        }
    </script>
</head>
<%

Function FormatDate(val)
	If val <> "" Then
     FormatDate = right(val,4)&"-"&mid(val,4,2)&"-"&left(val,2)
	Else
	 FormatDate = year(now)&"-"&right("0"&month(now),2)&"-"&right("0"&day(now),2)
	End If
End Function

%>
<!--#include virtual="/destroying/doc/html/progress.htm" -->
<body>
    <form name="Fm" enctype="multipart/form-data">
    <!--- modal popup approve status -->
    <div class="modal fade" id="modalScan" tabindex="-1" role="dialog" aria-labelledby="contactLabel"
        aria-hidden="true">
        <div class="modal-dialog" style="width: 900px;">
            <div class="panel panel-primary">
                <div class="panel-heading" id="ModalHeader">
                    <h4 class="panel-title" id="contactLabel">
                    </h4>
                </div>
                <div class="modal-body" style="padding: 5px;">
                    <!--  <div class="row">
                               <div class="span12"> -->
                    <div class="panel panel-default panel-table" id="dvTb">
                        <div class="panel-heading">
                            <div class="tr">
                                <div class="td" style="font-size: 15px;">
                                    NO</div>
                                <div class="td" style="font-size: 15px;">
                                    TID</div>
                                <div class="td" style="font-size: 15px;">
                                    NAME-LNAME</div>
                                <div class="td" style="font-size: 15px;">
                                    PRODUCT</div>
                                <div class="td" style="font-size: 15px;">
                                    APPLY DATE</div>
                                <div class="td" style="font-size: 15px;">
                                    STATUS DOC</div>
                                <div class="td" style="font-size: 15px;">
                                    DATE</div>
                                <div class="td" style="font-size: 15px;">
                                    DESTROY</div>
                                <div class="td" style="font-size: 15px;">
                                    QC BY</div>
                            </div>
                        </div>
                        <div class="panel-body" id="trDocScan">
                            <div class="tr">
                                <div style="text-align: center; font-size: large; font-weight: bold; margin: 10px 10px 10px 10px;
                                    width: 100px;">
                                    ไม่พบข้อมูล
                                </div>
                            </div>
                        </div>
                    </div>
                    <!-- </div>
                                </div>-->
                </div>
                <div class="panel-footer" style="margin-bottom: -14px; height: 52px; text-align: right;">
                    <input onclick="fncConfirm();" class="btn btn-confirm" type="button" value="confirm" />
                </div>
            </div>
        </div>
    </div>
    <!------ update -->
    <div class="container" style="margin-top: 40px">
        <br />
        <h2>
            ระบบตรวจสอบการทำลายเอกสาร</h2>
        <div class="row">
            <!-- Nav tabs -->
            <ul class="nav nav-tabs" role="tablist">
                <li role="presentation" class="active"><a href="#tabScan" aria-controls="home" role="tab"
                    data-toggle="tab">สแกนบาร์โค้ต</a></li>
                <li role="presentation"><a onclick="fncCallDocDesCondition()" href="#tbDocDestory"
                    aria-controls="profile" role="tab" data-toggle="tab">เอกสารทำลายแล้ว</a></li>
                <li role="presentation"><a onclick="fncNotDestroyEmp()" href="#tbDocNotDestroy" aria-controls="profile"
                    role="tab" data-toggle="tab">เอกสารที่ไม่ต้องทำลาย</a></li>
                <li role="presentation"><a onclick="fncShowUploadDoc()" href="#tbUploadFile" aria-controls="profile"
                    role="tab" data-toggle="tab">อัพโหลดเอกสาร</a></li>
                <li role="presentation"><a onclick="fncReportPenalty()" href="#tbPenalty" aria-controls="profile"
                    role="tab" data-toggle="tab">ค่าปรับ</a></li>
            </ul>
            <div class="tab-content">
                <div role="tabpanel" class="tab-pane active" id="tabScan">
                       <div class="form-group" style="margin-bottom: 0px;">
                        <label for="exampleInputEmail1">
                        </label>
                        <div style="text-align: right;">
                            <button type="button" id="btnKeyManual" class="btn btn-link">
                                <span class="glyphicon glyphicon-hand-left" aria-hidden="true"></span>&nbsp; Manual
                                มือ
                            </button>
                            <button type="button" id="btnScan" style="display: none;" class="btn btn-link">
                                <span class="glyphicon glyphicon-barcode" aria-hidden="true"></span>&nbsp; scan
                                barcode
                            </button>
                        </div>
                        <br />
                        <input type="text" class="form-control" placeholder="สแกนบาร์โค้ตสิค่ะ" name="txt_tid"
                            id="txt_tid" class="Box" size="13" maxlength="10">
                        <!--   <input type="text" class="form-control" placeholder="สแกนบาร์โค้ตสิค่ะ" name="txt_tid"
                            id="txt_tid" class="Box" size="13" maxlength="9" onkeypress="if(num_key(event) == true){toUpdate(this.value);}">-->
                    </div>
                    <div class="form-group">
                        <input type="text" style="display: none;" class="form-control" placeholder="กรอกข้อมูลด้วยมือ"
                            name="tbIdManual" id="tbIdManual" class="Box" size="13" maxlength="10" onkeypress="if(num_key(event) == true){toUpdate(this.value);}">
                    </div>
                    <!-- table เอกสารที่ต้องทำลายทั้งหมด -->
                    <div class="container">
                        <div class="row">
                            <div class="panel panel-primary filterable">
                                <div class="panel-heading">
                                    <h3 class="panel-title">
                                        ข้อมูลเอกสารที่ต้องทำลายทั้งหมด</h3>
                                    <div class="pull-right">
                                        <button type="button" class="btn btn-default btn-xs btn-filter">
                                            <span class="glyphicon glyphicon-filter"></span>Filter</button>
                                    </div>
                                </div>
                                <div style="text-align: right;">
                                    <button type="button" id="btnExport" class="btn btn-link">
                                        <span class="glyphicon glyphicon-list-alt" aria-hidden="true"></span>Export Excel
                                    </button>
                                </div>
                                <table class="table" id="tbname">
                                    <thead>
                                        <tr class="filters">
                                            <th>
                                                No.
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="TID" disabled>
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="NAME-LNAME" disabled>
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="PRODUCT" disabled>
                                            </th>
                                            <th>
                                                STATUS DOC
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="STATUS DATE" disabled>
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="QC BY" disabled>
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody id="tbodyAllDocDelete">
                                    </tbody>
                                </table>
                            </div>
                            <div class="col-md-12 text-right">
                                <ul class="pagination pagination-sm" id="myPagerTab1">
                                </ul>
                            </div>
                        </div>
                    </div>
                    <!-- end table-->
                    <!-- end table เอกสารที่ต้องทำลายทั้งหมด -->
                </div>
                <div role="tabpanel" class="tab-pane" id="tbDocDestory">
                    <div class="container">
                        <!-- start table -->
                        <!--<h3></h3>
                                            <hr> -->
                        <div class="row">
                            <div class="panel panel-primary filterable">
                                <div class="panel-heading">
                                    <h3 class="panel-title">
                                        ข้อมูลที่มีการสแกนและต้องทำลาย</h3>
                                    <div class="pull-right">
                                        <button type="button" class="btn btn-default btn-xs btn-filter">
                                            <span class="glyphicon glyphicon-filter"></span>Filter</button>
                                    </div>
                                </div>
                                <div style="text-align: right;">
                                    <button type="button" id="btnExportExcelDes" class="btn btn-link">
                                        <span class="glyphicon glyphicon-list-alt" aria-hidden="true"></span>Export Excel
                                    </button>
                                </div>
                                <table class="table" id="tbTab2">
                                    <thead>
                                        <tr class="filters">
                                            <th>
                                                No.
                                            </th>
                                            <th>
                                                <input type="text" onchange="fncFilter('tid',this)"  class="form-control" placeholder="TID" disabled>
                                            </th>
                                            <th>
                                                <input type="text" onchange="fncFilter('name',this)"  class="form-control" placeholder="NAME-LNAME" disabled>
                                            </th>
                                            <th>
                                                <input type="text" onchange="fncFilter('product',this)" class="form-control" placeholder="PRODUCT" disabled>
                                            </th>
                                            <th>
                                                STATUS DOC
                                            </th>
                                            <th>
                                                <input type="text" onchange="fncFilter('date',this)" class="form-control" placeholder="DATE" disabled>
                                            </th>
                                            <th>
                                                <input type="text"  onchange="fncFilter('qc',this)" class="form-control" placeholder="QC BY" disabled>
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody id="tbodyDocDes">
                                    </tbody>
                                </table>
                            </div>
                            <div class="col-md-12 text-right">
                                <ul class="pagination pagination-sm" id="myPagerTab2">
                                </ul>
                            </div>
                        </div>
                    </div>
                    <!-- end table-->
                </div>
                <div role="tabpanel" class="tab-pane" id="tbDocNotDestroy">
                    <div class="container">
                        <!-- start table -->
                        <!--<h3></h3>
                                            <hr> -->
                        <div class="row">
                            <div class="panel panel-primary filterable">
                                <div class="panel-heading">
                                    <h3 class="panel-title">
                                        ข้อมูลเอกสารที่ไม่ต้องทำลาย</h3>
                                    <div class="pull-right">
                                        <button type="button" class="btn btn-default btn-xs btn-filter">
                                            <span class="glyphicon glyphicon-filter"></span>Filter</button>
                                    </div>
                                </div>
                                <div style="text-align: right;">
                                    <button type="button" id="btnExportExceltab2" class="btn btn-link">
                                        <span class="glyphicon glyphicon-list-alt" aria-hidden="true"></span>Export Excel
                                    </button>
                                </div>
                                <table class="table" id="tbTabNotDesEmp">
                                    <thead>
                                        <tr class="filters">
                                            <th>
                                                No.
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="TID" disabled>
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="NAME-LNAME" disabled>
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="PRODUCT" disabled>
                                            </th>
                                            <th>
                                                STATUS DOC
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="DATE" disabled>
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="QC BY" disabled>
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody id="tbodyNotDesEmp">
                                    </tbody>
                                </table>
                            </div>
                            <div class="col-md-12 text-right">
                                <ul class="pagination pagination-sm" id="myPagerNotDesEmp">
                                </ul>
                            </div>
                        </div>
                    </div>
                    <!-- end table-->
                </div>
                <div role="tabpanel" class="tab-pane" id="tbUploadFile">
                    <!-- <div class="container">
                                 <iframe name='iframe1' id="iframe1" src="ReadExcel.asp" frameborder="0" border="0" cellspacing="0" style="border-style: none;width: 100%; height: 100%;"></iframe>
                             </div>-->
                    <div class="well">
                        <div class="row">
                            <div class="col-xs-6 col-md-4">
                                เอกสารใบประหน้า</div>
                            <div class="col-xs-6 col-md-4">
                                <!--<input type="FILE" id="fileInputCSV" size="40" name="fupload2"><br>-->
                                <div id="dvAddRows">
                                </div>
                            </div>
                            <div class="col-xs-6 col-md-4">
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-xs-6 col-md-4">
                            </div>
                            <div class="col-xs-6 col-md-4">
                                <input id="btnAddDocUpload" class="btn btn-info" type="button" value="เพิ่มเอกสารอัพโหลด" />
                                <input id="btnUpload" class="btn btn-success" type="button" value="อัพโหลดเอกสาร" />
                            </div>
                            <div class="col-xs-6 col-md-4">
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="panel panel-primary filterable">
                            <div class="panel-heading">
                                <h3 class="panel-title">
                                    เอกสารที่ upload</h3>
                                <div class="pull-right">
                                    <button type="button" class="btn btn-default btn-xs btn-filter">
                                        <span class="glyphicon glyphicon-filter"></span>Filter</button>
                                </div>
                            </div>
                            <table class="table" id="tbDocUpload">
                                <thead>
                                    <tr class="filters">
                                        <th>
                                            No.
                                        </th>
                                        <th>
                                            <input type="text" class="form-control" placeholder="Date Update Doc" disabled>
                                        </th>
                                        <th>
                                            view
                                        </th>
                                        <th>
                                            save
                                        </th>
                                    </tr>
                                </thead>
                                <tbody id="tbodyDocUpload">
                                </tbody>
                            </table>
                        </div>
                        <div class="col-md-12 text-right">
                            <ul class="pagination pagination-sm" id="myPagerDocUpload">
                            </ul>
                        </div>
                    </div>
                </div>
                <div role="tabpanel" class="tab-pane" id="tbPenalty">
                    <div class="container">
                        <div class="row">
                            <div class="panel panel-primary filterable">
                                <div class="panel-heading">
                                    <h3 class="panel-title">
                                        ข้อมูลค่าปรับ</h3>
                                    <!-- <div class="pull-right">
                                                        <button   type="button" class="btn btn-default btn-xs btn-filter"><span class="glyphicon glyphicon-filter"></span> Filter</button>
                                                    </div>-->
                                </div>
                                <div style="text-align: right;">
                                    <button type="button" id="btnExportPenalty" class="btn btn-link">
                                        <span class="glyphicon glyphicon-list-alt" aria-hidden="true"></span>Export Excel
                                    </button>
                                </div>
                                <table class="table" id="tbPenalty">
                                    <thead>
                                        <tr class="filters">
                                            <th>
                                            </th>
                                            <th>
                                                No.
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder="MONTH" disabled>
                                            </th>
                                            <th>
                                                <input type="text" class="form-control" placeholder=" Sum Penalty " disabled>
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody id="tbodyPenalty">
                                    </tbody>
                                </table>
                            </div>
                            <div class="col-md-12 text-right">
                                <ul class="pagination pagination-sm" id="myPagerPenalty">
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!--include virtual="/destroying/doc/html/footer.html" -->
    </form>
    <form name="FmSubUpload" method="post">
    <input type='hidden' name="TabType" value="UpdateWaitSup">
    <div style="display: none;" id="dvHideTid">
    </div>
    </form>
    <form name="f">

    <input type="hidden" value="1" name="ExportExcelFilter" id="hdfExportExcelFilter">  
    <input type="hidden"  name="nameFilter" id="hdfNameFilter"> 
    <input type="hidden" name="productFilter" id="hdfProduct">
    <input type="hidden" name="createDateFilter" id="hdfCreateDate">
    <input type="hidden" name="qcFilter" id="hdfQcBy">

    <input type="hidden" name="tid" id="hdfTid">
    <input type="hidden" name="TabType" id="hdfTabType">
    <input type="hidden" name="status" id="hdfStatus">
    <input type="hidden" name="name" id="hdfName">
    <input type="hidden" name="tbname" id="hdfTbname">
    <input type="hidden" name="objhtml" id="hdfObjhtml">
    </form>
    <script src="js/jsBlink.js" type="text/javascript"></script>
    <link href="css/CSPanel.css" rel="stylesheet" />
    <link href="css/CSTb.css" rel="stylesheet" />
    <script src="js/JSTb.js"></script>
    <script src="js/pHelper.js" type="text/javascript"></script>
    <script src="js/Pageing.js" type="text/javascript"></script>
    <iframe name="iF_Status" width="800" height="100" align="center" frameborder="0">
    </iframe>
    <%set rs = nothing%>
</body>
</html>
<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->
<style>
    .blink_second {
        color: red;
        font-weight: bold;
        font-size: 20px;
    }

    .error-notice {
        margin: 5px 5px; /* Making sure to keep some distance from all side */
    }

    .oaerror {
        width: 90%; /* Configure it fit in your design  */
        margin: 0 auto; /* Centering Stuff */
        background-color: #FFFFFF; /* Default background */
        padding: 20px;
        border: 1px solid #eee;
        border-left-width: 5px;
        border-radius: 3px;
        margin: 0 auto;
        font-family: 'Open Sans' , sans-serif;
        font-size: 16px;
    }

    .danger {
        border-left-color: #d9534f; /* Left side border color */
        background-color: rgba(217, 83, 79, 0.1); /* Same color as the left border with reduced alpha to 0.1 */
    }

    .danger strong {
        color: #d9534f;
    }

    .warning {
        border-left-color: #f0ad4e;
        background-color: rgba(240, 173, 78, 0.1);
    }

    .warning strong {
        color: #f0ad4e;
    }

    .info {
        border-left-color: #5bc0de;
        background-color: rgba(91, 192, 222, 0.1);
    }

    .info strong {
        color: #5bc0de;
    }

    .success {
        border-left-color: #2b542c;
        background-color: rgba(43, 84, 44, 0.1);
    }

    .success strong {
        color: #2b542c;
    }
</style>