<!--#include virtual="/silkspan_ssl/inc/conn.asp"-->
<%
conn.CommandTimeout = 1000
server.scripttimeout = 500
%>
<HTML>
<HEAD>
<TITLE>DESTROYING DOCUMENTS REPORT</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
  <!--#include virtual="/destroying/doc/html/head.html" -->   
<script language="javascript"> 
function num_key(evt){
	var iKeyCode;
	var IsValid = false;

	if(window.event) { // IE
		iKeyCode = evt.keyCode	
	}
	else if(evt.which)	{ // Netscape/Firefox/Opera
		iKeyCode = evt.which
	}
	//alert(iKeyCode)
	if (iKeyCode == 13){
		return true;
	}
	else{
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
    $("#btnExport").on("click", function () {
        document.getElementById("hdfTabType").value = "excel";
        document.getElementById("hdfObjhtml").value = $("#tbname").html();
        document.f.target = "_blank"
        document.f.method = "post"
        document.f.action = "excel.asp?xlsName=doc_destroy"
        document.f.submit();
    });
    
    $("#btnExportExceltab2").on("click", function () {  
       $("#hdfExportExcelFilter").val("export")  
       fncCallDocDesCondition()   
    });
    $("#btnExportTab3").on("click", function () {
        document.getElementById("hdfTabType").value = "excel";
        document.getElementById("hdfObjhtml").value = fncGetRowVisible("tbTab3", "tbodyDocCon");
        document.f.target = "_blank"
        document.f.method = "post"
        document.f.action = "excel.asp?xlsName=doc_notdestroy"
        document.f.submit();
    });
    $("#btnExportPenalty").on("click", function () { 
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

function fnbcExportExcel(str) {
    document.getElementById("hdfTabType").value = "excel";
    document.getElementById("hdfObjhtml").value = str;
    //document.getElementById("hdfObjhtml").value = fncGetRowVisible("tbPenalty", "tbodyPenalty");
    document.f.target = "_blank"
    document.f.method = "post"
    document.f.action = "excel.asp?xlsName=doc_penalty"
    document.f.submit();
}
function fncExportXlsPenatyEmp() {
    document.getElementById("hdfTabType").value = "excelPenalty";
    document.f.target = "iF_Status"
    document.f.method = "post"
    document.f.action = "Query.asp"
    document.f.submit();
}

function loadExpandFnc() {
    $(".btnExpand").on("click", function (e) {
        var id = $(this).attr("data-rowClass");
        $("#" + id).toggle();
        $(this).find("span").toggle(); 
    });
}

function toUpdate(_tid) {
    $("#processing-modal").modal("show")
    document.getElementById("hdfTid").value = _tid
    document.getElementById("hdfTabType").value = "Scan";
		document.f.target = "iF_Status"
		document.f.method = "post"
		document.f.action = "Query.asp"
		document.f.submit();
		
}
function fncCallReport(_tid) {
   
    $("#processing-modal").modal("show")
    document.getElementById("hdfTid").value = _tid
    document.getElementById("hdfTabType").value = "Report";
    document.f.target = "iF_Status"
    document.f.method = "post"
    document.f.action = "Query.asp"
    document.f.submit();
}
function fncCallDocDestro() { 
   $("#processing-modal").modal("show")
   document.getElementById("hdfTid").value = ""
    document.getElementById("hdfTabType").value = "DocDestroy";
    document.f.target = "iF_Status"
    document.f.method = "post"
    document.f.action = "Query.asp"
    document.f.submit();

    
}
function fncRenderPopupconfirmDestroy(count) { 
  
    $("#modalScan").modal("show");
    if (count == 0) {
        $("#contactLabel").text("เอกสารส่งคืน");
        $("#ModalHeader").css("background-color", "#D9534F"); 
       // $('.blink_second').blink({ delay: 800 });
    }
    else {
        $("#contactLabel").text("ตรวจสอบเอกสาร");
        $("#ModalHeader").css("background-color", "#337AB7"); 
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
function fncCallDocCondition() {
    $("#processing-modal").modal("show");
    document.getElementById("hdfTid").value = ""
    document.getElementById("hdfTabType").value = "DocDesCon";
    document.f.target = "iF_Status"
    document.f.method = "post"
    document.f.action = "Query.asp"
    document.f.submit(); 
}
function fncCallDocDesCondition() {
    $("#processing-modal").modal("show");
  //  document.getElementById("hdfTid").value = ""
    document.getElementById("hdfTabType").value = "AllDocDesCon";
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
function fncReportPenalty() {
    $("#processing-modal").modal("show");
    document.getElementById("hdfTid").value = ""
    document.getElementById("hdfTabType").value = "penalty";
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
function fncResponeAllDocDesctroy(str) {
    $("#processing-modal").modal("hide")
    $("#tbodyAllDocDelete").html(str);
    bindPageing();
}
function fncResponePenalty(str) {
    $("#processing-modal").modal("hide")
    $("#tbodyPenalty").html(str);
    bindPageing();
    loadExpandFnc();
}
function fncResponeAllDocCon(str) {
    $("#processing-modal").modal("hide")
    $("#tbodyDocCon").html(str);
    bindPageing();
}
function fncResponeAllDocDesCon(str) {
    $("#processing-modal").modal("hide")
    $("#tbodyDocDes").html(str);
     console.log("fncResponeAllDocDesCon export : ",$("#hdfExportExcelFilter").val());
    if($("#hdfExportExcelFilter").val() == "export"){ 
       $("#hdfExportExcelFilter").val("") 
        document.getElementById("hdfTabType").value = "excel";
        document.getElementById("hdfObjhtml").value = fncGetRowVisible("tbTab2", "tbodyDocDes");
        document.f.target = "_blank"
        document.f.method = "post"
        document.f.action = "excel.asp?xlsName=doc_scandestroy"
        document.f.submit();  
    } 
    $("#myPagerTab2").html("")
    $('#tbodyDocDes').pageMe({ pagerSelector: '#myPagerTab2', showPrevNext: true, hidePageNumbers: false, perPage: 20 });
     
   // bindPageing();
}
function fncAlert(str) {
    alert(str);
} 
 
</script>
</HEAD>
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
 <BODY>
<form name="f">
  

 <!--- modal popup approve status -->
    <div class="modal fade" id="modalScan" tabindex="-1" role="dialog" aria-labelledby="contactLabel" aria-hidden="true">
                <div class="modal-dialog" style=" width: 900px;">
                    <div class="panel panel-primary" id="panelTitle">
                        <div class="panel-heading" id="ModalHeader" > 
                            <h4 class="panel-title" id="contactLabel"><span class="glyphicon glyphicon-info-sign"></span> comfirm</h4>
                        </div> 
                        <div class="modal-body" style="padding: 5px;">
                            <!--  <div class="row">
                               <div class="span12"> -->
                           
                                         <div class="panel panel-default panel-table" id="dvTb">
                                        <div class="panel-heading">
                                            <div class="tr">
                                                <div class="td" style="font-size:15px;">NO</div>
                                                <div class="td" style="font-size:15px;">TID</div>
                                                <div class="td" style="font-size:15px;">NAME-LNAME</div>
                                                <div class="td" style="font-size:15px;">PRODUCT</div> 
                                                <div class="td" style="font-size:15px;">APPLY DATE</div>
                                                <div class="td" style="font-size:15px;">STATUS DOC</div>
                                                <div class="td" style="font-size:15px;">DATE</div>
                                                <div class="td" style="font-size:15px;">DESTROY</div>
                                                <div class="td" style="font-size:15px;">QC BY</div>
                                            </div>
                                        </div>
                                        <div class="panel-body"  id="trDocScan">
                                            <div class="tr"  >
                                               <div style="text-align:center;font-size:large;font-weight:bold;margin:10px 10px 10px 10px;width:100px;">
                                                  ไม่พบข้อมูล
                                                  </div>
                                            </div>
                                        </div> 
                                    </div>
                                <!-- </div>
                                </div>-->
                                
                            </div>  
                            <div class="panel-footer" style="margin-bottom:-14px;height: 52px;text-align:right;">
                               <!-- <input type="submit" class="btn btn-success" value="Send"/> 
                                <input type="reset" class="btn btn-danger" value="Clear" /> -->
                                <input onclick="fncConfirm();" class="btn btn-confirm" type="button" value="confirm" />
 
                            </div>
                        </div>
                    </div>
                </div> 
 <!------ update -->



           <div class="container" style="margin-top: 40px">
                <br />
                <h2>ระบบตรวจสอบการทำลายเอกสาร</h2>
                <div class="row">
                    <!-- Nav tabs -->
                    <ul class="nav nav-tabs" role="tablist">
                        <li role="presentation" class="active"><a href="#tabScan"    aria-controls="home" role="tab" data-toggle="tab">สแกนบาร์โค้ต</a></li>
                         <li role="presentation"><a onclick="fncCallDocDesCondition()" href="#tbDocDestory" aria-controls="profile" role="tab" data-toggle="tab">เอกสารต้องทำลาย</a></li>
                         <li role="presentation"><a onclick="fncCallDocCondition()" href="#tbDocDesCon" aria-controls="profile" role="tab" data-toggle="tab">เอกสารไม่อยู่เงือนไขทำลาย</a></li>
                        <li role="presentation"><a onclick="fncReportPenalty()" href="#tbPenalty" aria-controls="profile" role="tab" data-toggle="tab">ค่าปรับ</a></li>

                    </ul>

                    <!-- Tab panes -->
                    <div class="tab-content">
                        <div role="tabpanel" class="tab-pane active" id="tabScan">
                            <div class="form-group" style="margin-bottom: 0px;">
                                <label for="exampleInputEmail1"></label><br /> 
                                <div style="text-align:right;">
                                      <button type="button" id="btnKeyManual" class="btn btn-link">
                                     <span class="glyphicon glyphicon-hand-left" aria-hidden="true"></span>&nbsp; Manual มือ
                                     </button> 
                                      <button type="button" id="btnScan"  style="display:none;" class="btn btn-link">
                                     <span class="glyphicon glyphicon-barcode" aria-hidden="true"></span>&nbsp; scan barcode 
                                     </button>
                                 </div>

<!--                                 <input type="text" class="form-control" placeholder="สแกนบาร์โค้ตสิค่ะ"  name="txt_tid" id="txt_tid" class="Box" size="13" maxlength="9" onkeypress="if(num_key(event) == true){toUpdate(this.value);}">
-->
                                 <input type="text" class="form-control" placeholder="สแกนบาร์โค้ตสิค่ะ"  name="txt_tid" id="txt_tid" class="Box" size="13" maxlength="10">
                                  
                                  
                            </div> 

                              <div class="form-group">
                                <input type="text" style="display:none;" class="form-control" placeholder="กรอกข้อมูลด้วยมือ"  name="tbIdManual" id="tbIdManual" class="Box" size="13" maxlength="10" onkeypress="if(num_key(event) == true){toUpdate(this.value);}">
                             </div>
                    

                            <!-- table เอกสารที่ต้องทำลายทั้งหมด --> 

                            <div class="container">  
                                        <div class="row">   
                                            <div class="panel panel-primary filterable">
                                                <div class="panel-heading">
                                                    <h3 class="panel-title">ข้อมูลเอกสารที่ต้องทำลายทั้งหมด</h3>
                                                    <div class="pull-right">
                                                        <button   type="button" class="btn btn-default btn-xs btn-filter"><span class="glyphicon glyphicon-filter"></span> Filter</button>
                                                    </div>
                                                </div>
                                                 
                                                 <div  style="text-align:right;"><button type="button"  id="btnExport" class="btn btn-link">
                                            <span class="glyphicon glyphicon-list-alt" aria-hidden="true"></span> Export Excel
                                            </button></div>

                                                <table class="table" id="tbname">
                                                    <thead>
                                                        <tr class="filters">
                                                            <th>No.</th>
                                                            <th><input type="text" class="form-control" placeholder="TID" disabled></th>
                                                            <th><input type="text" class="form-control" placeholder="NAME-LNAME" disabled></th>
                                                            <th><input type="text" class="form-control" placeholder="PRODUCT" disabled></th> 
                                                            <th><input type="text" class="form-control" placeholder="APPLY DATE" disabled></th>
                                                            <th><input type="text" class="form-control" placeholder="STATUS DOC" disabled></th>
                                                            <th>STATUS DATE</th> 
                                                            <th><input type="text" class="form-control" placeholder="QC BY" disabled> </th>
                                                        </tr>
                                                    </thead>
                                                    <tbody id="tbodyAllDocDelete"> 
                                                    </tbody>
                                                </table>
                                                  

                                            </div>
                                              <div class="col-md-12 text-right" >
                                                   <ul class="pagination pagination-sm" id="myPagerTab1"></ul>
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
                                                        <h3 class="panel-title">ข้อมูลที่มีการสแกนและต้องทำลาย</h3>
                                                        <div class="pull-right">
                                                            <button   type="button" class="btn btn-default btn-xs btn-filter"><span class="glyphicon glyphicon-filter"></span> Filter</button>
                                                        </div>
                                                    </div>

                                                    <div  style="text-align:right;"><button type="button"  id="btnExportExceltab2" class="btn btn-link">
                                                        <span class="glyphicon glyphicon-list-alt" aria-hidden="true"></span> Export Excel
                                                        </button>
                                                    </div>

                                                    <table class="table" id="tbTab2">
                                                        <thead>
                                                            <tr class="filters">
                                                                <th>No.</th>
                                                                <th><input type="text" onchange="fncFilter('tid',this)" class="form-control" placeholder="TID" disabled></th>
                                                                <th><input type="text" onchange="fncFilter('name',this)" class="form-control" placeholder="NAME-LNAME" disabled></th>
                                                                <th><input type="text" onchange="fncFilter('product',this)" class="form-control" placeholder="PRODUCT" disabled></th>  
                                                                <th>STATUS DOC</th>
                                                                <th><input type="text" onchange="fncFilter('date',this)"  class="form-control" placeholder="DATE" disabled></th> 
                                                                <th><input type="text" onchange="fncFilter('qc',this)"  class="form-control" placeholder="QC BY" disabled /></th>
                                                            </tr>
                                                        </thead>
                                                        <tbody id="tbodyDocDes"> 
                                                        </tbody>
                                                    </table>
                                                </div>
                                                  <div class="col-md-12 text-right" >
                                                   <ul class="pagination pagination-sm" id="myPagerTab2"></ul>
                                              </div>

                                            </div>
                                        </div>
                                        <!-- end table-->
                        </div>

                        <div role="tabpanel" class="tab-pane" id="tbDocDesCon">
                                                <div class="container">
                        <!-- start table -->
                                        <!--<h3></h3>
                                        <hr> -->
                                        <div class="row">
                                            <div class="panel panel-primary filterable">
                                                <div class="panel-heading">
                                                    <h3 class="panel-title">ข้อมูลที่มีการสแกนและทำลายไม่ได้</h3>
                                                    <div class="pull-right">
                                                        <button   type="button" class="btn btn-default btn-xs btn-filter"><span class="glyphicon glyphicon-filter"></span> Filter</button>
                                                    </div>
                                                </div>
                                                  <div  style="text-align:right;"><button type="button"  id="btnExportTab3" class="btn btn-link">
                                                        <span class="glyphicon glyphicon-list-alt" aria-hidden="true"></span> Export Excel
                                                        </button>
                                                    </div>

                                                <table class="table" id="tbTab3">
                                                    <thead>
                                                        <tr class="filters">
                                                            <th>No.</th>
                                                            <th><input type="text" class="form-control" placeholder="TID" disabled></th>
                                                            <th><input type="text" class="form-control" placeholder="NAME-LNAME" disabled></th>
                                                            <th><input type="text" class="form-control" placeholder="PRODUCT" disabled></th>  
                                                            <th>STATUS DOC</th>
                                                            <th><input type="text" class="form-control" placeholder="DATE" disabled></th> 
                                                             <th><input type="text" class="form-control" placeholder="QC BY" disabled></th>  
                                                        </tr>
                                                    </thead>
                                                    <tbody id="tbodyDocCon"> 
                                                    </tbody>
                                                </table>
                                            </div>

                                              <div class="col-md-12 text-right" >
                                                   <ul class="pagination pagination-sm" id="myPagerTab3"></ul>
                                              </div>
											  
                                        </div>
                                    </div>
                                    <!-- end table-->
                        </div> 

                     
                         <div role="tabpanel" class="tab-pane" id="tbPenalty">
                                <div class="container">
                                    <div class="row">
                                            <div class="panel panel-primary filterable">
                                                <div class="panel-heading">
                                                    <h3 class="panel-title">ข้อมูลค่าปรับ</h3>
                                                   <!-- <div class="pull-right">
                                                        <button   type="button" class="btn btn-default btn-xs btn-filter"><span class="glyphicon glyphicon-filter"></span> Filter</button>
                                                    </div>-->
                                                </div>
                                                  <div  style="text-align:right;"><button type="button"  id="btnExportPenalty" class="btn btn-link">
                                                        <span class="glyphicon glyphicon-list-alt" aria-hidden="true"></span> Export Excel
                                                        </button>
                                                    </div>

                                                <table class="table" id="tbPenalty">
                                                    <thead>
                                                        <tr class="filters">
                                                           <th></th>
                                                            <th>No.</th>
                                                            <th><input type="text" class="form-control" placeholder="MONTH" disabled></th>
                                                            <th><input type="text" class="form-control" placeholder=" Sum Penalty " disabled></th>   
                                                        </tr>
                                                    </thead>
                                                    <tbody id="tbodyPenalty"> 
                                                    </tbody>
                                                </table>
                                            </div>

                                              <div class="col-md-12 text-right" >
                                                   <ul class="pagination pagination-sm" id="myPagerPenalty"></ul>
                                              </div>
											  
                                        </div>
                                </div> 
                            </div>

                    </div>
                </div>

 
            </div>
     

      <input type="hidden" value="1" name="ExportExcelFilter" id="hdfExportExcelFilter">  
      <input type="hidden"  name="nameFilter" id="hdfNameFilter"> 
      <input type="hidden" name="productFilter" id="hdfProduct">
      <input type="hidden" name="createDateFilter" id="hdfCreateDate">
      <input type="hidden" name="qcFilter" id="hdfQcBy">

    <input type="hidden" name="tid" id="hdfTid">
    <input type="hidden" name="TabType" id="hdfTabType"> 
    <input type="hidden" name="objhtml" id="hdfObjhtml">


    <!--include virtual="/destroying/doc/html/footer.html" -->
</form> 
  
   <link href="css/CSPanel.css" rel="stylesheet" />
   <link href="css/CSTb.css" rel="stylesheet" />
   <script src="js/JSTb.js"></script>
   <script src="js/pHelper.js" type="text/javascript"></script>
    <script src="js/Pageing.js" type="text/javascript"></script>
    <script src="js/jsBlink.js" type="text/javascript"></script>
    <script>
//        $(window).load(function () {

//        });
        function bindPageing() {
            $("#myPagerPenalty,#myPagerTab1,#myPagerTab2,#myPagerTab3").html("")
            $('#tbodyAllDocDelete').pageMe({ pagerSelector: '#myPagerTab1', showPrevNext: true, hidePageNumbers: false, perPage: 20 });
           // $('#tbodyDocDes').pageMe({ pagerSelector: '#myPagerTab2', showPrevNext: true, hidePageNumbers: false, perPage: 20 });
            $('#tbodyDocCon').pageMe({ pagerSelector: '#myPagerTab3', showPrevNext: true, hidePageNumbers: false, perPage: 20 });
            $('#tbPenalty').pageMe({ pagerSelector: '#myPagerPenalty', showPrevNext: true, hidePageNumbers: false, perPage: 20 });
        }
    </script>
    <style>
            #tbPenalty input.form-control {
                text-align: center;
        }
            .blink_second {
                color:red;
                font-weight:bold;
                font-size:20px;
            } 
    </style>
<iframe name="iF_Status" width="800" height="100" align="center" frameborder="0"></iframe> 
 <%set rs = nothing%>
</BODY>
</HTML> 

<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->
 