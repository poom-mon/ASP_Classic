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
////  
////$(function () { 
//////    $("#btnExport").on("click", function () {
//////        document.getElementById("hdfTabType").value = "excel";
//////        document.getElementById("hdfObjhtml").value = $("#tbname").html();
//////        document.f.target = "_blank"
//////        document.f.method = "post"
//////        document.f.action = "excel.asp?xlsName=doc_destroy"
//////        document.f.submit();
//////    });  
////});
 
function toUpdate(_tid) { 
    document.getElementById("hdfTid").value = _tid
    document.getElementById("hdfTabType").value = "tmpScan";
	document.f.target = "iF_Status"
	document.f.method = "post"
	document.f.action = "Query.asp"
	document.f.submit(); 
}
function RenderResult(_tid) {
    var e = formatDate(d);
    strRows = "";
    strRows = strRows + "<div><b>เอกสารที่สแกน</b></div>"
    strRows = strRows + "<div>tid:"+_tid+"</div>"
    strRows = strRows + "<div>time:"+e+"</div>"
    $("#dvResult").html(strRows);
    $("#txt_tid").val("");
}

var d = new Date();
function formatDate(date) {
    var hours = date.getHours();
    var minutes = date.getMinutes();
    var ampm = hours >= 12 ? 'pm' : 'am';
    hours = hours % 12;
    hours = hours ? hours : 12; // the hour '0' should be '12'
    minutes = minutes < 10 ? '0' + minutes : minutes;
    var strTime = hours + ':' + minutes + ' ' + ampm;
    return date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getFullYear() + "  " + strTime;
}

 
</script>
</HEAD> 
 

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


 <br /><br />
<br /><br /><br /> 
           <div class="container" style="margin-top: 40px"> 
                        <div role="tabpanel" class="tab-pane active" id="tabScan">
                            <div class="form-group">
                                <label for="exampleInputEmail1"></label><br /> 
                                 <input type="text" class="form-control" placeholder="สแกนบาร์โค้ตสิค่ะ"  name="txt_tid" id="txt_tid" class="Box" size="13" maxlength="9"
                                 style=" height: 64.22222137451172px;"
                                  onkeypress="if(num_key(event) == true){toUpdate(this.value);}">

                                  
                            </div>
                            <div class="alert alert-info"  id="dvResult" role="alert">  
                            </div>
 
                        </div>
                         
                    </div> 

  
     
      

    <input type="hidden" name="tid" id="hdfTid">
    <input type="hidden" name="TabType" id="hdfTabType"> 
    <input type="hidden" name="objhtml" id="hdfObjhtml">


    <!--include virtual="/destroying/doc/html/footer.html" -->
</form>  
   <script src="js/pHelper.js" type="text/javascript"></script> 
     
     
<iframe name="iF_Status" width="800" height="100" align="center" frameborder="0"></iframe> 
 <%set rs = nothing%>
</BODY>
</HTML> 

<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->
 