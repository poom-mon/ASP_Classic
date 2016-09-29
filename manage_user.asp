<!--#include virtual="/silkspan_ssl/inc/conn.asp"-->
<%
conn.CommandTimeout = 1000
server.scripttimeout = 500
%>
<html>
<head>
    <title>EditUser</title>
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

        $(function () {
            loadUser();
            $("#tbUserLogin_length").css("display", "none");
        });
        //        $(window).load(function () {
        //            $("#tbUserLogin_length").css("display", "none");
        //        });

        function fncSetdate() {
            $("#hdfName").val($("#tbName").val());
            $("#hdfLname").val($("#tbLname").val());
            $("#hdfUsrname").val($("#tbUsername").val());
            $("#hdfPassword").val($("#tbPassword").val());
            $("#hdfUserRole").val($("#ddlUserRole").val());
            $("#hdfEmail").val($("#tbEmail").val());
        }
        function fncDelete(va) {
            var r = confirm("ยืนยันการลบข้อมูล !");
            if (r == true) {
                var id = $(va).data("id");
                $("#hdfId").val(id);
                document.getElementById("hdfTabType").value = "deleteUserLogin";
                document.f.target = "iF_Status"
                document.f.method = "post"
                document.f.action = "Query.asp"
                document.f.submit();
            }
        }
        function fncEdit(va) {
            $("#hdfId").val($(va).data("id"));
            $("#tbName").val($(va).data("name"));
            $("#tbLname").val($(va).data("lname"));
            $("#tbEmail").val($(va).data("email"));
            $("#tbUsername").val($(va).data("username"));
            $("#ddlUserRole").val($(va).data("userrole"));
            $("#tbPassword").val($(va).data("password"));
            $("#tbUsername").prop('disabled', true);

            $("#modalEditTable").modal("show");
            document.getElementById("hdfTabType").value = "EditUsrLogin";

        }
        function fncAdd(va) {
            $("#tbName").val('');
            $("#tbLname").val('');
            $("#tbEmail").val('');
            $("#tbPassword").val('');
            $("#tbUsername").val('');
            $("#tbUsername").prop('disabled', false);
            $("#modalEditTable").modal("show");
            document.getElementById("hdfTabType").value = "AddUserLogin";
        }
        function fncChkUpdate() {
            fncSetdate();
            var status = $("#hdfTabType").val();
            document.f.target = "iF_Status"
            document.f.method = "post"
            document.f.action = "Query.asp"
            document.f.submit();
        }
        function loadUser() {
            // $("#modalScan").modal("show")
            document.getElementById("hdfTabType").value = "loadUserLogin";
            document.f.target = "iF_Status"
            document.f.method = "post"
            document.f.action = "Query.asp"
            document.f.submit();
        }
        function fncRenderTbody(str) {
            //$("#processing-modal").modal("hide")
            $("#tbody").html(str);
        }
        function fncEditSuccess() {
            alert("แก้ไขข้อมูลเรียบร้อยแล้วค่ะ !");
            loadUser();
        }
        function fncAddSuccess() {
            alert("เพิ่มข้อมูลเรียบร้อยแล้วค่ะ !");
            loadUser();
        }
        function fncDeleteSuccess() {
            alert("ลบข้อมูลเรียบร้อยแล้วค่ะ !");
            loadUser();
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
    <form name="f">
    <html>
    <head>
        <link href="//maxcdn.bootstrapcdn.com/bootstrap/3.3.4/css/bootstrap.min.css" rel="stylesheet">
        <script src="//code.jquery.com/jquery-1.11.1.min.js"></script>
        <script src="//cdn.datatables.net/1.10.7/js/jquery.dataTables.min.js"></script>
        <script src="//cdn.datatables.net/plug-ins/1.10.7/integration/bootstrap/3/dataTables.bootstrap.js"></script>
    </head>
    <div class="modal fade" id="modalEditTable" tabindex="-1" role="dialog" aria-labelledby="contactLabel"
        aria-hidden="true">
        <div class="modal-dialog">
            <div class="panel panel-primary">
                <div class="panel-heading">
                    <h4 class="panel-title" id="contactLabel">
                        <span class="glyphicon glyphicon-info-sign"></span>User Info</h4>
                </div>
                <div class="modal-body" style="padding: 5px;">
                    <div class="row" style="margin: 20px 20px 20px 20px;">
                        <div class="col-lg-6">
                            Name</div>
                        <div class="col-lg-6">
                            <input id="tbName" class="form-control" type="text" />
                        </div>
                        <div class="col-lg-6">
                            Last name</div>
                        <div class="col-lg-6">
                            <input id="tbLname" class="form-control" type="text" />
                        </div>
                        <div class="col-lg-6">
                            Email</div>
                        <div class="col-lg-6">
                            <input id="tbEmail" class="form-control" type="text" />
                        </div>
                        <div class="col-lg-6">
                            usename</div>
                        <div class="col-lg-6">
                            <input id="tbUsername"   class="form-control" type="text" />
                        </div>
                        <div class="col-lg-6">
                            USER ROLE</div>
                        <div class="col-lg-6">
                            <select id="ddlUserRole" class="form-control">
                                <option value="sup">SUP</option>
                                <option value="emp">EMP</option>
                                <option value="hr">HR</option>
                            </select>
                        </div>
                        <div class="col-lg-6">
                            password</div>
                        <div class="col-lg-6">
                            <input id="tbPassword" class="form-control" type="text" />
                        </div>
                    </div>
                </div>
                <div class="panel-footer" style="margin-bottom: -14px; height: 52px;">
                    <button style="float: right;" type="button" class="btn btn-default btn-close" data-dismiss="modal">
                        cancel</button>
                    <button style="float: right;" onclick="fncChkUpdate();" type="button" class="btn btn-default btn-close"
                        data-dismiss="modal">
                        confirm</button>
                </div>
            </div>
        </div>
    </div>
    <br />
    <body>
        <div class="container" style="margin-top: 40px">
            <div class="row">
                <!-- Nav tabs -->
                <!-- <ul class="nav nav-tabs" role="tablist">
                    <li role="presentation" class="active"><a href="#tbUser" aria-controls="home" role="tab"
                        data-toggle="tab">ข้อมูลเอกสารที่ต้องทำลายทั้งหมด</a></li>
                   <li role="presentation"><a href="#tabEmail" aria-controls="profile" role="tab" data-toggle="tab">
                        เอกสารทำลายแล้ว</a></li> 
                </ul>-->
                <div class="tab-content">
                    <div role="tabpanel" class="tab-pane active" id="tbUser">
                        <h2>
                            จัดการผู้เข้าใช้ระบบ</h2>
                        <div class="panel panel-primary filterable">
                            <div style="text-align: right;">
                                <button type="button" onclick="fncAdd(this);" class="btn btn-link btnEdit" aria-label="Left Align">
                                    <span class="glyphicon glyphicon-plus" aria-hidden="true">เพิ่มผู้ใช้งาน</span>
                                </button>
                            </div>
                            <div class="panel-heading">
                                <h3 class="panel-title">
                                    ข้อมูลผู้ใช้งาน</h3>
                                <div class="pull-right">
                                    <button type="button" class="btn btn-default btn-xs btn-filter">
                                        <span class="glyphicon glyphicon-filter"></span>Filter</button>
                                </div>
                            </div>
                            <table class="table" id="Table1">
                                <thead>
                                    <tr class="filters">
                                        <th>
                                            No.
                                        </th>
                                        <th>
                                            <input type="text" class="form-control" placeholder="Username" disabled>
                                        </th>
                                        <th>
                                            <input type="text" class="form-control" placeholder="Password" disabled>
                                        </th>
                                        <th>
                                            <input type="text" class="form-control" placeholder="Name" disabled>
                                        </th>
                                        <th>
                                            <input type="text" class="form-control" placeholder="Email" disabled>
                                        </th>
                                        <th>
                                            <input type="text" class="form-control" placeholder="User Role" disabled>
                                        </th>
                                        <th>
                                            update
                                        </th>
                                    </tr>
                                </thead>
                                <tbody id="tbody">
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div role="tabpanel" class="tab-pane" id="tabEmail">
                    <h2>จัดการข้อมูล Email</h2>
                        <ul class="nav nav-tabs" role="tablist">
                            <li role="presentation" class="active"><a href="#tabEditEmail" aria-controls="home" role="tab"
                                data-toggle="tab">เพิ่มข้อมูล email </a></li>
                            <li role="presentation"><a href="#tabAddMail" aria-controls="profile" role="tab" data-toggle="tab">
                                แก้ไขข้อมูล e-mail</a></li>
                        </ul>
                        <div class="tab-content">
                            <div role="tabpanel" class="tab-pane active" id="tabEditEmail">
                                
                                <div class="row" style="margin:20px 20px 20px 20px;" >
                                    <div class="col-lg-4" >
                                      ชืออ้างอิงถึงอีเมล์  </div>
                                    <div class="col-lg-8">
                                        <input id="tbEmailRef" class="form-control" type="text" />
                                    </div> 

                                  <div class="col-lg-4" >
                                      ผู้ส่ง  </div>
                                    <div class="col-lg-8">
                                        <input id="tbEmailSender" class="form-control" type="text" />
                                    </div> 


                                  <div class="col-lg-4" >
                                      subject  </div>
                                    <div class="col-lg-8">
                                        <input id="tbEmailSubject" class="form-control" type="text" />
                                    </div> 

                                   <div class="col-lg-4" >
                                      body  </div>
                                    <div class="col-lg-8">
                                        <input id="tbEmailBody" class="form-control" type="text" />
                                    </div> 



                                </div>

                            </div>
                         <!-- <div role="tabpanel" class="tab-pane active" id="tabAddMail"> 
                            </div>-->
                        </div>
                    </div>

                </div>
            </div>
        </div>
    </body>
    </html>
    <input type="hidden" name="tid" id="hdfTid">
    <input type="hidden" name="TabType" id="hdfTabType">
    <input type="hidden" name="Name" id="hdfName">
    <input type="hidden" name="Lname" id="hdfLname">
    <input type="hidden" name="UserName" id="hdfUsrname">
    <input type="hidden" name="Password" id="hdfPassword">
    <input type="hidden" name="UsrRole" id="hdfUserRole">
    <input type="hidden" name="id" id="hdfId">
    <input type="hidden" name="Email" id="hdfEmail">
    </form>
    <link href="css/CSPanel.css" rel="stylesheet" />
    <link href="css/CSTb.css" rel="stylesheet" />
    <script src="js/JSTb.js"></script>
    <script src="js/bootstrap.min.js" type="text/javascript"></script>
    <script src="js/pHelper.js" type="text/javascript"></script>
    <iframe name="iF_Status" width="800" height="100" align="center" frameborder="0">
    </iframe>
    <%set rs = nothing%>
</body>
</html>
<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->