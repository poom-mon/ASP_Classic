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

        $(function () {
            $("#btnLogin").on("click", function () {
                fncLogin();
            });
            $("#tbPassword").keypress(function (e) {
                if (e.which == 13) {
                    fncLogin();
                }
            });
          
        });

       

        function fncLogin() { 
          var _usrname = $("#tbUsername").val();
          var _password = $("#tbPassword").val();
          document.getElementById("hdfTabType").value = "login";
           document.getElementById("hdfUsername").value = _usrname;
          document.getElementById("hdfPassword").value = _password; 
           document.f.target = "iF_Status"
           document.f.method = "post"
           document.f.action = "Query.asp"
           document.f.submit(); 
       }

       function fncLoginFail() {
           alert("Sorry loginFail");
       }
       function fncLoginSuccess(url) {
           window.location.href = url;
       }
    </script>
</head>
<%
 


if Request("logout")  ="true"  then
 	Session("usrname") =""
    Session("typeEmp") =""
    Session("name") = "" 
    Response.Redirect("login.asp")
end if 


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
            <div class="container">
                <div class="row">
                    <div class="col-md-6 col-md-offset-3">
                        <div class="panel panel-login">
                            <div class="panel-heading">
                                <div class="row">
                                    <div class="col-xs-12">
                                        <h2>เข้าสู่ระบบ</h2>
                                    </div>
                                     
                                </div>
                                <hr>
                            </div>
                            <div class="panel-body">
                                <div class="row">
                                    <div class="col-lg-12"> 
                                            <div class="form-group">
                                                <input type="text" name="username" id="tbUsername" tabindex="1" class="form-control" placeholder="Username" value="">
                                            </div>
                                            <div class="form-group">
                                                <input type="password" name="password" id="tbPassword" tabindex="2" class="form-control" placeholder="Password">
                                            </div>
                                            
                                            <div class="form-group">
                                                <div class="row">
                                                    <div class="col-sm-6 col-sm-offset-3">
                                                        <input type="button" name="login-submit"    id="btnLogin" tabindex="4" class="form-control btn btn-login" value="Log In">
                                                    </div>
                                                </div>
                                            </div>  
                                          
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <style>
                body {
                    padding-top: 90px;
                }

                .panel-login {
                    border-color: #ccc;
                    -webkit-box-shadow: 0px 2px 3px 0px rgba(0,0,0,0.2);
                    -moz-box-shadow: 0px 2px 3px 0px rgba(0,0,0,0.2);
                    box-shadow: 0px 2px 3px 0px rgba(0,0,0,0.2);
                }

                    .panel-login > .panel-heading {
                        color: #00415d;
                        background-color: #fff;
                        border-color: #fff;
                        text-align: center;
                    }

                        .panel-login > .panel-heading a {
                            text-decoration: none;
                            color: #666;
                            font-weight: bold;
                            font-size: 15px;
                            -webkit-transition: all 0.1s linear;
                            -moz-transition: all 0.1s linear;
                            transition: all 0.1s linear;
                        }

                            .panel-login > .panel-heading a.active {
                                color: #029f5b;
                                font-size: 18px;
                            }

                        .panel-login > .panel-heading hr {
                            margin-top: 10px;
                            margin-bottom: 0px;
                            clear: both;
                            border: 0;
                            height: 1px;
                            background-image: -webkit-linear-gradient(left,rgba(0, 0, 0, 0),rgba(0, 0, 0, 0.15),rgba(0, 0, 0, 0));
                            background-image: -moz-linear-gradient(left,rgba(0,0,0,0),rgba(0,0,0,0.15),rgba(0,0,0,0));
                            background-image: -ms-linear-gradient(left,rgba(0,0,0,0),rgba(0,0,0,0.15),rgba(0,0,0,0));
                            background-image: -o-linear-gradient(left,rgba(0,0,0,0),rgba(0,0,0,0.15),rgba(0,0,0,0));
                        }

                    .panel-login input[type="text"], .panel-login input[type="email"], .panel-login input[type="password"] {
                        height: 45px;
                        border: 1px solid #ddd;
                        font-size: 16px;
                        -webkit-transition: all 0.1s linear;
                        -moz-transition: all 0.1s linear;
                        transition: all 0.1s linear;
                    }

                    .panel-login input:hover,
                    .panel-login input:focus {
                        outline: none;
                        -webkit-box-shadow: none;
                        -moz-box-shadow: none;
                        box-shadow: none;
                        border-color: #ccc;
                    }

                .btn-login {
                    background-color: #59B2E0;
                    outline: none;
                    color: #fff;
                    font-size: 14px;
                    height: auto;
                    font-weight: normal;
                    padding: 14px 0;
                    text-transform: uppercase;
                    border-color: #59B2E6;
                }

                    .btn-login:hover,
                    .btn-login:focus {
                        color: #fff;
                        background-color: #53A3CD;
                        border-color: #53A3CD;
                    }

                .forgot-password {
                    text-decoration: underline;
                    color: #888;
                }

                    .forgot-password:hover,
                    .forgot-password:focus {
                        text-decoration: underline;
                        color: #666;
                    }

                .btn-register {
                    background-color: #1CB94E;
                    outline: none;
                    color: #fff;
                    font-size: 14px;
                    height: auto;
                    font-weight: normal;
                    padding: 14px 0;
                    text-transform: uppercase;
                    border-color: #1CB94A;
                }

                    .btn-register:hover,
                    .btn-register:focus {
                        color: #fff;
                        background-color: #1CA347;
                        border-color: #1CA347;
                    }
            </style>
 
    <input type="hidden" name="TabType" id="hdfTabType"> 
    <input type="hidden" name="usr_login" id="hdfUsername">
    <input type="hidden" name="pass_login" id="hdfPassword">  

    <!--#include virtual="/destroying/doc/html/footer.html" -->
    </form> 
    <iframe name="iF_Status" width="800" height="100" align="center" frameborder="0">
    </iframe>
    <%set rs = nothing%>
</body>
</html>
<!--#include virtual="/silkspan_ssl/inc/closeconn.asp" -->
    <style>
                body {
                    padding-top: 90px;
                }

                .panel-login {
                    border-color: #ccc;
                    -webkit-box-shadow: 0px 2px 3px 0px rgba(0,0,0,0.2);
                    -moz-box-shadow: 0px 2px 3px 0px rgba(0,0,0,0.2);
                    box-shadow: 0px 2px 3px 0px rgba(0,0,0,0.2);
                }

                    .panel-login > .panel-heading {
                        color: #00415d;
                        background-color: #fff;
                        border-color: #fff;
                        text-align: center;
                    }

                        .panel-login > .panel-heading a {
                            text-decoration: none;
                            color: #666;
                            font-weight: bold;
                            font-size: 15px;
                            -webkit-transition: all 0.1s linear;
                            -moz-transition: all 0.1s linear;
                            transition: all 0.1s linear;
                        }

                            .panel-login > .panel-heading a.active {
                                color: #029f5b;
                                font-size: 18px;
                            }

                        .panel-login > .panel-heading hr {
                            margin-top: 10px;
                            margin-bottom: 0px;
                            clear: both;
                            border: 0;
                            height: 1px;
                            background-image: -webkit-linear-gradient(left,rgba(0, 0, 0, 0),rgba(0, 0, 0, 0.15),rgba(0, 0, 0, 0));
                            background-image: -moz-linear-gradient(left,rgba(0,0,0,0),rgba(0,0,0,0.15),rgba(0,0,0,0));
                            background-image: -ms-linear-gradient(left,rgba(0,0,0,0),rgba(0,0,0,0.15),rgba(0,0,0,0));
                            background-image: -o-linear-gradient(left,rgba(0,0,0,0),rgba(0,0,0,0.15),rgba(0,0,0,0));
                        }

                    .panel-login input[type="text"], .panel-login input[type="email"], .panel-login input[type="password"] {
                        height: 45px;
                        border: 1px solid #ddd;
                        font-size: 16px;
                        -webkit-transition: all 0.1s linear;
                        -moz-transition: all 0.1s linear;
                        transition: all 0.1s linear;
                    }

                    .panel-login input:hover,
                    .panel-login input:focus {
                        outline: none;
                        -webkit-box-shadow: none;
                        -moz-box-shadow: none;
                        box-shadow: none;
                        border-color: #ccc;
                    }

                .btn-login {
                    background-color: #59B2E0;
                    outline: none;
                    color: #fff;
                    font-size: 14px;
                    height: auto;
                    font-weight: normal;
                    padding: 14px 0;
                    text-transform: uppercase;
                    border-color: #59B2E6;
                }

                    .btn-login:hover,
                    .btn-login:focus {
                        color: #fff;
                        background-color: #53A3CD;
                        border-color: #53A3CD;
                    }

                .forgot-password {
                    text-decoration: underline;
                    color: #888;
                }

                    .forgot-password:hover,
                    .forgot-password:focus {
                        text-decoration: underline;
                        color: #666;
                    }

                .btn-register {
                    background-color: #1CB94E;
                    outline: none;
                    color: #fff;
                    font-size: 14px;
                    height: auto;
                    font-weight: normal;
                    padding: 14px 0;
                    text-transform: uppercase;
                    border-color: #1CB94A;
                }

                    .btn-register:hover,
                    .btn-register:focus {
                        color: #fff;
                        background-color: #1CA347;
                        border-color: #1CA347;
                    }
            </style>