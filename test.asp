
<script src="js/jquery-1.11.3.min.js" type="text/javascript"></script> 
<!--<script src="js/jquery.MultiFile.js" type="text/javascript"></script>
    <div> 
        <input type="file" class="multi"  />
        <input id="btnSave" type="button" value="button" /> 
    </div> -->
     <link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />
     <form name="Fm" enctype="multipart/form-data">
      
     
         <div id="dvAddRows"></div>
          <input id="btnAdd" type="button" value="add file upload" /> 
           <input id="btnSave" type="button" value="upload" />

             
        <script>
          var i = 1;
            $("#btnAdd").on("click", function () {
                $("#dvAddRows").append("<div class='alert alert-info' style='width:400px;margin-bottom: 5px;'>  <input type='FILE'  size='40' name='fupload" + i + "'></div>  "); 
              i=i+1;
            });

            $("#btnSave").on("click", function () {
               // $("#MultiFile1_wrap_list > .alert >.MultiFile-title ").each(function (index) {
               //     console.log(index + ":" + $(this).text() + ":" + $(this).attr("title").substring(15, $(this).attr("title").lenght));
                  fncUploadFile("xxxx");
               // });

            });
            function fncUpdateDataUpload() {
                var str = document.getElementById("fileInputCSV").value;
                var pieces = str.split('\\');
                var filename = pieces[pieces.length - 1]  
                document.f.target = "iF_Status"
                document.f.method = "post"
                document.f.action = "Query.asp?filename=" + filename;
                document.f.submit();
                alert("upload complete !");

            }
            function fncUploadFile(folodername) {
                document.Fm.target = "iF_Status"
                document.Fm.method = "post"
                document.Fm.action = "uploadFile.asp?filename=" + folodername;
                document.Fm.submit();

            }
        </script>
       
            <iframe name="iF_Status" width="800" height="100" align="center" frameborder="0">
            </iframe>
        </form>