<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID="1033"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<style type="text/css"> 
    .banner {overflow: hidden;}
     .cImg {border-style:none;}
     body { margin:0px;padding:0px; } 
</style>
  <script src="jquery-1.10.2.js" type="text/javascript"></script>
  <%

        ''Random function 
            function Shuffle (ByRef arrInput) 
	        Dim arrIndices, iSize, x
	        Dim arrOriginal 
	        iSize = UBound(arrInput)+1
	 
	        arrIndices = RandomNoDuplicates(0, iSize-1, iSize)
	 
	        arrOriginal = CopyArray(arrInput)
	 
	        For x=0 To UBound(arrIndices)
		        arrInput(x) = arrOriginal(arrIndices(x))
	        Next
            Shuffle = arrInput
        End function

        ''Randdom áººäÁè«Óé
        Function RandomNoDuplicates (iMin, iMax, iElements) 
	        If (iMax-iMin+1)>iElements Then
		        Exit Function
	        End If 
	        Dim RndArr(), x, curRand
	        Dim iCount, arrValues()
	 
	        Redim arrValues(iMax-iMin)
	        For x=iMin To iMax
		        arrValues(x-iMin) = x
	        Next
	 
	        Redim RndArr(iElements-1)
	 
	        For x=0 To UBound(RndArr)
		        RndArr(x) = iMin-1
	        Next 
	        Randomize
	        iCount=0 
	        Do Until iCount>=iElements 
		        curRand = arrValues(CLng((Rnd*(iElements-1))+1)-1) 
 		        If Not(InArray(RndArr, curRand)) Then
			        RndArr(iCount)=curRand
			        iCount=iCount+1
		        End If 
		        If Not(Response.IsClientConnected) Then
			        Exit Function
		        End If
	        Loop 
	        RandomNoDuplicates = RndArr
        End Function
  
        Function InArray(arr, val)
	        Dim x
	        InArray=True
	        For x=0 To UBound(arr)
		        If arr(x)=val Then
			        Exit Function
		        End If
 	        Next
	        InArray=False
        End Function

        Function CopyArray (arr)
	        Dim result(), x
	        ReDim result(UBound(arr))
	        For x=0 To UBound(arr)
		        If IsObject(arr(x)) Then
			        Set result(x) = arr(x)
		        Else  
			        result(x) = arr(x)
		        End If
	        Next
	        CopyArray = result
        End Function
         
     ''end function Random
  
    Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")    
    objXMLDoc.async = False    
    objXMLDoc.load Server.MapPath("webpartnerBannerAdmin.xml")
   

    function getNodeCount(banner,banType)
       count = 0
        For Each xmlProduct In objXMLDoc.documentElement.selectNodes("/Table/Admin_webpartnerBanner[bannerSize=" & banner &" and bannerType='" & banType &"' and numberDisplay > 0]") 
            Dim numberDisplay : numberDisplay = xmlProduct.selectSingleNode("numberDisplay").text   
            count= count + cint(Server.HTMLEncode(numberDisplay))
         Next 
         getNodeCount = count
    end function
     
         
      hostname =  Request.ServerVariables("server_name") 
      redirectUrl ="http://www.silkspan.com/banner/log_cbanner_counter.asp?fname="
     ' redirectUrl ="http://192.168.0.2/banner/log_cbanner_counter.asp?fname="
 
       strdivValue =""
       if (len(Request("bSize"))>0) then  
                banner = Request("bSize")
                banType = Request("bType")
                  
               Dim arr() 
                ReDim arr(getNodeCount(banner,banType))
                i = 0 
                cc= 0
                
                wid = 0
                heig =0
                Dim xmlProduct       
                For Each xmlProduct In objXMLDoc.documentElement.selectNodes("/Table/Admin_webpartnerBanner[bannerSize=" & banner &" and bannerType='" & banType &"' and numberDisplay > 0]")
                     Dim bannerSize : bannerSize = xmlProduct.selectSingleNode("bannerSize").text   
                     Dim bannerType : bannerType = xmlProduct.selectSingleNode("bannerType").text   
 
                     Dim filePath : filePath = xmlProduct.selectSingleNode("filePath").text   
                     Dim numberDisplay : numberDisplay = xmlProduct.selectSingleNode("numberDisplay").text   
                     Dim url : url = xmlProduct.selectSingleNode("url").text   
           
                    
                     Dim width : width = xmlProduct.selectSingleNode("width").text   
                     Dim height : height = xmlProduct.selectSingleNode("height").text   
                      

                    strFilePath = filePath
                    floderName = "/banner/partner/IF_Banner/PathImage/"&  Request("btype")
                    dim fs,p
                    set fs=Server.CreateObject("Scripting.FileSystemObject")
                    path=fs.getfilename(strFilePath) 

                     Response.Write("<!-- xx: " & filePath  & ": " & InStr( filePath,"www2.silkspan.com") & "-->")
                    if(InStr( filePath,"www2.silkspan.com") > 0 ) then
                       filePath =  filePath   
                    else
                         filePath = floderName &  "\" & path   'Server.MapPath(floderName) &  "\" & path  
                    end if

                  

                     'if( banner =  Server.HTMLEncode(bannerSize) and  banType =  Server.HTMLEncode(bannerType) ) then  
                             wid = Server.HTMLEncode(width) 
                             heig = Server.HTMLEncode(height)  

                            FOR  a = 1 to cint(Server.HTMLEncode(numberDisplay))     
                              '''''''check other typedealr
                              oldUrl = url
                                 leQ = InStr(url,"?")
                                 adsTypeDeal = "" 
                                 if(leQ > 0 ) then
                                    adsTypeDeal ="&" &  Mid(url, leQ+1, len(url)-leQ+1)
                                    url = Mid(url, 1, leQ-1)
                                 end if 
                             '' end check other typedeler       
                             ' Response.Write(url &"<br>")

                               if banType = "credit" then  
                                  arr(i) =  "<div class='model a'><a href='" & redirectUrl & url & "&typedealer=" & Request("parner") & adsTypeDeal & "' target='_blank' ><img class='cImg'  data-number='" & i & "' src='" & filePath & "'></a></div>"
                               else
                                   arr(i) = "<div class='model a'><a href='" & url & "?typedealer=" &  Request("parner") & adsTypeDeal &  "' target='_blank' ><img class='cImg'  data-number='" & i & "' src='" & filePath & "'></a></div>"
                               end if 
                              url= oldUrl 
                                i=i+1 
                            NEXT   
                    ' END if 
                Next   

               arrc =  Shuffle(arr)
               Response.Write("<div><div class='banner' >")
                for i=0 to uBound(arrc) 
                   Response.Write  arrc(i) 
                Next 
              Response.Write("</div></div>")
       

               strSript ="" 
               strSript = strSript &"<script>" 
               strSript = strSript &"$(function () {  "
               strSript = strSript &"document.getElementById('hdfwidth').value='" & wid & "px';  "
               strSript = strSript &"document.getElementById('hdfheight').value='" & heig & "px';  "
               strSript = strSript &"document.getElementById('hdfSpeed').value='" & 6  & "000';  "
               strSript = strSript &"}); "
               strSript = strSript &"</script> "
               Response.Write(strSript) 



       else 
 
          Response.Write("null")
       end if 
 
  
  %>

   <script type="text/javascript">
           $(function () {
               var speed = document.getElementById('hdfSpeed').value;
               $(".banner > div:gt(0)").hide();
               setInterval(function () { $('.banner > div:first').fadeOut(900).next().fadeIn(900).end().appendTo('.banner'); }, speed);
               setSize();
           });
           function shuffleArray(array) {
               for (var i = array.length - 1; i > 0; i--) {
                   var j = Math.floor(Math.random() * (i + 1));
                   var temp = array[i];
                   array[i] = array[j];
                   array[j] = temp;
               }
               return array;
           }

           function setSize() {
               var height = document.getElementById('hdfheight').value;
               var width = document.getElementById('hdfwidth').value;
               $('.cImg').css('width', width);
               $('.cImg').css('height', height);
               $('.banner').css('width', width);
               $('.banner').css('height', height);
           }
</script>
 
    <input id="hdfheight" type="hidden" />
    <input id="hdfwidth" type="hidden" />
    <input id="hdfSpeed" type="hidden" />
 

     
   