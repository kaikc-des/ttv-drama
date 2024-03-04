<!--#include virtual="/include/inc/killChars.inc"-->
<!--#Include Virtual="/include/inc/HomeV2ADO.inc" -->
<!--#Include Virtual="/include/inc/RemoveHtml.inc" -->
<!--#Include Virtual="/include/inc/adovbs.inc" -->
<%
'必須修改項目
AppName="MongaWoman_NewsView"
PgmName="艋舺的女人"

InfoID = killChars(Request.QueryString("id"))

On Error Resume Next
InfoID = int(InfoID)
if err.number <> 0 then
	Err.Clear
	Response.End
end if	
On Error GoTo 0

RemoveMemCached=LCase(request("Action"))
if RemoveMemCached="removememcached" then
    call RemoveString(AppName,InfoID)
	Response.end
end if

setRV = -100
removeRV = -100
tmpStr = ""
xmlString = ""

xmlString = GetDataString(InfoID, "GetData")
If setRV = 200 Then
    tmpStr = "Cached"
Else
    tmpStr = "DB"
End If
'response.Write "<!--data from : "+ tmpStr +xmlString+"-->"

set oXML = Server.CreateObject("Microsoft.XMLDOM")
oXML.async = False
oXML.loadXML(XmlString)
If oXML.parseError.errorCode <> 0 Then
    response.Write "<!--Error Reason : "+ oXML.parseError.reason + "-->"
    response.Write "<!--Error Line : "+ cstr(oXML.parseError.line) + "-->"
	Response.End
End If

Set oRC = oXML.selectNodes("/DataList/RecordCount")
RecordCount = 0
if oRC.length > 0 then
    for each x in oRC
        RecordCount = x.selectSingleNode("rc").text
    next
end if
set oRC = nothing

if RecordCount <= 0 then
    response.Redirect "https://www.ttv.com.tw/"
end if
set oData = nothing

Function GetDataString(key, getFunc)
    Dim o, val
    Set o = Server.CreateObject("MemcachedCOM.MemCache")
    val = o.GetString(AppName , key)
    If IsEmpty(val)  Then
	    Dim afunc
	    Set afunc = GetRef(getFunc)
	    val = afunc()
	    setRV = o.SetString(AppName, key, val)
    Else
        setRV = 200
    End If
    Set o = Nothing
    GetDataString = val
End Function

Function GetData()
	tmpStr= "<?xml version='1.0' encoding='UTF-8'?>"
	tmpStr=tmpStr& "<DataList>"

	OpenHomeV2 dbcon
	sqlstr = "Select * From vTTVInfoDept where datediff(d,BeginDate,getdate()) >= 0 and datediff(d,getdate(),ExpireDate) >= 0 and ShowOn='1' and ID ="&InfoID
	Set RsInfo = dbcon.Execute(sqlstr)
    tmpStr=tmpStr& "<RecordCount><rc>"& RsInfo.recordcount &"</rc></RecordCount>"
	if not rsInfo.eof then
'	    set rsimg = dbcon.execute("select * from tblTTVInfoImg where imode='1' and ID="+ cstr(RsInfo("ID")) +" order by dimention desc, PID desc")
'	    if not rsimg.eof then
'	        image_src = "https://img.ttv.com.tw/ttvinfoimg/" + RemoveBraces(rsimg("uniqueID")) + "-800x600.jpg"
'       end if
'        set rsimg = nothing
	    tmpStr=tmpStr& "<DataInfo>"
	    tmpStr=tmpStr& "<InfoID>" & RsInfo("ID") & "</InfoID>"
   	    tmpStr=tmpStr& "<Heading><![CDATA[" & replace(rsInfo("Heading"),"《"+PgmName+"》","") & "]]></Heading>"
	    tmpStr=tmpStr& "<Description><![CDATA[" & RsInfo("Text") & "]]></Description>"
	    tmpStr=tmpStr& "<CreateDate><![CDATA[" & formatdatetime(RsInfo("CreateDate"),2) & "]]></CreateDate>"
        tmpStr=tmpStr& "<Photo>"
	    set rsimg = dbcon.execute("select * from tblTTVInfoImg where imode='1' and ID="+ cstr(infoID) +" order by PID desc")
	    do while not rsimg.eof
	        imgSrc = "https://img.ttv.com.tw/ttvinfoimg/" + RemoveBraces(rsimg("uniqueID")) + "-800x600.jpg"
	        rsImg.movenext
	        if rsImg.eof then
	            tmpStr=tmpStr & imgSrc
            else		        
		        tmpStr=tmpStr & imgSrc & "+++"
            end if			    
        loop		    
        tmpStr=tmpStr & "</Photo>"
	    rsImg.close
	    set rsImg=nothing

	    tmpStr=tmpStr& "</DataInfo>"
	    RsInfo.movenext
	end if
	RsInfo.close
	set RsInfo=nothing
	dbcon.close
	set dbcon=nothing
	tmpStr=tmpStr& "</DataList>"
	GetData = tmpStr
End Function

Sub RemoveString(Src,key)
    Dim o
    Set o = Server.CreateObject("MemcachedCOM.MemCache")
    removeRV = o.RemoveString(Src, key)
    Set o = Nothing
End Sub

if RecordCount > 0 then
    Set oData = oXML.selectNodes("/DataList/DataInfo")    
    for each x in oData
        InfoID=int(x.selectSingleNode("InfoID").text)
        Heading=x.selectSingleNode("Heading").text
        Photo=x.selectSingleNode("Photo").text
        CreateDate=x.selectSingleNode("CreateDate").text
        Description=x.selectSingleNode("Description").text
        PhotoAry = split(Photo,"+++")
        Photo0=""
        if ubound(PhotoAry) > 0 then
            Photo0 = PhotoAry(0)
        end if   
        htm_url = "http://www.ttv.com.tw/drama14/MongaWoman/newsview.asp?id="& cstr(InfoID)
    next
end if
set oData = nothing
'set oXML = nothing

function RemoveBraces(Bstr)
	RemoveBraces = replace(replace(Bstr,"{",""),"}","")
end function
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="author" content="TTV">
    <meta name="keywords" content="艋舺的女人,狄鶯,李興文,陳仙梅,黃仲崑,傅天穎,安定亞,台視" />
    <meta name="Description" content="<%=left(Description,100) %>.." />
    <title><%=Heading%> - 新聞 - 艋舺的女人</title>
    <meta property="og:title" content="<%=Heading%> - 新聞 - 艋舺的女人" />
    <meta property="og:type" content="tv_show" />
    <meta property="og:image" content="<%=Photo0 %>"/>
    <meta property="og:url" content="<%=htm_url %>" />
    <meta property="og:site_name" content="台視" />
    <meta property="fb:app_id" content="173424616047838" />
    <meta property="og:description" content="<%=left(Description,100) %>.." />

    <!-- Bootstrap -->
    <link rel="stylesheet" href="/drama14/css/bootstrap.min.css">
    <link type="text/css" media="screen" rel="stylesheet" href="/include/js/colorbox-master/example5/colorbox.css" />      
    <link rel="stylesheet" href="/drama14/css/bootstrap.min.css" type="text/css" media="screen" />
    <style type="text/css">
        .HeaderTitle {
            text-align:center;
        }    
        .br
        {
            CLEAR: both;
            PADDING-TOP: 5px;
        }
        .br2x
        {
            CLEAR: both;
            PADDING-TOP: 10px;
        }
        .pdb
        {
            PADDING-BOTTOM: 15px;
        }
        
  
.gold {color:#FFD700;}
    
#Photo {
	margin-left:0px;
}

#Photo a {
	outline: none;
}

#Photo a img {

	padding-top: 11px;
	vertical-align: top;
}

#PhotoMore {
	margin-left:0px;
}


#PhotoMore a {
	outline: none;
    padding-top: 5px;
}

#PhotoMore a img {

	padding-top: 11px;
	vertical-align: top;
}
#Photo a img.last {
	margin-right: 0;	
}
    
.ZoonIn {
    background-image:url("/taiwan/images/ZoonIn.png");
    background-repeat: no-repeat;
    background-position:center 0px;

	height:90px;
	border: 1px solid #BBB;
    padding:0 5px 0 5px;
	float:left;
	margin:5px 5px 0 0;
	text-align:center;
}
    
.br
{
    CLEAR: both;
    PADDING-TOP: 5px
}
.br2x
{
    CLEAR: both;
    PADDING-TOP: 10px;
}
    
    
    </style>

    <link href="2014.css" rel="stylesheet" />

    <!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->
<!-- Bootstrap core JavaScript
================================================== -->
<!-- Placed at the end of the document so the pages load faster -->
<script src="https://code.jquery.com/jquery-1.10.2.min.js"></script>
<script src="/include/js/bootstrap-3.0.0/js/bootstrap.min.js"></script>
<script type="text/javascript" src="/include/js/jquery.autopager-1.0.0.min.js"></script>
<script type="text/javascript" src="/include/js/colorbox-master/jquery.colorbox-min.js"></script>
<script type="text/javascript" src="/drama14/js/mosaic.1.0.1.min.js"></script>
<script type="text/javascript" src="https://apis.google.com/js/plusone.js">
  {"parsetags": "explicit"}
</script>
</head>

<body>
<div id="fb-root"></div>
<script>(function(d, s, id) {
  var js, fjs = d.getElementsByTagName(s)[0];
  if (d.getElementById(id)) return;
  js = d.createElement(s); js.id = id;
  js.src = "//connect.facebook.net/zh_TW/all.js#xfbml=1&appId=173424616047838";
  fjs.parentNode.insertBefore(js, fjs);
}(document, 'script', 'facebook-jssdk'));</script>
<!--#include file="top.asp"-->

    <div class="container">
        <div class="row" id="pagerow" style="margin-top:2em">

            <div class="col-md-8 col-lg-8">

                <div class="panel panel-default" style="margin-bottom:10px;">
                    <div class="panel-body">
                        <div id="panel-body">
                            <div class="row" style="margin:0 auto;padding:0em 2em" id="container">

                                <div id="Heading"><h3 style="color:#1144bb;line-height:180%;font-weight:800"><%=Heading%></h3></div>
                                <div style="padding:5px 10px 0 10px;">發佈日期：<%=CreateDate%></div>
                                <div class="br"></div>
    			                <div style="margin:10px 0 10px 10px;width:640px;">
    			                    <div style="float: left;" id="gplusbtn"><g:plusone size="medium" href="<%=htm_url%>"></g:plusone></div>
    			                    <div style="float: left;margin-top:-4px;" id="fblikebtn">
                                        <fb:like href="<%=htm_url%>" send="true" layout="button_count" width="100" show_faces="false" font="arial"></fb:like>
                                    </div>
			                    </div>
                                <div class="br"></div>

                                <div id="Photo">
                                    <center>
                        <%
                                    On Error Resume Next
                                    for i = 0 to ubound(PhotoAry)
                                        if err.number <> 0 then
                                        Err.Clear
                                        Response.Redirect htm_url + "&action=removememcached"
                                        end if%>		
                                        <div class="ZoonIn"><a class="group1" href="<%=PhotoAry(i)%>"><img src="<%=replace(PhotoAry(i),"800x600","72x72") %>" <%if i = ubound(PhotoAry) then response.write "class='last'" end if%> /></a></div>
                            <%      next %>

                                    </center>
                                </div>
                                <div class="br"></div>
                                <div><%=Description %></div>
                            </div><!--/row end -->
                                              
                        </div>
                        <div style="padding-top:20px;text-align:center" ><a class="btn btn-sm btn-default" href="javascript:history.back()">回上頁 back</a></div>
                    </div>
			    </div> <!--/panel end -->
			    
            </div>
            
            <div class="col-md-4 col-lg-4 ">
                <div>
<!-- 回應式廣告單元 -->
<ins class="adsbygoogle"
     style="display:block"
     data-ad-client="ca-pub-5342396684826977"
     data-ad-slot="3553673898"
     data-ad-format="auto"></ins>
<script>
(adsbygoogle = window.adsbygoogle || []).push({});
</script>
                </div>
                <div class="br2x" >
<!-- 回應式廣告單元 -->
<ins class="adsbygoogle"
     style="display:block"
     data-ad-client="ca-pub-5342396684826977"
     data-ad-slot="3553673898"
     data-ad-format="auto"></ins>
<script>
(adsbygoogle = window.adsbygoogle || []).push({});
</script>
                </div>
                <div class="br2x" >
<!-- 回應式廣告單元 -->
<ins class="adsbygoogle"
     style="display:block"
     data-ad-client="ca-pub-5342396684826977"
     data-ad-slot="3553673898"
     data-ad-format="auto"></ins>
<script>
(adsbygoogle = window.adsbygoogle || []).push({});
</script>
                </div>
                
<!-- 眾多回應式廣告單元，本行只需下一次即可 -->
<script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
<!-- 眾多回應式廣告單元，本行只需下一次即可 -->

            </div>             
        </div>
        <!-- /row of columns -->

    </div> <!-- /container -->
    
<!-- FOOTER -->
<div id="footer">
    <div class="container">
        <!--#include file="under.htm"-->
    </div>
</div>

<script type="text/javascript">
    $(document).ready(function() {
        $(".group1").colorbox({rel:'group1'});
    });	

    function PageHit(pageid){
        $.ajax({
            type: "GET",
            crossDomain: true,
            url: "/AD/PageHits.asp",
            data: { PgID: pageid},
            cache: false
        });
    }
    
    gapi.plusone.go();        
    PageHit('1111-563');

    $('#news').addClass("active");
</script>    

    <script type="text/javascript" src="https://apis.google.com/js/platform.js">
        { lang: 'zh-TW' }
    </script>
    <!--#include virtual="/group/js/ga.js"-->
  </body>
</html>
