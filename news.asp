<!--#include virtual="/include/inc/adovbs.inc" -->  
<!--#include virtual="/include/inc/killChars.inc"-->
<!--#include virtual="/include/inc/HomeV2ADO.inc"-->
<!--#include virtual="/include/inc/HomeV3ADO.inc"-->
<!--#Include Virtual="/Include/inc/RemoveHTML.inc" -->
<%
'必須修改項目
AppName="MongaWoman_News"
PgmName="艋舺的女人"
imageNotExist = "/drama14/MongaWoman/images/index.jpg"   '圖片不存在時，顯示此圖

page = KillChars(Request.QueryString("page"))

if page = "" then
	pageindex = 0
	page = 1
else
	On Error Resume Next
	pageindex = page - 1
	if Err.number <> 0 then
		Err.Clear
		Response.End
	end if	
	On Error GoTo 0
end if

pagesize = 6
AppName= AppName + replace(cstr(date()),"/","") + cstr(hour(now))
RemoveMemCached=LCase(request("Action"))
if RemoveMemCached="removememcached" then
    for page = 1 to 100
    call RemoveString(AppName,page)
    next
	Response.end
end if

setRV = -100
removeRV = -100
tmpStr = ""
xmlString = ""

xmlString = GetDataString(page, "GetData")
If setRV = 200 Then
    tmpStr = "Cached"
Else
    tmpStr = "DB"
End If

'response.Write "<!--data from : "+ tmpStr +"-->"
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
        PageCount = x.selectSingleNode("pagecount").text
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

'    query=" ShowOn='1' and datediff(d,Createdate,getdate()) < 90 and (Heading like '%《艋舺的女人》%')"
    query=" ShowOn='1' and (Heading like '%《"+PgmName+"》%' )"
    OpenHomeV2 dbcon
    sql = "SELECT CEILING(count(ID) / "+ cstr(pagesize) +".0) as rcount FROM tblTTVInfo where " & query
    set rs=dbcon.execute(sql)
    pagecount = rs("rcount")
    sql="select TOP "+ cstr(pagesize) +" * from tblTTVInfo where ID NOT IN (SELECT TOP "+ cstr(pagesize * pageindex) +" ID FROM tblTTVInfo where "+ query +" order by ID desc) and "+ query +" order by ID desc"
    set rs=dbcon.execute(sql)
    
    tmpStr=tmpStr& "<RecordCount><rc>"& rs.recordcount &"</rc><pagecount>"& cstr(pagecount) &"</pagecount></RecordCount>"
	
	do while not rs.eof
	    set rsimg = dbcon.execute("select * from tblTTVInfoImg where imode='1' and ID="+ cstr(rs("ID")) +" order by dimention desc, PID desc")
	    image_src = imageNotExist
	    if not rsimg.eof then
	        image_src = "https://img.ttv.com.tw/ttvinfoimg/" + RemoveBraces(rsimg("uniqueID")) + "-800x600.jpg"
        end if
        set rsimg = nothing
	    tmpStr=tmpStr& "<DataInfo>"
	    tmpStr=tmpStr& "<ID>" & rs("ID") & "</ID>"
	    tmpStr=tmpStr& "<Heading><![CDATA[" & replace(rs("Heading"),"《"+PgmName+"》","") & "]]></Heading>"
	    tmpStr=tmpStr& "<Description><![CDATA[" & rs("Text") & "]]></Description>"
	    tmpStr=tmpStr& "<Photo><![CDATA[" & image_src & "]]></Photo>"
	    tmpStr=tmpStr& "<CreateDate><![CDATA[" & formatdatetime(rs("CreateDate"),2) & "]]></CreateDate>"
	    tmpStr=tmpStr& "</DataInfo>"
	    rs.movenext
	loop
	rs.close
	set rs=nothing
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

sub ShowData()
    if RecordCount > 0 then
        Set oData = oXML.selectNodes("/DataList/DataInfo")    
        y = 0
        for each x in oData
            y = y + 1
            ID=x.selectSingleNode("ID").text
            Heading=x.selectSingleNode("Heading").text
            imgSrc=x.selectSingleNode("Photo").text
            htm_url="newsview.asp?id=" + ID
'   定義 臉書「讚」的連結位置
            fb_url="http://www.ttv.com.tw/drama14/MongaWoman/" + htm_url
    %>

            <div class="col-sm-6 col-sm-6 col-md-6 col-lg-6 pdb">
                <div class="mosaic-block bar2">
			        <a href="<%=htm_url %>" target="_top" class="mosaic-overlay">
				        <div class="details">
					        <p class="gold"><%=Heading%></p>
				        </div>
			        </a>
                    <div class="mosaic-backdrop"><img src="<%=imgSrc%>" class="img-thumbnail img-responsive"  alt="<%=Heading%>" /></div>
                </div>

                <div style="padding:0 0 0 10px;">
                    <div style="float: left;width:80px;margin-top:5px;"><g:plusone size="medium" href="<%=fb_url%>"></g:plusone></div>
                    <div style="float: left;margin-top:5px;">
                        <iframe src="https://www.facebook.com/plugins/like.php?href=<%=fb_url%>&amp;layout=button_count&amp;show_faces=false&amp;width=100&amp;action=like&amp;colorscheme=light&amp;height=20" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:100px; height:20px;" allowTransparency="true"></iframe>
                    </div>
                </div>
            </div>
    <%            
        next
    else
    response.Redirect "https://www.ttv.com.tw/"
    end if
    set oData = nothing

end sub

function RemoveBraces(Bstr)
	RemoveBraces = replace(replace(Bstr,"{",""),"}","")
end function
%>
<!DOCTYPE html>
<html lang="zh-tw">
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="author" content="TTV">
    <meta name="keywords" content="艋舺的女人,狄鶯,李興文,陳仙梅,黃仲崑,傅天穎,安定亞,台視" />
    <meta name="Description" content="艋舺的女人，台視優質戲劇。一段發生在保守的時代，想愛卻無法愛的故事。兩個失去至愛，必須相濡以沫、共組家庭的女人；雖然她們沒有丈夫，但卻擁有彼此。" />
    <title>艋舺的女人 - 新聞</title>
    <meta property="og:title" content="艋舺的女人 - 新聞" />
    <meta property="og:type" content="tv_show" />
    <meta property="og:image" content="https://www.ttv.com.tw/drama14/MongaWoman/images/200x200_www.jpg"/>
    <meta property="og:url" content="https://www.ttv.com.tw/drama14/MongaWoman/index.htm" />
    <meta property="og:site_name" content="台視" />
    <meta property="fb:app_id" content="173424616047838" />
    <meta property="og:description" content="艋舺的女人，台視優質戲劇。一段發生在保守的時代，想愛卻無法愛的故事。兩個失去至愛，必須相濡以沫、共組家庭的女人；雖然她們沒有丈夫，但卻擁有彼此。" />

    <!-- Bootstrap -->
    <link rel="stylesheet" href="/drama14/css/bootstrap.min.css" />
    <link rel="stylesheet" href="/drama14/css/mosaic.css" type="text/css" media="screen" />
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
        
    /*General Mosaic Styles*/
    .mosaic-block {
	    float:left;
	    position:relative;
	    overflow:hidden;
	    width:100%;
	    height:250px;
	    margin:10px;
	    background:#111 url(/group/images/ajax-loader.gif) no-repeat center center;
	    border:1px solid #fff;
	    -webkit-box-shadow:0 1px 3px rgba(0,0,0,0.5);
    }
        
	.mosaic-backdrop {
		display:none;
		position:absolute;
		top:0;
		height:100%;
		width:100%;
		background:#111;
	}
        
        
	.bar2 .mosaic-overlay {
		bottom:-25px;
		height:75px;
		opacity:0.8;
		-ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=80)";
		filter:alpha(opacity=80);
			padding:0 5px 5px 5px;
	}
	
	.bar2 .mosaic-overlay:hover {
		opacity:1;
		-ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=100)";
		filter:alpha(opacity=100);
	}
    
    .gold {color:#FFD700;}
    
        .r2p {paddig-right:2px;}
        .ellipsis{
            width:100%;
            overflow : hidden;
            text-overflow : ellipsis;
            white-space : nowrap;
        }    
        .ad300x2 {margin:0 auto;max-width:620px;}            
    
    </style>

    <!-- Custom styles for this template -->
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
<!--<script src="//netdna.bootstrapcdn.com/bootstrap/3.1.1/js/bootstrap.min.js"></script>-->
<script src="/include/js/bootstrap-3.0.0/js/bootstrap.min.js"></script>
<script type="text/javascript" src="/include/js/jquery.autopager-1.0.0.min.js"></script>
<script type="text/javascript" src="/include/js/colorbox-master/jquery.colorbox-min.js"></script>
<script type="text/javascript" src="/drama14/js/mosaic.1.0.1.min.js"></script>
<script type="text/javascript" src="https://apis.google.com/js/plusone.js">
  {"parsetags": "explicit"}
</script>
</head>
<body>
<!--#include file="top.asp"-->

    <div class="container">
        <div class="row" id="pagerow" style="margin-top:2em">

            <div class="col-xs-12 col-sm-12 col-md-8 col-lg-8">
                <div class="panel panel-default" style="margin-bottom:10px;">
                    <div class="panel-body">
                        <div id="panel-body">
                            <div class="row" style="margin:0 auto;" id="container">
                            <% ShowData() %>  
                            </div><!--/row end -->                  
                        </div>
                    </div>
			    </div> <!--/panel end -->
                <div><!-- pagination start -->
                    <%
                    if page = 1 then
                        prevPage = 1
                    else
                        prevPage = page - 1
                    end if

                    if cint(page) = cint(PageCount) then
                        nextPage = PageCount
                    else
                        nextPage = page + 1
                    end if

                    if page - 5 <= 1 then
                        Pstart = 1
                    else
                        Pstart = page - 5
                    end if        

                    if page + 5 >= cint(PageCount) then
                        Pend = cint(PageCount)
                    else
                        Pend = page + 5
                    end if        

                    function PageClass(page,currentPage,status)
                        if cint(page) = cint(currentPage) then
                            PageClass = status
                        end if        
                    end function
                     %>			    
                     <center>
                    <ul class="pagination">
                        <li class="<%=PageClass(page,1,"disabled") %>"><a href="news.asp?page=1" data-toggle="tooltip" data-placement="top" title="跳至第一頁"><span class="glyphicon glyphicon-fast-backward"></span></a></li>
                        <li class="<%=PageClass(page,1,"disabled") %>"><a href="news.asp?page=<%=prevPage %>" data-toggle="tooltip" data-placement="top" title="跳至第<%=prevPage %>頁"><span class="glyphicon glyphicon-backward"></span></a></li>
                        <%for Ploop = Pstart to Pend %>
                        <li class="<%=PageClass(page,Ploop,"active") %>"><a href="news.asp?page=<%=Ploop %>" data-toggle="tooltip" data-placement="top" title="跳至第<%=Ploop %>頁"><%=Ploop %></a></li>
                        <%next %>
                        <li class="<%=PageClass(page,PageCount,"disabled") %>"><a href="news.asp?page=<%=nextPage %>" data-toggle="tooltip" data-placement="top" title="跳至第<%=nextPage %>頁"><span class="glyphicon glyphicon-forward"></span></a></li>
                        <li class="<%=PageClass(page,PageCount,"disabled") %>"><a href="news.asp?page=<%=PageCount %>" data-toggle="tooltip" data-placement="top" title="跳至最後一頁"><span class="glyphicon glyphicon-fast-forward"></span></a></li>
                    </ul>			    
                    </center>			    
                </div><!-- pagination end -->
            </div>
            
            <div class="clearfix hidden-sm hidden-xs col-md-4 col-lg-4">
                <div>
            <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
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
            <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
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
            <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
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
            </div>
            
        </div>
        <div style="padding-top:10px;"></div><!-- /row of columns -->

    </div> <!-- /container -->
    
<!-- FOOTER -->
<div id="footer">
    <div class="container">
        <!--#include file="under.htm"-->
    </div>
</div>

<script type="text/javascript">
    $(document).ready(function() {
        var page = 1;
        if(page >= <%=PageCount %>) $("#next").css("display","none");	

		$('.bar2').mosaic({
			animation	:	'slide'		//fade or slide
		});
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
