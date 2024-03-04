<!--#include virtual="/include/inc/HomeV2ADO.inc"-->
<!--#include virtual="/include/inc/KillChars.inc"-->
<%
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

ID = KillChars(request.QueryString("id"))
AppName="PrinceWilliamFBPagesPost"
if ID = "" then
    AppName = AppName + replace(cstr(date()),"/","") + cstr(hour(now))
end if

RemoveMemCached=LCase(request("Action"))
if RemoveMemCached="removememcached" then
'	call RemoveString(AppName,0)
	if ID = "" then
	    for  page = 0 to 30
    	    call RemoveString(AppName,page)
	    next
	else
	    call RemoveString(AppName,ID)
	end if
	Response.end
end if

setRV = -100
removeRV = -100
tmpStr = ""
xmlString = ""

if ID = "" then
    xmlString = GetDataString(page, "GetData")
else
    xmlString = GetDataString(ID, "GetData1")
end if    
If setRV = 200 Then
    tmpStr = "Cached"
Else
    tmpStr = "DB"
End If
'response.Write "<!--data from : "+tmpStr+ "-->"

set oXML = Server.CreateObject("Microsoft.XMLDOM")
oXML.async = False
oXML.loadXML(XmlString)
If oXML.parseError.errorCode <> 0 Then
    response.Write "<!--Error Reason : "+ oXML.parseError.reason + "-->"
    response.Write "<!--Error Line : "+ oXML.parseError.line + "-->"
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

if RecordCount > 0 then
    Set oRC = oXML.selectNodes("/DataList/FirstRecord")
    if oRC.length > 0 then
        for each x in oRC
            FirstID = x.selectSingleNode("ID").text
        next
    end if
    set oRC = nothing
    if ID <> "" then
        found = false
        Set oData = oXML.selectNodes("/DataList/DataInfo")    
        for each x in oData
            PostID=x.selectSingleNode("PostID").text
            if PostID = ID then              
                PageName=x.selectSingleNode("PageName").text
                Message=x.selectSingleNode("Message").text
                Picture=x.selectSingleNode("Picture").text
                found = true
                FirstID = ID
                exit for   
            end if
        next
    end if                
else
    response.Redirect "https://www.ttv.com.tw/"
end if
set oData = nothing
'set oXML = nothing

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
	
    pagesize = 12
    
    query = "Post_from_id='602181683203965' and PageShowOn = '1' and ShowOn = '1' and TTVShowUrl <> ''"
    sql = "Select CEILING(count(post_id) / "+ cstr(pagesize) +".0) as rcount FROM vFBPagePost WHERE "& query
    'response.Write sql + "<br>"
    set rs=dbcon.execute(sql)
    pagecount = rs("rcount")
'    response.Write pagecount
    sql = "Select TOP "+ cstr(pagesize) +" * FROM vFBPagePost where post_id NOT IN (SELECT TOP "+ cstr(pagesize * pageindex) +" post_id FROM vFBPagePost where "+ query +"  ORDER BY post_created_time desc) and "+ query +" order by post_created_time desc"
'    response.Write sql
    set rs=dbcon.execute(sql)
	
	
'    sql = "Select top 30 * from vFBPagePost Where ShowOn = '1' order by post_created_time desc"
'	set rs = dbcon.execute(sql)
	tmpStr=tmpStr& "<RecordCount><rc>"& rs.recordcount &"</rc></RecordCount>"
    if not rs.eof then
	    tmpStr=tmpStr& "<FirstRecord><ID>"& rs("post_id") &"</ID></FirstRecord>"
    end if
	while not rs.eof
		post_id = rs("post_id")
		post_message = replace(rs("post_message"),"&","&amp;")
		post_picture = replace(rs("post_picture"),"&","&amp;")
		post_icon = replace(rs("post_icon"),"&","&amp;")
		post_link = replace(rs("post_link"),"&","&amp;")
		post_created_time = rs("post_created_time")
		post_type = rs("post_type")
		page_name = rs("page_name")
		page_id = rs("post_from_id")
		TTVShowUrl = rs("TTVShowUrl")
	
		tmpStr=tmpStr& "<DataInfo>"
		tmpStr=tmpStr& "<PostID>" & post_id & "</PostID>"
		tmpStr=tmpStr& "<PostType>" & post_type & "</PostType>"
		tmpStr=tmpStr& "<Link>" & post_link & "</Link>"
		tmpStr=tmpStr& "<Picture>" & post_picture & "</Picture>"
		tmpStr=tmpStr& "<Icon>" & post_icon & "</Icon>"
		tmpStr=tmpStr& "<Message><![CDATA[" & post_message & "]]></Message>"
		tmpStr=tmpStr& "<CreatedTime>" & post_created_time & "</CreatedTime>"
		tmpStr=tmpStr& "<PageName>" & page_name & "</PageName>"
		tmpStr=tmpStr& "<PageID>" & page_id & "</PageID>"
		tmpStr=tmpStr& "<TTVShowUrl>" & TTVShowUrl & "</TTVShowUrl>"
		tmpStr=tmpStr& "</DataInfo>"
		rs.movenext
	wend
	dbcon.close
	set dbcon=nothing
	tmpStr=tmpStr& "</DataList>"
	GetData = tmpStr
End Function

Function GetData1()
	tmpStr= "<?xml version='1.0' encoding='UTF-8'?>"
	tmpStr=tmpStr& "<DataList>"
	OpenHomeV2 dbcon
	
    query = "PageShowOn = '1' and ShowOn = '1' and TTVShowUrl <> ''"
    sql = "Select * FROM vFBPagePost where "+ query +" and post_id='"+ ID +"'"
'    response.Write sql
    set rs=dbcon.execute(sql)
	
	
'    sql = "Select top 30 * from vFBPagePost Where ShowOn = '1' order by post_created_time desc"
'	set rs = dbcon.execute(sql)
	tmpStr=tmpStr& "<RecordCount><rc>"& rs.recordcount &"</rc></RecordCount>"
    if not rs.eof then
	    tmpStr=tmpStr& "<FirstRecord><ID>"& rs("post_id") &"</ID></FirstRecord>"
    end if
	while not rs.eof
		post_id = rs("post_id")
		post_message = replace(rs("post_message"),"&","&amp;")
		post_picture = replace(rs("post_picture"),"&","&amp;")
		post_icon = replace(rs("post_icon"),"&","&amp;")
		post_link = replace(rs("post_link"),"&","&amp;")
		post_created_time = rs("post_created_time")
		post_type = rs("post_type")
		page_name = rs("page_name")
		page_id = rs("post_from_id")
		TTVShowUrl = rs("TTVShowUrl")
	
		tmpStr=tmpStr& "<DataInfo>"
		tmpStr=tmpStr& "<PostID>" & post_id & "</PostID>"
		tmpStr=tmpStr& "<PostType>" & post_type & "</PostType>"
		tmpStr=tmpStr& "<Link>" & post_link & "</Link>"
		tmpStr=tmpStr& "<Picture>" & post_picture & "</Picture>"
		tmpStr=tmpStr& "<Icon>" & post_icon & "</Icon>"
		tmpStr=tmpStr& "<Message><![CDATA[" & post_message & "]]></Message>"
		tmpStr=tmpStr& "<CreatedTime>" & post_created_time & "</CreatedTime>"
		tmpStr=tmpStr& "<PageName>" & page_name & "</PageName>"
		tmpStr=tmpStr& "<PageID>" & page_id & "</PageID>"
		tmpStr=tmpStr& "<TTVShowUrl>" & TTVShowUrl & "</TTVShowUrl>"
		tmpStr=tmpStr& "</DataInfo>"
		rs.movenext
	wend
	dbcon.close
	set dbcon=nothing
	tmpStr=tmpStr& "</DataList>"
	GetData1 = tmpStr
End Function

Sub RemoveString(Src,key)
    Dim o
    Set o = Server.CreateObject("MemcachedCOM.MemCache")
    removeRV = o.RemoveString(Src, key)
    Set o = Nothing
End Sub
 %>

<!DOCTYPE html>
<html lang="zh-tw">
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="author" content="">
    <meta name="keywords" content="艋舺的女人,狄鶯,李興文,陳仙梅,黃仲崑,傅天穎,安定亞,台視" />
    <meta name="Description" content="艋舺的女人，台視優質戲劇。一段發生在保守的時代，想愛卻無法愛的故事。兩個失去至愛，必須相濡以沫、共組家庭的女人；雖然她們沒有丈夫，但卻擁有彼此。" />
    <title>艋舺的女人 - 粉絲塗鴉牆</title>
    <%if ID = "" then %>
    <meta name="keywords" content="塗鴉牆,粉絲團" />
    <meta name="description" content="粉絲塗鴉牆集合了艋舺的女人臉書粉絲團的貼文，讓您不須點來點去遊走各處，即可輕鬆查閱各粉絲團的貼文資訊。" />
    <meta property="og:title" content="艋舺的女人粉絲塗鴉牆" />
    <meta property="og:url" content="https://www.ttv.com.tw/drama14/PrinceWilliam/fb.asp" />
    <meta property="og:image" content="https://www.ttv.com.tw/drama14/PrinceWilliam/images/200x200.jpg"/>
    <meta property="og:description" content="粉絲塗鴉牆集合了艋舺的女人臉書粉絲團的貼文，讓您不須點來點去遊走各處，即可輕鬆查閱各粉絲團的貼文資訊。" />
    <%else %>
    <meta name="keywords" content="塗鴉牆,粉絲團,<%=PageName %>" />
    <meta name="description" content="<%=Message %>" />
    <meta property="og:title" content="<%=PageName %> - 台視粉絲塗鴉牆" />
    <meta property="og:url" content="https://www.ttv.com.tw/group/fbpost/default.asp?id=<%=ID %>" />
    <meta property="og:image" content="<%=Picture %>"/>
    <meta property="og:description" content="<%=Message %>" />
    <%end if %>
    <meta property="og:type" content="website" />
    <meta property="og:site_name" content="台視" />
    <meta property="fb:app_id" content="173424616047838" />

    <link rel="stylesheet" type="text/css" href="/include/js/VerticalTimeline/css/default.css" />
    <link rel="stylesheet" type="text/css" href="/include/js/VerticalTimeline/css/component.css" />
<style type="text/css">
    .content_AutoPager {
        display:none;
        height:100px;
        margin-bottom: 1.5em;
    }
    
    a[rel=next] {
        font-size: 1em;
    }

    .moredata {
        margin:10px auto;
	    bottom: 10px;
	    color: #0066ff;
	    font-size: 1.5em;
	    padding: 1em;
        border-radius: 3px;
        border: 1px solid #CCCCCC;
    }
    .moredata:hover {background-color: rgba(0, 0, 0, 0.3);}
    .moredata A:link, A:visited
    {
        COLOR: #0066ff;
        TEXT-DECORATION: none
    }
    a.external
    {
      color:#ffffff;
      font-size:0.8em;
      background: url(https://i.ttv.com.tw/ProgramInfoImg/external.png) center right no-repeat;
      padding-right: 15px;
      border-bottom:none;
    }
    a.external:hover
    {
        COLOR: #ff0000;
        TEXT-DECORATION: none;
    }

    /* Right content */
    .cbp_tmtimeline > li .AD728 {
	    margin: 0 0 10px 25%;
	    line-height: 1.4;
	    position: relative;
    }

    .content_tag{font-size: 1em;}	
    }

    .sharefrom {text-align:right;padding-top:10px;}

        body
        {
            background:url('images/bgd.png'),linear-gradient(-70deg,#fffbe0 30%,#ffffff 75%,#9d8383 100%);
            background:url('images/bgd.png'),-moz-linear-gradient(-70deg,#fffbe0 30%,#ffffff 75%,#9d8383 100%);
            background:url('images/bgd.png'),-webkit-linear-gradient(-70deg,#fffbe0 30%,#ffffff 75%,#9d8383 100%);        
        }
        .HeaderTitle {
            text-align:center;
        }  
</style>    
    

    
    

    <!-- Bootstrap core CSS -->
    <link href="/drama14/css/bootstrap.min.css" rel="stylesheet">

    <!-- Custom styles for this template -->
    <link href="2014.css" rel="stylesheet">
    <link href="/drama14/css/StyleSheet.css" rel="stylesheet">

    <!-- Just for debugging purposes. Don't actually copy this line! -->
    <!--[if lt IE 9]><script src="js/ie8-responsive-file-warning.js"></script><![endif]-->

    <!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->
    <script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
    <script>    window.jQuery || document.write('<script type="text/javascript" src="https://i.ttv.com.tw/include/js/jquery-1.9.1.min.js"><\/script>')</script>
    <!--[if lt IE 9]><script src="http://html5shim.googlecode.com/svn/trunk/html5.js"></script><![endif]-->
    <script src="js/bootstrap.min.js"></script>
    <script type="text/javascript" src="/include/js/VerticalTimeline/js/modernizr.custom.js"></script>
    <script type="text/javascript" src="/include/js/jquery.autopager-1.0.0.min.js"></script>

    <script type="text/javascript" src="https://apis.google.com/js/plusone.js">
      {"parsetags": "explicit"}
    </script>
</head>

<body>
<!--#include file="top.asp"-->

<!-- Fans Wall Start -->    
<div class="container">
    <a name="PageTop"></a>
	<div class="main">
        <ul class="cbp_tmtimeline" id="listviewU">
        <%if page=1 then
            ShowData()
        end if%>                    
               
        <div class="content_AutoPager" id="container">
        <%if page>1 then ShowData() end if%>
        </div>    	

        <div class="AutoPager_frame">
            <script type="text/javascript">
                var page = 1;
                $(function() {
                    $.autopager({
                        start: function() {
                            //            ajaxstart();
                        },
                        autoLoad: false,
                        content: '.content_AutoPager',
                        link: '#next',
                        load: function(current, next) {
                            $("#listviewU").append($(this).html());
                            page++;
                            $("[id^='msg_" + page.toString() + "']").each(function(index) {
                                $(this).html(replaceURLWithHTMLLinks($(this).html()));
                                //                    console.log($(this).text());
                            });
                            $(".content_AutoPager").html("");
                            PageHit('1111');

                            gapi.plusone.go();
                        }
                    });

                    $('#next').click(function() {
                        $.autopager('load');
                        return false;
                    });

                });
            </script>
        </ul>
				
		<div style="text-align:center;" >
            <a href="fb.asp?page=<%=page+1 %>" rel="next" id="next" class="moredata">：↓：顯示更多內容：↓：</a>&nbsp;&nbsp;        </div>	

        <!--FB+1回到網頁頂端 start-->
        <style>
        .ad-back-to-top {
            position: fixed;
            padding:10px 2px 10px 10px;
            border-radius: 3px;
            border: 1px solid #CCCCCC;

            top: 200px;
            left: 10px;
            display: none;
        }
         </style>
        <div class="ad-back-to-top hidden-xs hidden-sm" >
        <div>
		        <div style="padding:0 0 10px 4px;"><iframe src="https://www.facebook.com/plugins/like.php?href=http://www.ttv.com.tw/group/fbpost/default.asp&amp;layout=box_count&amp;show_faces=false&amp;width=55&amp;action=like&amp;colorscheme=light&amp;height=65" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:55px; height:65px;" allowTransparency="true"></iframe></div>
		        <div class="g-plusone" data-size="tall" data-href="https://www.ttv.com.tw/group/fbpost/default.asp"></div>
        </div>
        </div>
        <script>
            //$(document).ready(function() {
            // 滾動視窗來判斷按鈕顯示或隱藏
            $(window).scroll(function() {
                if ($(this).scrollTop() > 150) {
                    $('.ad-back-to-top').fadeIn(100);
                } else {
                    $('.ad-back-to-top').fadeOut(100);
                }
            });

            // jQuery實現動畫滾動
            $('.back-to-top').click(function(event) {
                event.preventDefault();
                $('html, body').animate({ scrollTop: 0 }, 500);
            })
            //});
        </script>
        <!--FB+1回到網頁頂端 end-->
				

        <!--回到網頁頂端 start-->
        <style>
        .back-to-top {
            position: fixed;
            bottom: 20px;
            right: 10px;
            text-decoration: none;
            color: #EEEEEE;
            background-color: rgba(0, 0, 0, 0.3);
            font-size: 12px;
            padding: 1em;
            display: none;
            border-radius: 3px;
            border: 1px solid #CCCCCC;
        }
        .back-to-top:hover {background-color: rgba(0, 0, 0, 0.3);}
        </style>
        <a href="#" class="back-to-top hidden-xs hidden-sm" >▲回到頂端</a>
        <script>
            //$(document).ready(function() {
            // 滾動視窗來判斷按鈕顯示或隱藏
            $(window).scroll(function() {
                if ($(this).scrollTop() > 150) {
                    $('.back-to-top').fadeIn(100);
                } else {
                    $('.back-to-top').fadeOut(100);
                }
            });

            // jQuery實現動畫滾動
            $('.back-to-top').click(function(event) {
                event.preventDefault();
                $('html, body').animate({ scrollTop: 0 }, 500);
            })
            //});
        </script>
        <!--回到網頁頂端 end-->
		</div>
	</div>

<!-- Fans Wall Stop -->

<!-- FOOTER -->
<div id="footer">
    <div class="container">
        <!--#include file="under.htm"-->
    </div>
</div>

    <script type="text/javascript">
        $('#fb').addClass("active");     
    </script>

    <script type="text/javascript" src="https://apis.google.com/js/plusone.js">
        { lang: 'zh-TW' }
    </script>
  </body>
</html>

<%sub ShowData()
                    if RecordCount > 0 then
                        y = 1
                        Set oData = oXML.selectNodes("/DataList/DataInfo")    
                        for each x in oData
                            icon=x.selectSingleNode("Icon").text
                            c_date=formatdatetime(x.selectSingleNode("CreatedTime").text,2)
                            c_time=formatdatetime(x.selectSingleNode("CreatedTime").text,4)
                            message=x.selectSingleNode("Message").text
                            picture=x.selectSingleNode("Picture").text
                            Link=x.selectSingleNode("Link").text
                            page_name=x.selectSingleNode("PageName").text
                            page_id=x.selectSingleNode("PageID").text
                            post_id=x.selectSingleNode("PostID").text
                            
                            TTVShowUrl = x.selectSingleNode("TTVShowUrl").text
                            if instr(TTVShowUrl,"?") > 0 then
                                TTVShowUrl = TTVShowUrl +"&id="+x.selectSingleNode("PostID").text
                            else
                                TTVShowUrl = TTVShowUrl +"?id="+x.selectSingleNode("PostID").text
                            end if                
                            
                            PostType = x.selectSingleNode("PostType").text
                            if PostType = "link" then
                                cbp_tmicon = "earth"
                            else
                                cbp_tmicon = "screen"
                            end if
                            
                    %>
					<li>
						<time class="cbp_tmtime" datetime="<%=c_date %> <%=c_time %>"><span><%=c_date %></span> <span><%=c_time %></span></time>
						<a href="/group/fbpost/default.asp?id=<%=post_id %>"><div class="cbp_tmicon cbp_tmicon-<%=cbp_tmicon %>"></div></a>
						<div class="cbp_tmlabel">
                            <div style="float:right;"><a href="<%=Link %>" target="_blank" class="external">原文連結</a></div>
							<h2><a href="<%=TTVShowUrl %>" target="_blank"><%=page_name %></a></h2>
                            <%if PostType = "photo" or PostType = "link" then%>
                                <img src="<%=picture %>" style="margin:auto;width:100%;max-width:595px;" id="img_<%=page %>_<%=y %>" /><br />
			                <%
			                end if
			                if PostType = "video" then
			                'https://www.facebook.com/video/embed?video_id=10201324025840261
                            'https://www.facebook.com/photo.php?v=10201324025840261
			                    if instr(Link,"photo.php?v=") > 0 then%>
			                    <iframe src="<%=replace(Link,"photo.php?v=","video/embed?video_id=") %>" width="100%" height="450" frameborder="0"></iframe>
			                <%  elseif instr(Link,"youtube") > 0 or instr(Link,"youtu.be") > 0 then
			                        post_link = replace(replace(replace(replace(replace(replace(Link,"&feature=player_embedded",""),"feature=player_embedded&",""),"watch?v=","embed/"),"https:",""),"&feature=youtu.be",""),"&hd=1","")
			                        post_link_a1 = split(post_link,"&list=")
			                        if ubound(post_link_a1) >= 1 then			                        
			                            post_link = post_link_a1(0)
			                            post_link_a2 = split(post_link,"&feature=")
                                        if ubound(post_link_a2) >= 1 then
                                            post_link = post_link_a2(0)
			                            end if
			                        end if
			                        
                                    if instr(Link,"youtu.be") > 0 then
                                        post_link = replace(post_link,"youtu.be/","www.youtube.com/embed/")
                                    end if
			                '    https://www.youtube.com/watch?v=4UOmV-myo_U
			                '    //www.youtube.com/embed/4UOmV-myo_U?rel=0%>
			                    <iframe width="100%" height="450" src="<%=post_link+"?rel=0" %>" frameborder="0" allowfullscreen></iframe>
			                <%  elseif instr(Link,"youtu.be") > 0 then
			                        post_link = "//www.youtube.com/embed/" + replace(Link,"http://youtu.be/","")
			                            
			                '    http://youtu.be/rnYY4cO61eA  %>
			                    <iframe width="100%" height="450" src="<%=post_link+"?rel=0" %>" frameborder="0" allowfullscreen></iframe>			                    
			                <%  elseif instr(Link,"yahoo.com") > 0 then
			                '    https://tw.omg.yahoo.com/video/%E6%84%9B%E7%9A%84%E7%94%9F%E5%AD%98%E4%B9%8B%E9%81%93-%E7%8D%A8%E5%AE%B6%E8%8A%B1%E7%B5%AE%E9%A6%96%E6%92%AD-151126643.html
			                %>
			                    <iframe width="100%" height="450" src="<%=Link%>?format=embed&player_autoplay=false" frameborder="0" allowfullscreen></iframe>

                            <%   end if
                            end if%>						        
    			            <div id="msg_<%=page %>_<%=y %>"><%=message %></div>
                            <div style="padding-top:10px;"><a href="https://www.facebook.com/<%=page_id %>" target="_blank" class="external">-- 分享自 <%=page_name %> 粉絲團</a></div>


                            <div style="padding-top:10px;">
                		        <div style="padding:0 0 10px 4px;float:left;"><iframe src="https://www.facebook.com/plugins/like.php?href=http://www.ttv.com.tw/group/fbpost/default.asp?id=<%=post_id %>&amp;layout=button_count&amp;show_faces=false&amp;width=90&amp;action=like&amp;colorscheme=light&amp;height=20" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:90px; height:20px;" allowTransparency="true"></iframe></div>
		                        <div class="g-plusone" data-size="medium" data-href="https://www.ttv.com.tw/group/fbpost/default.asp?id=<%=post_id %>"></div>
                            </div>
						</div>
						<%if y / 3 = int(y / 3) then %>
<div class="AD728" width="100%;">
<div style="float:right;">
<script type="text/javascript" src="http://adsense.scupio.com/adpinline/ADmediaJS/ttv_612_1608_2274_1.js"></script>
<!--
<script>
width = document.documentElement.clientWidth;
if (width >= 1000) {
var q = "z=161&w=300&h=250";
}
</script>
<script src="http://ads.doublemax.net/adx/rt_publisher.js" ></script>
-->
</div>

<div>
<script type="text/javascript"><!--
    width = document.documentElement.clientWidth;
    if (width >= 1000) {
        google_ad_client = "pub-5342396684826977";
        /* 300x250, &#65533;w&#65533;&#1573;&#65533; 2008/11/18 */
        google_ad_slot = "6389628022";
        google_ad_width = 300;
        google_ad_height = 250;
    }
//-->
        </script>
<script type="text/javascript" src="https://pagead2.googlesyndication.com/pagead/show_ads.js"></script>
</div>
</div>						
						<%end if %>
					</li>
                <%
                    y=y+1
                    next
                end if
                set oData = nothing
                set oXML = nothing
end sub%>

<script type="text/javascript">
    var loaded = 0;

    $(document).ready(function() {

        //    jQuery("img.lazy").lazy({
        //        beforeLoad: function(element) {
        //console.log(element.attr('id'));
        //		    id_msg = "#msg_1_"+loaded.toString();
        //			$(id_msg).html(replaceURLWithHTMLLinks($(id_msg).text()));
        //			loaded++;
        //        }
        //    });
        PageHit('1100');
        gapi.plusone.go();
    });

    function PageHit(pageid) {
        $.ajax({
            type: "GET",
            crossDomain: true,
            url: "/AD/PageHits.asp",
            data: { PgID: pageid },
            cache: false
        });
    }

    function FBPostClick(PostID) {
        $.ajax({
            type: "GET",
            crossDomain: true,
            url: "/AD/FBPostClick.asp",
            data: { post_id: PostID },
            cache: false
        });
    }

    function replaceURLWithHTMLLinks(text) {
        var exp = /(\b(https?|ftp|file):\/\/[-A-Z0-9+&@#\/%?=~_|!:,.;]*[-A-Z0-9+&@#\/%=~_|])/ig;
        return text.replace(exp, "<a href='$1' target='_blank'>$1</a>");
    }
</script>