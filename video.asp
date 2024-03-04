<!--#include virtual="/include/inc/adovbs.inc" -->
<!--#include virtual="/include/inc/killChars.inc"-->
<!--#include virtual="/include/inc/HomeV2ADO.inc"-->
<!--#include virtual="/include/inc/HomeV3ADO.inc"-->
<!--#Include Virtual="/Include/inc/RemoveHTML.inc" -->
<% page=KillChars(Request.QueryString("page")) if page="" then pageindex=0 page=1 else On Error Resume Next
    pageindex=page - 1 if Err.number <> 0 then
    Err.Clear
    Response.End
    end if
    On Error GoTo 0
    end if

    pagesize = 6
    AppName="MongaWoman_default_"+ replace(cstr(date()),"/","") + cstr(hour(now))

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
    response.Write "<!--data from : "+ tmpStr +"-->"

    set oXML = Server.CreateObject("Microsoft.XMLDOM")
    oXML.async = False
    oXML.loadXML(XmlString)
    If oXML.parseError.errorCode <> 0 Then
        response.Write "<!--Error Reason : "+ oXML.parseError.reason + "-->"
        response.Write "<!--Error Line : "+ cstr(oXML.parseError.line) + "-->"
        Response.End
        End If

        'Set oRC = oXML.selectNodes("/DataList/PageTopCount")
        'PageTopCount = 0
        'if oRC.length > 0 then
        ' for each x in oRC
        ' PageTopCount = x.selectSingleNode("rc").text
        ' next
        'end if
        'set oRC = nothing

        Set oRC = oXML.selectNodes("/DataList/RecordCount")
        RecordCount = 0
        if oRC.length > 0 then
        for each x in oRC
        RecordCount = x.selectSingleNode("rc").text
        PageCount = x.selectSingleNode("pagecount").text
        next
        end if
        set oRC = nothing

        if RecordCount <= 0 then response.Redirect "https://www.ttv.com.tw/" end if set oData=nothing Function
            GetDataString(key, getFunc) Dim o, val Set o=Server.CreateObject("MemcachedCOM.MemCache")
            val=o.GetString(AppName , key) If IsEmpty(val) Then Dim afunc Set afunc=GetRef(getFunc) val=afunc()
            setRV=o.SetString(AppName, key, val) Else setRV=200 End If Set o=Nothing GetDataString=val End Function
            Function GetData() tmpStr="<?xml version='1.0' encoding='UTF-8'?>" tmpStr=tmpStr& "<DataList>"
            query=" taglist like '%MongaWoman%' AND YouTubeVideoPublic=1 AND TTVVideoPublic=1 AND YouTubeStatus=3 AND datediff(n, StartTime, getdate())>=0 AND datediff(n, EndTime, getDate())<=0"
            OpenHomeV3 dbcon sql="SELECT CEILING(count(VideoID) / " + cstr(pagesize) +".0) as rcount FROM vVideoAll
            where " & query
    set rs=dbcon.execute(sql)
    pagecount = rs(" rcount") sql="select TOP " + cstr(pagesize) +" * from vVideoAll where VideoID NOT IN (SELECT
            TOP "+ cstr(pagesize * pageindex) +" VideoID FROM vVideoAll where "+ query +" order by
            datediff(d,StartTime,getdate()),title) and "+ query +" order by datediff(d,StartTime,getdate()),title" set
            rs=dbcon.execute(sql) tmpStr=tmpStr& "<RecordCount><rc>" & rs.recordcount &"</rc>
            <pagecount>"& cstr(pagecount) &"</pagecount>
            </RecordCount>"

            do while not rs.eof
            tmpStr=tmpStr& "<DataInfo>"
                tmpStr=tmpStr& "<VideoID>" & rs("VideoID") & "</VideoID>"
                tmpStr=tmpStr& "<YouTubeVideoID>" & rs("YouTubeVideoID") & "</YouTubeVideoID>"
                tmpStr=tmpStr& "<Title>
                    <![CDATA[" & replace(replace(rs("Title"),"台視",""),"HD官方版","") & "]]>
                </Title>"
                tmpStr=tmpStr& "<Intro>
                    <![CDATA[" & rs("Intro") & "]]>
                </Intro>"
                tmpStr=tmpStr& "<Hits>" & rs("hits") & "</Hits>"
                tmpStr=tmpStr& "<Duration>" & rs("Duration") & "</Duration>"
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
            %>

            <!DOCTYPE html>
            <html lang="zh-tw">

            <head>
                <meta charset="utf-8">
                <meta http-equiv="X-UA-Compatible" content="IE=edge">
                <meta name="viewport" content="width=device-width, initial-scale=1">
                <meta name="author" content="">
                <meta name="keywords" content="艋舺的女人,狄鶯,李興文,陳仙梅,黃仲崑,傅天穎,安定亞,台視" />
                <meta name="Description"
                    content="艋舺的女人，台視優質戲劇。一段發生在保守的時代，想愛卻無法愛的故事。兩個失去至愛，必須相濡以沫、共組家庭的女人；雖然她們沒有丈夫，但卻擁有彼此。" />
                <title>艋舺的女人 - 影音</title>
                <meta property="og:title" content="艋舺的女人 - 影音" />
                <meta property="og:type" content="tv_show" />
                <meta property="og:image" content="https://www.ttv.com.tw/drama14/MongaWoman/images/200x200.jpg" />
                <meta property="og:url" content="https://www.ttv.com.tw/drama14/MongaWoman/video.asp" />
                <meta property="og:site_name" content="台視" />
                <meta property="fb:app_id" content="173424616047838" />
                <meta property="og:description"
                    content="艋舺的女人，台視優質戲劇。一段發生在保守的時代，想愛卻無法愛的故事。兩個失去至愛，必須相濡以沫、共組家庭的女人；雖然她們沒有丈夫，但卻擁有彼此。" />
                <link rel="shortcut icon" href="/favicon.ico" />

                <!-- Bootstrap core CSS -->
                <link href="css/bootstrap.min.css" rel="stylesheet">
                <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.2.1/js/bootstrap.min.js"
                    integrity="sha384-B0UglyR+jN6CkvvICOB2joaf5I4l3gm9GU6Hc1og6Ls7i6U/mkkaduKaBhlAXv9k"
                    crossorigin="anonymous"></script>
                <style type="text/css">
                    body {
                        background: linear-gradient(-70deg, #fffbe0 30%, #ffffff 75%, #9d8383 100%);
                        background: -moz-linear-gradient(-70deg, #fffbe0 30%, #ffffff 75%, #9d8383 100%);
                        background: -webkit-linear-gradient(-70deg, #fffbe0 30%, #ffffff 75%, #9d8383 100%);
                    }

                    .HeaderTitle {
                        text-align: center;
                    }

                    .r2p {
                        paddig-right: 2px;
                    }

                    .ellipsis {
                        width: 100%;
                        overflow: hidden;
                        text-overflow: ellipsis;
                        white-space: nowrap;
                    }

                    .ad300x2 {
                        margin: 0 auto;
                        max-width: 620px;
                    }
                </style>
                <link type="text/css" media="screen" rel="stylesheet"
                    href="/include/js/colorbox-master/example5/colorbox.css" />

                <!-- Custom styles for this template -->
                <link href="2014.css" rel="stylesheet">
                <link href="css/StyleSheet.css" rel="stylesheet">

                <!-- Just for debugging purposes. Don't actually copy this line! -->
                <!--[if lt IE 9]><script src="js/ie8-responsive-file-warning.js"></script><![endif]-->

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
                <script type="text/javascript" src="/include/js/jquery.sticky.js"></script>
                <script type="text/javascript" src="/include/js/colorbox-master/jquery.colorbox-min.js"></script>

                <script type="text/javascript" src="https://apis.google.com/js/plusone.js">
                    { "parsetags": "explicit" }
                </script>
            </head>

            <body>
                <div id="fb-root"></div>
                <script>    (function (d, s, id) {
                        var js, fjs = d.getElementsByTagName(s)[0];
                        if (d.getElementById(id)) return;
                        js = d.createElement(s); js.id = id;
                        js.src = "//connect.facebook.net/zh_TW/all.js#xfbml=1&appId=173424616047838";
                        fjs.parentNode.insertBefore(js, fjs);
                    }(document, 'script', 'facebook-jssdk'));</script>

                <!--#include file="top.asp"-->

                <div class="container">
                    <div class="row" id="pagerow" style="margin-top:2em">
                        <div class="col-xs-12 col-sm-12 col-md-8 col-lg-8">
                            <div class="panel panel-default" style="margin-bottom:10px;">
                                <div class="panel-body">
                                    <div class="row" id="container">
                                        <% ShowData() %>
                                    </div><!--/row end -->
                                </div>
                            </div>
                            <div><!-- pagination start -->
                                <% if page=1 then prevPage=1 else prevPage=page - 1 end if if cint(page)=cint(PageCount)
                                    then nextPage=PageCount else nextPage=page + 1 end if if page - 5 <=1 then Pstart=1
                                    else Pstart=page - 5 end if if page + 5>= cint(PageCount) then
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
                                            <li class="<%=PageClass(page,1," disabled") %>"><a href="video.asp?page=1"
                                                    data-toggle="tooltip" data-placement="top" title="跳至第一頁"><span
                                                        class="glyphicon glyphicon-fast-backward"></span></a></li>
                                            <li class="<%=PageClass(page,1," disabled") %>"><a
                                                    href="video.asp?page=<%=prevPage %>" data-toggle="tooltip"
                                                    data-placement="top" title="跳至第<%=prevPage %>頁"><span
                                                        class="glyphicon glyphicon-backward"></span></a></li>
                                            <%for Ploop=Pstart to Pend %>
                                                <li class="<%=PageClass(page,Ploop," active") %>"><a
                                                        href="video.asp?page=<%=Ploop %>" data-toggle="tooltip"
                                                        data-placement="top" title="跳至第<%=Ploop %>頁">
                                                        <%=Ploop %>
                                                    </a></li>
                                                <%next %>
                                                    <li class="<%=PageClass(page,PageCount," disabled") %>"><a
                                                            href="video.asp?page=<%=nextPage %>" data-toggle="tooltip"
                                                            data-placement="top" title="跳至第<%=nextPage %>頁"><span
                                                                class="glyphicon glyphicon-forward"></span></a></li>
                                                    <li class="<%=PageClass(page,PageCount," disabled") %>"><a
                                                            href="video.asp?page=<%=PageCount %>" data-toggle="tooltip"
                                                            data-placement="top" title="跳至最後一頁"><span
                                                                class="glyphicon glyphicon-fast-forward"></span></a>
                                                    </li>
                                        </ul>
                                    </center>
                            </div><!-- pagination end -->

                        </div>

                        <div class="clearfix hidden-sm hidden-xs col-md-4 col-lg-4">
                            <div>
                                <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
                                <!-- 回應式廣告單元 -->
                                <ins class="adsbygoogle" style="display:block" data-ad-client="ca-pub-5342396684826977"
                                    data-ad-slot="3553673898" data-ad-format="auto"></ins>
                                <script>
                                    (adsbygoogle = window.adsbygoogle || []).push({});
                                </script>
                            </div>
                            <div style="CLEAR: both;PADDING-TOP: 10px;">
                                <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
                                <!-- 回應式廣告單元 -->
                                <ins class="adsbygoogle" style="display:block" data-ad-client="ca-pub-5342396684826977"
                                    data-ad-slot="3553673898" data-ad-format="auto"></ins>
                                <script>
                                    (adsbygoogle = window.adsbygoogle || []).push({});
                                </script>
                            </div>
                            <div style="CLEAR: both;PADDING-TOP: 10px;" id="sticker">
                                <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
                                <!-- 回應式廣告單元 -->
                                <ins class="adsbygoogle" style="display:block" data-ad-client="ca-pub-5342396684826977"
                                    data-ad-slot="3553673898" data-ad-format="auto"></ins>
                                <script>
                                    (adsbygoogle = window.adsbygoogle || []).push({});
                                </script>
                            </div>
                        </div>
                        <div style="padding-top:10px;"></div>
                    </div> <!-- /container -->

                    <!-- FOOTER -->
                    <div id="footer">
                        <div class="container">
                            <!--#include file="under.htm"-->
                        </div>
                    </div>

                    <script type="text/javascript">
                        $('#video').addClass("active");     
                    </script>
                    <script type="text/javascript" src="https://apis.google.com/js/plusone.js">
                        { lang: 'zh-TW' }
                    </script>


                    <script type="text/javascript">
                        $(document).ready(function () {
                            var page = 1;
                            if (page >= <%= PageCount %>) $("#next").css("display", "none");

                        setColorBoxVideo();

                        $(document).on("click", ".ytvideo", function () {
                            $.ajax({
                                type: "GET",
                                crossDomain: true,
                                url: "/AD/VideoClick.asp",
                                data: { ID: $(this).data("videoid") },
                                cache: false
                            });
                        });        
    });

                        $('.dropdown-menu').find('#google_searchbox').click(function (e) {
                            e.stopPropagation();
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

                        function setColorBoxVideo() {
                            var iWidth = 960;
                            var iHeight = 540;
                            width = document.documentElement.clientWidth;
                            if (width < 960) {
                                iWidth = width * 0.8;
                                iHeight = iWidth * 0.56;
                            }
                            $(".ytvideo").colorbox({ iframe: true, innerWidth: iWidth, innerHeight: iHeight });
                        }

                        gapi.plusone.go();
                        PageHit('1111-562');

                        $(window).load(function () {
                            $("#sticker").sticky({ topSpacing: 50 });
                        });
                    </script>

                    <!--#include virtual="/group/js/ga.js"-->
            </body>

            </html>
            <% sub ShowData() if RecordCount> 0 then
                Set oData = oXML.selectNodes("/DataList/DataInfo")
                y = 0
                for each x in oData
                y = y + 1
                VideoID=x.selectSingleNode("VideoID").text
                YouTubeVideoID=x.selectSingleNode("YouTubeVideoID").text
                Title=x.selectSingleNode("Title").text
                Intro=x.selectSingleNode("Intro").text
                Duration=x.selectSingleNode("Duration").text
                imgSrc="https://img.youtube.com/vi/"+ YouTubeVideoID +"/0.jpg"
                hits=CLng(x.selectSingleNode("Hits").text)
                htm_url="https://www.ttv.com.tw/videocity/video_play.asp?id=" + VideoID
                %>

                <div class="col-sm-6 col-sm-6 col-md-6 col-lg-6">
                    <a class="ytvideo"
                        href="https://www.youtube.com/embed/<%=YouTubeVideoID%>?fs=1&hl=zh_TW&rel=0&autoplay=1&showinfo=0&autohide=1"
                        title="<%=Title%>" data-videoid="<%=VideoID %>"><img src="<%=imgSrc%>" class="img-thumbnail"
                            style="width:100%;border:0;" alt="<%=Title%>" /></a>
                    <div style="padding-bottom:10px;height:2em;">
                        <a class="ytvideo"
                            href="https://www.youtube.com/embed/<%=YouTubeVideoID%>?fs=1&hl=zh_TW&rel=0&autoplay=1&showinfo=0&autohide=1"
                            title="<%=Title%>" data-videoid="<%=VideoID %>">
                            <%=Title%>
                        </a>
                    </div>

                    <div>
                        <button class="btn btn-default btn-sm" disabled="disabled"><span
                                class="glyphicon glyphicon-hd-video"></span>&nbsp;<%=Duration %></button>
                        <div style="float: left;width:80px;margin-top:5px;">
                            <g:plusone size="medium" href="<%=htm_url%>"></g:plusone>
                        </div>
                        <div style="float: left;margin-top:5px;">
                            <iframe
                                src="https://www.facebook.com/plugins/like.php?href=<%=htm_url%>&amp;layout=button_count&amp;show_faces=false&amp;width=100&amp;action=like&amp;colorscheme=light&amp;height=20"
                                scrolling="no" frameborder="0"
                                style="border:none; overflow:hidden; width:100px; height:20px;"
                                allowTransparency="true"></iframe>

                        </div>
                        <div class="br2x"></div>
                    </div>
                </div>
                <% next else response.Redirect "https://www.ttv.com.tw/" end if set oData=nothing end sub %>