<!--#include virtual="/include/inc/HomeV2ADO.inc"-->
<% epi=int(request("epi")) PgmName="艋舺的女人2017" MCachedKey="MongaWoman" 'Cached資料ID, 用節目英文名

if epi=0 then
	OpenHomeV2 dbcon
	sqlE = "Select Top 1 episode from tblPgmInfo Where PgmName = '" & PgmName & "' And ShowOn = ' 1' and
    datediff(d,ShowDate,getdate())>=0 order by episode desc"
    epi = dbcon.execute(sqlE).fields(0).value
    dbcon.close
    set dbcon=nothing
    end if

    setRV = -100
    removeRV = -100
    tmpStr = ""
    xmlString = ""
    MaxEpi=0
    RemoveMemCached=LCase(request("Action"))

    if RemoveMemCached="removememcached" then
    OpenHomeV2 dbcon
    sqlE = "Select Top 1 episode from tblPgmInfo Where PgmName = '" & PgmName & "' And ShowOn = '1' order by episode
    desc"
    epi = dbcon.execute(sqlE).fields(0).value
    dbcon.close
    set dbcon=nothing
    for sh=1 to epi
    RKeyName=MCachedKey & "-EP" & sh 'MemCached Key
    call RemoveString(MCachedKey & "-Plot",RKeyName)
    next
    Response.end
    end if

    KeyName=MCachedKey & "-EP" & epi 'MemCached Key
    If epi<>0 Then
        xmlString = GetString(KeyName, "GetPlot") 'o
        If setRV = 200 Then
        tmpStr = "Cached"
        Else
        tmpStr = "DB"
        End If
        End If

        Function GetPlot()
        tmpStr= "
        <?xml version='1.0' encoding='UTF-8'?>"
        tmpStr=tmpStr& "<Plot>"
            OpenHomeV2 dbcon

            sqlE = "Select Top 1 episode from tblPgmInfo Where PgmName = '" & PgmName & "' And ShowOn = '1' and
            datediff(d,ShowDate,getdate())>=0 order by episode desc"
            MaxEpi = dbcon.execute(sqlE).fields(0).value

            tmpStr=tmpStr& "<Episode>" & MaxEpi & "</Episode>"
            sql = "Select InfoID, ShowTime,Description, CreateDate from tblPgmInfo Where episode = '" & epi & "' and
            PgmName = '" & PgmName & "' and ShowOn = '1' Order By InfoID"
            set rs = dbcon.execute(sql)
            while not rs.eof
            if trim(rs("CreateDate")) <> "" then
                Parent = year(rs("CreateDate")) & "-" & month(rs("CreateDate"))
                else
                Parent = "NODate"
                end if
                Title = trim(rs("ShowTime"))
                Description = rs("Description")
                InfoID=rs("InfoID")

                tmpStr=tmpStr& "<Information>"
                    tmpStr=tmpStr& "<Info_ID>" & InfoID & "</Info_ID>"
                    tmpStr=tmpStr& "<Title>
                        <![CDATA[" & Title & "]]>
                    </Title>"
                    tmpStr=tmpStr& "<Description>
                        <![CDATA[" & Description & "]]>
                    </Description>"
                    tmpStr=tmpStr& "<Photo>:::::"
                        ' IMAGE
                        sqlImg="select SID from tblPgmInfoImg where InfoID='" & InfoID & "' and (ContentType is NOT NULL
                        and rtrim(ltrim(ContentType)) <> '')"
                            set rsImg=dbcon.execute(sqlImg)
                            while not rsImg.eof
                            imgSrc="https://www.ttv.com.tw/homev3/ProgramInfoImg/" & Parent & "/" & InfoID & "/" &
                            rsImg("SID") & ".jpg"
                            tmpStr=tmpStr & imgSrc & ":::::"
                            rsImg.movenext
                            wend
                            rsImg.close
                            set rsImg=nothing
                            tmpStr=tmpStr& "</Photo>"
                    tmpStr=tmpStr& "</Information>"
                rs.movenext
                wend
                dbcon.close
                set dbcon=nothing
                tmpStr=tmpStr& "</Plot>"
        GetPlot = tmpStr
        End Function

        Function GetString(key, getPlotFuncName)
        Dim o, val
        Set o = Server.CreateObject("MemcachedCOM.MemCache")
        val = o.GetString(MCachedKey & "-Plot", key) 'q memcached &#362; ("" `{W, Key &#357;V)
        If IsEmpty(val) Then
        Dim afunc
        Set afunc = GetRef(getPlotFuncName)
        val = afunc()
        setRV = o.SetString(MCachedKey & "-Plot", key, val) 'N&#422;sJ memcached
        Else
        setRV = 200
        End If
        Set o = Nothing
        GetString = val
        End Function

        Sub RemoveString(Src,key)
        Dim o
        Set o = Server.CreateObject("MemcachedCOM.MemCache")
        removeRV = o.RemoveString(Src, key) 'N&#433;q memcached
        Set o = Nothing
        End Sub

        set objXML = Server.CreateObject("Microsoft.XMLDOM")
        objXML.async = False
        objXML.loadXML(XmlString)
        If objXML.parseError.errorCode <> 0 Then
            Response.End
            End If

            Set objLst = objXML.getElementsByTagName("Information")
            noOfInfo = objLst.length
            set objList=nothing

            if noOfInfo>0 then
            MaxEpi=int(objXML.getElementsByTagName("Episode").item(i).childNodes(0).text)
            InfoID=objXML.getElementsByTagName("Info_ID").item(i).childNodes(0).text
            Title=objXML.getElementsByTagName("Title").item(i).childNodes(0).text
            Description=objXML.getElementsByTagName("Description").item(i).childNodes(0).text
            PhotoList=objXML.getElementsByTagName("Photo").item(i).childNodes(0).text
            end if
            set objXML=nothing
            %>
            <!DOCTYPE html>
            <html>

            <head>
                <meta charset="utf-8">
                <meta http-equiv="X-UA-Compatible" content="IE=edge">
                <meta name="viewport" content="width=device-width, initial-scale=1">
                <meta name="author" content="">
                <meta name="keywords" content="艋舺的女人,狄鶯,李興文,陳仙梅,黃仲崑,傅天穎,安定亞,台視" />
                <meta name="Description"
                    content="艋舺的女人，台視優質戲劇。一段發生在保守的時代，想愛卻無法愛的故事。兩個失去至愛，必須相濡以沫、共組家庭的女人；雖然她們沒有丈夫，但卻擁有彼此。" />
                <title>艋舺的女人 - 劇情</title>
                <meta property="og:title" content="艋舺的女人 - 劇情" />
                <meta property="og:type" content="tv_show" />
                <meta property="og:image" content="https://www.ttv.com.tw/drama14/MongaWoman/images/200x200.jpg" />
                <meta property="og:url" content="https://www.ttv.com.tw/drama14/MongaWoman/plot.asp" />
                <meta property="og:site_name" content="台視" />
                <meta property="fb:app_id" content="173424616047838" />
                <meta property="og:description"
                    content="艋舺的女人，台視優質戲劇。一段發生在保守的時代，想愛卻無法愛的故事。兩個失去至愛，必須相濡以沫、共組家庭的女人；雖然她們沒有丈夫，但卻擁有彼此。" />
                <script src="https://ads.ttv.com.tw/AD/PageHits.asp?PgID=1111-560"></script>

                <!-- Bootstrap core CSS -->
                <link rel="stylesheet" href="css/bootstrap.min.css">
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

                    .advertisement {
                        background-color: #cccccc;
                        font-size: 9px;
                        text-align: center;
                        margin: 5px;
                        padding: 1px 1px 5px 1px;
                    }

                    .advertisement::after {
                        content: "ADVERTISEMENT";
                    }
                </style>

                <!-- Custom styles for this template -->
                <link href="2014.css" rel="stylesheet">

                <!-- Just for debugging purposes. Don't actually copy this line! -->
                <!--[if lt IE 9]><script src="js/ie8-responsive-file-warning.js"></script><![endif]-->

                <!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
                <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->

                <script src="https://ads.ttv.com.tw/AD/PageHits.asp?PgID=1111"></script>

                <script type="text/javascript">
                    function fblike() {
                        var ttvwebpage = document.getElementById("ttvwebpage");
                        ttvwebpage.style.display = "block";
                    }
                </script>
            </head>

            <body>
                <!--#include file="top.asp"-->
                <div class="jumbotron2">
                    <div class="container">
                        <!-- row of columns -->
                        <div class="row">

                            <div class="col-md-8 col-lg-8" style="padding-top:1.5em">
                                <%PhotoListA=split(PhotoList,":::::") if ubound(PhotoListA)>1 then%>
                                    <img src="<%=PhotoListA(1)%>" class="featurette-image img-circle pull-right img512"
                                        style="border:1px #ffffff solid;margin:0px 16px 0px 0px;">
                                    <%end if if ubound(PhotoListA)>2 then
                                        for j=2 to (ubound(PhotoListA)-1)%>
                                        <img src="<%=PhotoListA(j)%>"
                                            class="featurette-image img-circle pull-left img512"
                                            style="border:1px #ffffff solid;margin:10px 0px 20px 0px">
                                        <%next end if%>

                                            <!--start選擇集數-->
                                            <select name="02"
                                                onChange="location.href=this.options[this.selectedIndex].value">
                                                <option value="#">請選擇集數</option>
                                                <%for ii=1 to MaxEpi%>
                                                    <option value="/drama14/MongaWoman/plot.asp?epi=<%=ii%>&unit=2#tip"
                                                        <%if epi=ii then%>selected<%end if%>> 第
                                                            <%=ii%> 集 </option>
                                                    <% next%>
                                            </select>
                                            <!--end選擇集數-->
                                            <h2 class="title">
                                                <%=Title%>
                                            </h2></span>
                                            <div>
                                                <%=Description%>
                                            </div>

                                            <div>
                                                <div id="fb-root"></div>
                                                <script
                                                    src="https://connect.facebook.net/zh_TW/all.js#appId=121631524541537&amp;xfbml=1"></script>
                                                <div style="height:30px"></div>
                                                <fb:like
                                                    href="https://www.ttv.com.tw/drama14/MongaWoman/plot.asp?epi=<%=epi %>"
                                                    send="false" layout="standard" width="100%" show_faces="false"
                                                    action="like" font=""></fb:like>
                                            </div>

                                            <script type="text/javascript">
                                                FB.Event.subscribe('edge.create',
                                                    function (response) {
                                                        fblike();
                                                    }
                                                ); 
                                            </script>

                                            <div id="ttvwebpage"
                                                style="width:400px;margin-top:5px;background:#fbb7ff;padding:5px  5px 5px 5px;text-align:left;-webkit-border-radius: 3px;-moz-border-radius: 3px;border-radius: 3px;display:none;">
                                                <span
                                                    style="margin:5px 5px 0 5px;height:21px;float:left;font-size:13px;"><a
                                                        href="https://www.facebook.com/ttvweb" target="_blank"><span
                                                            style="color:#3B5998;">台視網站給您最優質的官網： </span></a></span>
                                                <iframe
                                                    src="https://www.facebook.com/plugins/like.php?app_id=129791680451367&amp;href=http%3A %2F%2Fwww.facebook.com %2Fttvweb&amp;send=false&amp;layout=standard&amp;width=450&amp;show_faces=false&amp;action=like&amp;colorscheme=lig ht&amp;font&amp;height=21"
                                                    scrolling="no" frameborder="0"
                                                    style="border:none; overflow:hidden; width:450px;  height:24px;"
                                                    allowTransparency="true"></iframe>
                                            </div>
                            </div>
                            <div class="col-md-3 col-lg-3" style="margin-right:60px;padding-top:1.5em">
                                <script type="text/javascript"><!--
google_ad_client = "ca-pub-5342396684826977";
/* 300 x 600 大型摩天大廣告 */
                                    google_ad_slot = "9112337892";
                                    google_ad_width = 300;
                                    google_ad_height = 600;
                                    //-->
                                </script>
                                <script type="text/javascript"
                                    src="https://pagead2.googlesyndication.com/pagead/show_ads.js">
                                    </script>

                            </div>

                        </div>
                        <!-- /row of columns -->

                    </div> <!-- /container -->
                </div>

                <!-- FOOTER -->
                <div id="footer">
                    <div class="container">
                        <!--#include file="under.htm"-->
                    </div>
                </div>


                <!-- Bootstrap core JavaScript
    ================================================== -->
                <!-- Placed at the end of the document so the pages load faster -->
                <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js"></script>
                <script src="js/bootstrap.min.js"></script>


                <script type="text/javascript">
                    $('#plot').addClass("active");     
                </script>
            </body>

            </html>