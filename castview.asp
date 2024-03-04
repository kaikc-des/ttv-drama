<% Cast=Request.querystring("Cast") On Error Resume Next x=int(Cast) if err.number <> 0 then
    Err.Clear
    Response.Redirect "https://www.ttv.com.tw/drama14/MongaWoman/cast.asp"
    end if
    On Error GoTo 0

    IF Cast="" then
    randomize()
    rNum = Int(15*Rnd+1)
    Cast=rNum
    End IF
    %>
    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
    <html lang="zh-tw">

    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta http-equiv="x-ua-compatible" content="ie=7" />
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <meta name="author" content="TTV">
        <meta name="keywords" content="艋舺的女人,狄鶯,李興文,陳仙梅,黃仲崑,傅天穎,安定亞,台視" />
        <meta name="Description" content="艋舺的女人，台視優質戲劇。一段發生在保守的時代，想愛卻無法愛的故事。兩個失去至愛，必須相濡以沫、共組家庭的女人；雖然她們沒有丈夫，但卻擁有彼此。" />
        <title>艋舺的女人 - 人物</title>
        <meta property="og:title" content="艋舺的女人 - 人物" />
        <meta property="og:type" content="tv_show" />
        <meta property="og:image" content="https://www.ttv.com.tw/drama14/MongaWoman/images/cast<%=cast%>.jpg" />
        <meta property="og:url" content="https://www.ttv.com.tw/drama14/MongaWoman/castview.asp?cast=<%=cast%>" />
        <meta property="og:site_name" content="台視" />
        <meta property="fb:app_id" content="173424616047838" />
        <meta property="og:description"
            content="艋舺的女人，台視優質戲劇。一段發生在保守的時代，想愛卻無法愛的故事。兩個失去至愛，必須相濡以沫、共組家庭的女人；雖然她們沒有丈夫，但卻擁有彼此。" />
        <script type="text/javascript" language="javascript">    AC_FL_RunContent = 0;</script>
        <script type="text/javascript" src="/include/js/AC_RunActiveContent.js" language="javascript"></script>
        <script src="https://ads.ttv.com.tw/AD/PageHits.asp?PgID=1111-561"></script>

        <!-- Bootstrap -->
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

            .cast_title {
                margin: -88px 0px 0px 0px;
                padding: 3em 0.5em 0.5em 0.5em;
                color: White;
                width: 9em;
                font-weight: bold;
                font-size: 1.8em;
                text-align: center;
                letter-spacing: 0.1em;
                text-shadow: 3px 3px 4px #555555;
                font-family: 微軟正黑體, ;
                margin-top: 55%\9
            }

            .color1 {
                background-color: #9d8383;
            }

            .color2 {
                background-color: #879d83;
            }

            .color3 {
                background-color: #a9a171;
            }

            .castview {
                padding: 0.5em;
                overflow: hidden:
            }

            @media screen and (max-width: 480px) {
                .castview {
                    padding-top: 90%
                }

                .cast_title {
                    margin-top: -28px;
                    padding: 0.2em;
                    width: 100%
                }
            }

            @media screen and (min-width: 480px) and (max-width: 767px) {
                .castview {
                    min-height: 500px;
                    padding-top: 70%
                }

                .cast_title {
                    margin-top: -48px;
                    padding: 1em 0.2em 0.2em 0.2em;
                    width: 80%
                }
            }

            @media screen and (min-width: 768px) and (max-width: 979px) {
                .castview {
                    min-height: 700px;
                    padding-top: 60%
                }

                .cast_title {
                    margin-top: -68px;
                    padding: 2em 0.3em 0.3em 0.3em
                }
            }

            @media screen and (min-width: 980px) {
                .castview {
                    min-height: 970px;
                    padding-top: 55%
                }
            }
        </style>

        <!-- Custom styles for this template -->
        <link href="2014.css" rel="stylesheet">

        <!-- Just for debugging purposes. Don't actually copy this line! -->
        <!--[if lt IE 9]><script src="/GoldenMelody25/js/ie8-responsive-file-warning.js"></script><![endif]-->

        <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
        <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
        <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->

        <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
        <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js"></script>
        <!-- Include all compiled plugins (below), or include individual files as needed -->
        <script src="js/bootstrap.min.js"></script>

    </head>

    <body>
        <!--#include file="top.asp"-->

        <a name="c"></a>
        <div class="container">
            <div class="row">
                <div class="well clearfix" style="margin:2em">
                    <%if Cast>=1 and Cast <= 17 then server.Execute "/drama14/MongaWoman/cast/" + cstr(CAST) +".htm"
                            else server.Execute "/drama14/MongaWoman/cast/1.htm" end if%>
                </div>
            </div><!-- /row of columns -->

        </div> <!-- /container -->

        <!-- FOOTER -->
        <div id="footer">
            <div class="container">
                <!--#include file="under.htm"-->
            </div>
        </div>

        <script type="text/javascript">
            $('#cast').addClass("active");     
        </script>

    </body>

    </html>