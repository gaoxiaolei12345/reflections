<!--#include file="include/conn.asp"-->
<%
set zxdtrs=server.createobject("adodb.recordset")
zxdtrs.open "SELECT * FROM scusm_news where news_type=2",conn,1,1
set productrs=server.createobject("adodb.recordset")
productrs.open "SELECT * FROM scusm_products WHERE productDM=true;",conn,1,1
set lifetipsrs=server.createobject("adodb.recordset")
lifetipsrs.open "SELECT TOP 8 * FROM scusm_news where news_type=3 ORDER BY news_date DESC;",conn,1,1
set hotproductrs=server.createobject("adodb.recordset")
hotproductrs.open "SELECT * FROM scusm_products where ProductHot=true ORDER BY producthits DESC;",conn,1,1
set producttypers=server.createobject("adodb.recordset")
producttypers.open "SELECT * FROM scusm_type where parentid=0 ORDER BY typeid aSC;",conn,1,1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>四川大学教育超市</title>
<style type="text/css">
.STYLE16 {
	font-family: "宋体";
	font-size: 12px;
	color: #000000;
}
.STYLE20 {font-size: 12px}
.STYLE24 {
	font-size: 13px;
	color: #CC002C;
}
.STYLE25 {font-family: "宋体"; font-size: 12px; color: #000000; }
#Layer3 {
	position:absolute;
	width:205px;
	height:118px;
	z-index:1;
	left: 810px;
	top: 511px;
}
a {
	text-decoration: none;
	font-family: "宋体";
	font-size: 13px;
}
body {
	background-color: #FFFFFF;
	font-family: "宋体";
	font-size: 13px;
	color: #000000;
	text-decoration: none;
}
a:link {
	font-family: "宋体";
	font-size: 13px;
	color: #0000FF;
	text-decoration: none;
}
a:visited {
	font-family: "宋体";
	font-size: 13px;
	color: #0000FF;
	text-decoration: none;
}
a:hover {
	font-family: "宋体";
	font-size: 13px;
	color: #6600FF;
	text-decoration: underline;
}
td img {display: block;}
.STYLE27 {
	font-size: 14px;
	color: #000000;
}
.STYLE28 {color: #FFFFFF}
</style>
</head>
<script type="text/javascript">
function hidead()
{document.getElementById("ad").style.display="none";}
</script>
<div id="ad" style="position:absolute">
<a href="http://cdjycs.scu.edu.cn" target="_blank">
<img src="gif/商品类型下广告.gif"  border="0"></a>
<DIV onclick="hidead();" style="FONT-SIZE: 9pt; CURSOR: hand" align=right>关闭广告×</DIV>
</div>
<script>
var x = 50,y = 60
var xin = true, yin = true
var step = 1
var delay = 1
var obj=document.getElementById("ad")
function floatAD() {
var L=T=0
var R= document.body.clientWidth-obj.offsetWidth
var B = document.body.clientHeight-obj.offsetHeight
obj.style.left = x + document.body.scrollLeft
obj.style.top = y + document.body.scrollTop
x = x + step*(xin?1:-1)
if (x < L) { xin = true; x = L}
if (x > R){ xin = false; x = R}
y = y + step*(yin?1:-1)
if (y < T) { yin = true; y = T }
if (y > B) { yin = false; y = B }
}
var itl= setInterval("floatAD()", delay)
obj.onmouseover=function(){clearInterval(itl)}
obj.onmouseout=function(){itl=setInterval("floatAD()", delay)}
</script>
<body>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FF0033">
  <!--DWLayoutTable-->
  <tr>
    <td height="164" colspan="7" valign="top" bordercolor="#FF0033"><img src="gif/banner2.gif" width="1000" height="164" /></td>
  </tr>
  <tr>
    <td width="496" height="35" valign="top" background="gif/banner-little.gif" bgcolor="#FFFFFF"><table width="476" height="29" border="1" cellpadding="3" cellspacing="0" bordercolor="#cc002c">
      <tr>
        <td width="47">&nbsp;</td>
        <td width="187" background="gif/时间框.gif"><div align="center" class="STYLE16">
         <SCRIPT language=Javascript>                       
<!--                       
calendar = new Date();                       
day = calendar.getDay();                       
month = calendar.getMonth();                       
date = calendar.getDate();                       
year = calendar.getYear();                       
if (year< 100) year = 1900 + year;                       
cent = parseInt(year/100);                       
g = year % 19;                       
k = parseInt((cent - 17)/25);                       
i = (cent - parseInt(cent/4) - parseInt((cent - k)/3) + 19*g + 15) % 30;                       
i = i - parseInt(i/28)*(1 - parseInt(i/28)*parseInt(29/(i+1))*parseInt((21-g)/11));                       
j = (year + parseInt(year/4) + i + 2 - cent + parseInt(cent/4)) % 7;                       
l = i - j;                       
emonth = 3 + parseInt((l + 40)/44);                       
edate = l + 28 - 31*parseInt((emonth/4));                       
emonth--;                       
var dayname = new Array ("星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六");                       
var monthname =                        
new Array ("1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月" );                       
document.write("<font color=#000000>"+year +"年");                       
document.write(monthname[month]);                       
document.write(date + "日");                       
document.write(dayname[day]+" "+"</font>");                       
document.write("<br></font>");                       
//-->                       
</SCRIPT>
        </div></td>
        <td width="216">&nbsp;</td>
      </tr>
    </table></td>
    <td width="52" valign="middle" background="gif/banner-little.gif" bgcolor="#FFFFFF"><div align="center"><span class="STYLE24"><a href="default.asp">首页</a></span></div>    </td>
    <td width="90" valign="middle" background="gif/banner-little.gif" bgcolor="#FFFFFF"><div align="center"><span class="STYLE24"><a href="about SM/About supermarket on.asp">关于超市</a></span></div></td>
    <td width="92" valign="middle" background="gif/banner-little.gif" bgcolor="#FFFFFF"><div align="center"><span class="STYLE24"><a href="special price/special price.asp">特价商品</a></span></div></td>
    <td width="93" valign="middle" background="gif/banner-little.gif" bgcolor="#FFFFFF"><div align="center" class="STYLE24"><a href="product show/product show.asp">商品展示</a></div></td>
    <td width="86" valign="middle" background="gif/banner-little.gif" bgcolor="#FFFFFF"><div align="center"><span class="STYLE24"><a href="newss/index.asp">最新动态</a></span></div></td>
    <td width="91" valign="middle" background="gif/banner-little.gif" bgcolor="#FFFFFF"><div align="center"><span class="STYLE24"><a href="contact us/contact_us.html">联系我们</a></span></div></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <!--DWLayoutTable-->
  <tr>
    <td width="7" height="10"></td>
    <td width="274"></td>
    <td width="7"></td>
    <td width="15"></td>
    <td width="399"></td>
    <td width="9"></td>
    <td width="90"></td>
    <td width="19"></td>
    <td width="118"></td>
    <td width="17"></td>
    <td width="33"></td>
    <td width="12"></td>
  </tr>
  <tr>
    <td height="28"></td>
    <td valign="middle" bgcolor="#FFFFFF"><img src="gif/cute-banner2.gif" width="274" height="28" /></td>
    <td></td>
    <td colspan="2" rowspan="5" valign="top" bgcolor="#FFFFFF"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" id="scriptmain" name="scriptmain" codebase="http://download.macromedia.com/pub/shockwave/cabs/
flash/swflash.cab#version=6,0,29,0" width="413" height="352">
      <param name="movie" value="bcastr.swf?bcastr_xml_url=xml/bcastr.xml" />
      <param name="quality" value="high" />
      <param name="scale" value="noscale" />
      <param name="LOOP" value="false" />
      <param name="menu" value="false" />
      <param name="wmode" value="transparent" />
      <embed src="bcastr.swf?bcastr_xml_url=xml/bcastr.xml" width="413" height="352" loop="False" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" salign="T" name="scriptmain" menu="false" wmode="transparent"></embed>
    </object></td>
    <td>&nbsp;</td>
    <td colspan="5" valign="middle" bgcolor="#FFFFFF"><img src="gif/商品排行.gif" width="277" height="28" /></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="257"></td>
    <td valign="top" background="gif/最新动态下.gif" bgcolor="#CCCCCC"><table width="270" border="0" align="center" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#FFFFFF">
	<%if  zxdtrs.bof and  zxdtrs.eof then
		no_msg="暂时还没有最新动态信息"
		else
		for i=0 to 11
		if i<zxdtrs.recordcount-1 then
    %>
	  <tr>
        <td><span class="STYLE20">·<a href="newss/text.asp?id=<%=zxdtrs("news_id")%>"><%=zxdtrs("news_title")%></a></span></td>
      </tr>
	<%zxdtrs.movenext
	   else
	   %>
	  <tr>
        <td>&nbsp;</td>
      </tr>
	   <%
	   end if
	next
	end if
	set zxdtrs=nothing%>
    </table>  </td>
    <td></td>
    <td>&nbsp;</td>
    <td colspan="5" rowspan="4" valign="top" background="gif/商品排行下.gif" bgcolor="#CCCCCC"><table width="272" border="0" align="center" cellpadding="5" cellspacing="0" bordercolor="#CCCCCC">
	
	
		     <%
    if hotproductrs.eof and hotproductrs.bof then
		response.Write "<tr><td>暂时没有产品信息</td></tr>"
	else
	for i=1 to 10
  if i<=hotproductrs.recordcount then%>
      <tr>
        <td width="22"><span class="STYLE20">NO<%=i%></span></td>
        <td width="226" class="STYLE16"><a href="新建文件夹/show.asp?id=<%=hotproductrs("productid")%>" class="STYLE16"><%=hotproductrs("productname")%></a> </td>
      </tr>
	<%
	hotproductrs.movenext
		else%>
      <tr>
        <td width="22"><span class="STYLE20">&nbsp;</span></td>
        <td width="226" class="STYLE16"><a href="新建文件夹/show.html" class="STYLE16">&nbsp;</a> </td>
      </tr>
	<%end if
	next
	end if
	set hotproductrs=nothing
	%>

    </table></td>
  <td></td>
  </tr>
  
  <tr>
    <td height="18">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="28"></td>
    <td valign="middle" bgcolor="#FFFFFF"><img src="gif/商品类型.gif" width="274" height="28" /></td>
    <td></td>
    <td>&nbsp;</td>
    <td></td>
  </tr>
  <tr>
    <td height="21"></td>
    <td rowspan="6" valign="top" background="gif/商品类型下.gif">
	
	
	
	
	
	<table width="269" border="1" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
	<%
	if producttypers.eof and producttypers.bof then
		response.Write "暂时没有商品分类信息"
	else
	   for i=0 to producttypers.recordcount-1
		set productsubtypers=server.createobject("adodb.recordset")
		productsubtypers.open "SELECT top 3 * FROM scusm_type where parentid="&producttypers("typeid")&" ORDER BY typeid aSC;",conn,1,1
'		if productsubtypers.recordcount<3 and productsubtypers.recordcount<>0 then
'			row=productsubtypers.recordcount
'		else
'			row=3
'		end if
	%>
		<tr>
        <td width="106" rowspan="4"><div align="center" class="STYLE25"><%=producttypers("typename")%></div>        </td>
         <%for j=0 to 2
		 	if j<productsubtypers.recordcount then
			if j=1 then%><td><div align="center" class="STYLE16"><a href="product show/product show.asp?type=<%=productsubtypers("typeid")%>"><%=productsubtypers("typename")%></a></div></td>
        </tr>
		<%else%>
		<tr>
        <td width="145"><div align="center" class="STYLE16"><a href="product show/product show.asp?type=<%=productsubtypers("typeid")%>"><%=productsubtypers("typename")%></a></div></td>
        </tr>
		<%
			end if
					productsubtypers.movenext
		else
			if j=1 then%><td><div align="center" class="STYLE16">&nbsp;</div></td>
        </tr>
		<%else%>
		<tr>
        <td width="145"><div align="center" class="STYLE16">&nbsp;</div></td>
        </tr>
		<%
			end if
		end if
		next
		%>
	<%
		producttypers.movenext
		next
	end if
		set productsubtypers=nothing
	set producttypers=nothing
%>
    </table></td>
    <td></td>
    <td>&nbsp;</td>
    <td></td>
  </tr>
  
  
  
  
  
  
  <tr>
    <td height="8"></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td height="66"></td>
    <td></td>
    <td></td>
    <td colspan="5" valign="top" bgcolor="#333399"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="634" height="66" title="中间广告">
      <param name="movie" value="gif/中间广告.swf" />
      <param name="quality" value="high" />
	  <param name="wmode" value="opaque">
      <embed src="gif/中间广告.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="634" height="66"></embed>
    </object></td>
    <td>&nbsp;</td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td height="9"></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td height="30"></td>
    <td></td>
    <td colspan="4" valign="bottom" bgcolor="#FFFFFF"><img src="gif/特价商品.gif" width="512" height="30" align="bottom" /></td>
    <td>&nbsp;</td>
    <td colspan="3" valign="top"><a href="special price/special price.asp"><img src="gif/更多.gif" width="167" height="30" border="0" align="bottom" /></a></td>
    <td></td>
  </tr>
  <tr>
    <td height="473"></td>
    <td></td>
    <td colspan="8" rowspan="6" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#333399" bgcolor="#CCCCCC">
      <!--DWLayoutTable-->
      <tr>
        <td width="699" height="648"><table width="699" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<!-- fwtable fwsrc="背景.png" fwbase="背景.png" fwstyle="Dreamweaver" fwdocid = "2057540629" fwnested="0" -->
  <tr>
   <td><img src="gif/spacer.gif" width="12" height="1" border="0" alt="" /></td>
   <td><img src="gif/spacer.gif" width="160" height="1" border="0" alt="" /></td>
   <td><img src="gif/spacer.gif" width="10" height="1" border="0" alt="" /></td>
   <td><img src="gif/spacer.gif" width="160" height="1" border="0" alt="" /></td>
   <td><img src="gif/spacer.gif" width="10" height="1" border="0" alt="" /></td>
   <td><img src="gif/spacer.gif" width="160" height="1" border="0" alt="" /></td>
   <td><img src="gif/spacer.gif" width="10" height="1" border="0" alt="" /></td>
   <td><img src="gif/spacer.gif" width="160" height="1" border="0" alt="" /></td>
   <td><img src="gif/spacer.gif" width="17" height="1" border="0" alt="" /></td>
   <td><img src="gif/spacer.gif" width="1" height="1" border="0" alt="" /></td>
  </tr>

  <tr>
   <td colspan="9"><img name="n_r1_c1" src="gif/背景_r1_c1.gif" width="699" height="11" border="0" id="n_r1_c1" alt="" /></td>
   <td><img src="gif/spacer.gif" width="1" height="11" border="0" alt="" /></td>
  </tr>
  <tr>
   <td rowspan="9" ><img name="n_r2_c1" src="gif/背景_r2_c1.gif" width="12" height="671" border="0" id="n_r2_c1" alt="" /></td>
   <td colspan="7" rowspan="8" valign="top">
   <table border="0" align="center" width="100%" cellpadding="0" cellspacing="0">
   <%
    if productrs.eof and productrs.bof then
		response.Write "<tr><td>暂时没有产品信息</td></tr>"
	else
	for i=0 to 11
	if i mod 4=0 then
	%>
  <tr>
  <%end if%>
  <%if i<productrs.recordcount then%>
	<td align="center"><table border="0" cellpadding="0" cellspacing="0"><tr>
	  <%if productrs("productimage")<>"" then%><td><a href="新建文件夹/show.asp?id=<%=productrs("productid")%>"><img src="up_files/product/<%=productrs("productimage")%>" width="160" height="170" border="0"/></a>
	    <%else%><td width="160" height="170"><a href="新建文件夹/show.asp?id=<%=productrs("productid")%>"><%=productrs("productname")%></a><%end if%></td>
	</tr><tr><td align="center"><div align="center" class="STYLE20">原价:<%=productrs("productprice")%>元 <span class="STYLE24">现价<%=productrs("productdmprice")%>元</span></div></td>
	</tr>
	    
	</table></td>
	<%
	productrs.movenext
		else%>
	<td align="center"><table border="0" cellpadding="0" cellspacing="0"><tr><td width="160" height="170">&nbsp;</td></tr><tr><td align="center">&nbsp;</td>
	</tr>
	    
	</table></td>
	
	<%end if%>
	<%if i mod 4=3 then%>
	</tr>
  <%
	end if

	next
	end if
	set productrs=nothing
	%>
   </table>   </td>
   <td rowspan="9"><img name="n_r2_c9" src="gif/背景_r2_c9.gif" width="17" height="671" border="0" id="n_r2_c9" alt="" /></td>
   <td><img src="gif/spacer.gif" width="1" height="170" border="0" alt="" /></td>
  </tr>
  <tr>
   <td><img src="gif/spacer.gif" width="1" height="40" border="0" alt="" /></td>
  </tr>
  <tr>
   <td><img src="gif/spacer.gif" width="1" height="8" border="0" alt="" /></td>
  </tr>
  <tr>
   <td><img src="gif/spacer.gif" width="1" height="170" border="0" alt="" /></td>
  </tr>
  <tr>
   <td><img src="gif/spacer.gif" width="1" height="40" border="0" alt="" /></td>
  </tr>
  <tr>
   <td><img src="gif/spacer.gif" width="1" height="8" border="0" alt="" /></td>
  </tr>
  <tr>
   <td><img src="gif/spacer.gif" width="1" height="170" border="0" alt="" /></td>
  </tr>
  <tr>
   <td><img src="gif/spacer.gif" width="1" height="40" border="0" alt="" /></td>
  </tr>
  <tr>
   <td><img name="n_r10_c2" src="gif/背景_r10_c2.gif" width="160" height="25" border="0" id="n_r10_c2" alt="" /></td>
   <td><img name="n_r10_c3" src="gif/背景_r10_c3.gif" width="10" height="25" border="0" id="n_r10_c3" alt="" /></td>
   <td><img name="n_r10_c4" src="gif/背景_r10_c4.gif" width="160" height="25" border="0" id="n_r10_c4" alt="" /></td>
   <td><img name="n_r10_c5" src="gif/背景_r10_c5.gif" width="10" height="25" border="0" id="n_r10_c5" alt="" /></td>
   <td><img name="n_r10_c6" src="gif/背景_r10_c6.gif" width="160" height="25" border="0" id="n_r10_c6" alt="" /></td>
   <td><img name="n_r10_c7" src="gif/背景_r10_c7.gif" width="10" height="25" border="0" id="n_r10_c7" alt="" /></td>
   <td><img name="n_r10_c8" src="gif/背景_r10_c8.gif" width="160" height="25" border="0" id="n_r10_c8" alt="" /></td>
   <td><img src="gif/spacer.gif" width="1" height="25" border="0" alt="" /></td>
  </tr>
</table></td>
      </tr>
    </table></td>
    <td></td>
  </tr>
  <tr>
    <td height="25" rowspan="2"></td>
    <td><a href="contact us/contact_us.asp"><img src="gif/商品类型下下.gif" width="274" height="150" border="0" /></a></td>
    <td rowspan="2"></td>
    <td rowspan="2"></td>
  </tr>
  <tr>
    <td><!--DWLayoutEmptyCell-->&nbsp;</td>
  </tr>
  
  <tr>
    <td height="28"></td>
    <td valign="top"><img src="gif/最新加入商品.gif" width="274" height="28" /></td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td height="97"></td>
    <td rowspan="2" valign="top"><table width="274" border="1" cellpadding="3" cellspacing="0" bordercolor="#333399">
	 <tr><td height="100" align="left" valign="top">    <%
    if lifetipsrs.eof and lifetipsrs.bof then
		response.Write "暂时没有信息"
	else
	for i=0 to 5
	%>
  
  <%if i<lifetipsrs.recordcount then%>
        <span class="STYLE20">&middot;<a href="newss/text.asp?id=<%=lifetipsrs("news_id")%>"><%=lifetipsrs("news_title")%></a></span><br />
	<%
	lifetipsrs.movenext
       end if%>
	
  <%
	next
	end if
	set lifetipsrs=nothing
	%></td></tr>
    </table></td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td height="2"></td>
    <td></td>
    <td></td>
  </tr>
  <tr>
    <td height="31"></td>
    <td></td>
    <td></td>
    <td></td>
    <td colspan="6" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <!--DWLayoutTable-->
      <tr>
        <td width="651" height="31"><table width="651" border="1" cellspacing="0" cellpadding="3">
          <tr>
            <td width="215"><span class="STYLE27">相关链接:</span></td>
            <td width="100"><a href="http://www.scu.edu.cn/"><img src="gif/链接1.gif" width="100" height="23" border="0" /></a></td>
            <td width="100"><a href="http://202.115.32.129/hq/main/index.php"><img src="gif/链接2.gif" width="100" height="23" border="0" /></a></td>
            <td width="100"><a href="http://www.mmtea.com/"><img src="gif/0.jpg" width="100" height="23" border="0" /></a></td>
            <td width="100"><a href="http://www.pepsico.com.cn/"><img src="gif/1.jpg" width="100" height="23" border="0" /></a></td>
            <td width="100"><a href="http://www.zhenmeifoods.com/newEbiz1/EbizPortalFG/portal/html/main.html"><img src="gif/2.jpg" width="100" height="23" border="0" /></a></td>
          </tr>
        </table></td>
      </tr>
    </table>    </td>
    <td></td>
    <td></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#333399">
  <!--DWLayoutTable-->
  <tr>
    <td width="1000" height="25" valign="top"><img src="gif/下部banner.gif" width="1000" height="50" /></td>
  </tr>
  <tr>
    <td height="25" align="center"><span class="STYLE28"> 版权所有 Copyright &copy;2009  All Right Preserved <a href="http://cdjycs.scu.edu.cn/admin/index.html">管理入口</a> </span></td>
  </tr>
  <tr>
    <td height="12" align="center"><span class="STYLE28">地址：四川大学望江校区文华大道 电话:028-85466006  </span></td>
  </tr>
  <tr>
    <td height="13" align="center"><span class="STYLE28">技术支持：腾格尔 QQ：55538511 E-Mail:<a href="mailto:tenger12345@163.com">tenger12345@163.com</a> 崔巍 QQ：283827681</span></td>
  </tr>
</table>
</body>
</html>