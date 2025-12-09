<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Demo.aspx.cs" Inherits="EInvoice.Demo" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 15">
<link rel=File-List href="HOA%20DON%20_YUPOONG_c_files/filelist.xml">
<!--[if !mso]>
<style>
v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
x\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style>
<![endif]-->
<style id="HOA DON _YUPOONG_c_17145_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.font517145
	{color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font617145
	{color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font717145
	{color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font817145
	{color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font917145
	{color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1017145
	{color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1117145
	{color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1217145
	{color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1317145
	{color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1417145
	{color:#993300;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1517145
	{color:#333399;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1617145
	{color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1717145
	{color:red;
	font-size:11.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1817145
	{color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font1917145
	{color:red;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font2017145
	{color:red;
	font-size:10.0pt;
	font-weight:700;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font2117145
	{color:windowtext;
	font-size:11.5pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font2217145
	{color:#0066CC;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font2317145
	{color:red;
	font-size:17.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font2417145
	{color:red;
	font-size:17.0pt;
	font-weight:700;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.xl6517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl6717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl6817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl6917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl7017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl7117145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl7217145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl7317145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl7417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl7517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl7617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:general;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl7717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl7817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl7917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:13.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:right;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8117145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8217145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8317145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl8717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl8817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"_\(* \#\,\#\#0_\)\;_\(* \\\(\#\,\#\#0\\\)\;_\(* \0022-\0022??_\)\;_\(\@_\)";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl8917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"_\(* \#\,\#\#0\.00_\)\;_\(* \\\(\#\,\#\#0\.00\\\)\;_\(* \0022-\0022??_\)\;_\(\@_\)";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl9017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"_\(* \#\,\#\#0\.00_\)\;_\(* \\\(\#\,\#\#0\.00\\\)\;_\(* \0022-\0022??_\)\;_\(\@_\)";
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl9117145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"_\(* \#\,\#\#0\.00_\)\;_\(* \\\(\#\,\#\#0\.00\\\)\;_\(* \0022-\0022??_\)\;_\(\@_\)";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl9217145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl9317145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl9417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl9517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:right;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl9617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:right;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl9717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl9817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl9917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl10017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl10117145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl10217145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl10317145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:right;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl10417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl10517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl10617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl10717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:right;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl10817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl10917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl11017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl11117145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl11217145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl11317145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl11417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl11517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl11617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl11717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl11817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl11917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl12017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl12117145
	{padding:0px;
	mso-ignore:padding;
	color:#C00000;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl12217145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl12317145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl12417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl12517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl12617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:right;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl12717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl12817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl12917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:18.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl13017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:18.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl13117145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl13217145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl13317145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl13417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:13.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl13517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:13.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl13617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:13.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl13717145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl13817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:13.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl13917145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl14017145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl14117145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl14217145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl14317145
	{padding:0px;
	mso-ignore:padding;
	color:#0070C0;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl14417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:18.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl14517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl14617145
	{padding:0px;
	mso-ignore:padding;
	color:#002060;
	font-size:18.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl14717145
	{padding:0px;
	mso-ignore:padding;
	color:#002060;
	font-size:18.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl14817145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:17.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl14917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl15017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl15117145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl15217145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl15317145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl15417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl15517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl15617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl15717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl15817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl15917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl16017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl16117145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl16217145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl16317145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl16417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl16517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl16617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl16717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:right;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl16817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:right;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl16917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl17017145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl17117145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl17217145
	{padding:0px;
	mso-ignore:padding;
	color:red;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl17317145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl17417145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:normal;}
.xl17517145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl17617145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl17717145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;
	mso-text-control:shrinktofit;}
.xl17817145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:italic;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl17917145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl18017145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:13.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
.xl18117145
	{padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:white;
	mso-pattern:black none;
	white-space:nowrap;}
-->
</style>
</head>

<body>
<!--[if !excel]>&nbsp;&nbsp;<![endif]-->
<!--The following information was generated by Microsoft Excel's Publish as Web
Page wizard.-->
<!--If the same item is republished from Excel, all information between the DIV
tags will be replaced.-->
<!----------------------------->
<!--START OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD -->
<!----------------------------->

<div id="HOA DON _YUPOONG_c_17145" align=center x:publishsource="Excel">

<table border=0 cellpadding=0 cellspacing=0 width=742 class=xl6617145
 style='border-collapse:collapse;table-layout:fixed;width:556pt'>
 <col class=xl6617145 width=6 style='mso-width-source:userset;mso-width-alt:
 199;width:4pt'>
 <col class=xl6617145 width=33 style='mso-width-source:userset;mso-width-alt:
 1166;width:25pt'>
 <col class=xl6617145 width=70 style='mso-width-source:userset;mso-width-alt:
 2474;width:52pt'>
 <col class=xl6617145 width=55 style='mso-width-source:userset;mso-width-alt:
 1962;width:41pt'>
 <col class=xl6617145 width=41 style='mso-width-source:userset;mso-width-alt:
 1450;width:31pt'>
 <col class=xl6617145 width=98 style='mso-width-source:userset;mso-width-alt:
 3498;width:74pt'>
 <col class=xl6617145 width=27 style='mso-width-source:userset;mso-width-alt:
 967;width:20pt'>
 <col class=xl6617145 width=78 style='mso-width-source:userset;mso-width-alt:
 2759;width:58pt'>
 <col class=xl6617145 width=56 style='mso-width-source:userset;mso-width-alt:
 1991;width:42pt'>
 <col class=xl6617145 width=34 style='mso-width-source:userset;mso-width-alt:
 1223;width:26pt'>
 <col class=xl6617145 width=49 style='mso-width-source:userset;mso-width-alt:
 1735;width:37pt'>
 <col class=xl6617145 width=42 style='mso-width-source:userset;mso-width-alt:
 1479;width:31pt'>
 <col class=xl6617145 width=13 style='mso-width-source:userset;mso-width-alt:
 455;width:10pt'>
 <col class=xl6617145 width=49 style='mso-width-source:userset;mso-width-alt:
 1735;width:37pt'>
 <col class=xl6617145 width=78 style='mso-width-source:userset;mso-width-alt:
 2759;width:58pt'>
 <col class=xl6617145 width=13 style='mso-width-source:userset;mso-width-alt:
 455;width:10pt'>
 <tr height=18 style='height:13.2pt'>
  <td height=18 class=xl6717145 width=6 style='height:13.2pt;width:4pt'>&nbsp;</td>
  <td class=xl6817145 width=33 style='width:25pt'>&nbsp;</td>
  <td class=xl6817145 width=70 style='width:52pt'>&nbsp;</td>
  <td width=55 style='width:41pt' align=left valign=top><!--[if gte vml 1]><v:shapetype
   id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t"
   path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
   <v:stroke joinstyle="miter"/>
   <v:formulas>
    <v:f eqn="if lineDrawn pixelLineWidth 0"/>
    <v:f eqn="sum @0 1 0"/>
    <v:f eqn="sum 0 0 @1"/>
    <v:f eqn="prod @2 1 2"/>
    <v:f eqn="prod @3 21600 pixelWidth"/>
    <v:f eqn="prod @3 21600 pixelHeight"/>
    <v:f eqn="sum @0 0 1"/>
    <v:f eqn="prod @6 1 2"/>
    <v:f eqn="prod @7 21600 pixelWidth"/>
    <v:f eqn="sum @8 21600 0"/>
    <v:f eqn="prod @7 21600 pixelHeight"/>
    <v:f eqn="sum @10 21600 0"/>
   </v:formulas>
   <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
   <o:lock v:ext="edit" aspectratio="t"/>
  </v:shapetype><v:shape id="Picture_x0020_1" o:spid="_x0000_s4402" type="#_x0000_t75"
   style='position:absolute;margin-left:-.6pt;margin-top:1.8pt;width:62.4pt;
   height:42pt;z-index:1;visibility:visible' o:gfxdata="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
">
   <v:imagedata src="HOA%20DON%20_YUPOONG_c_files/HOA%20DON%20_YUPOONG_c_17145_image001.png"
    o:title=""/>
   <x:ClientData ObjectType="Pict">
    <x:SizeWithCells/>
    <x:CF>Bitmap</x:CF>
    <x:AutoPict/>
   </x:ClientData>
  </v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout;
  position:absolute;z-index:1;margin-left:0px;margin-top:2px;width:83px;
  height:56px'><img width=83 height=56
  src="HOA%20DON%20_YUPOONG_c_files/HOA%20DON%20_YUPOONG_c_17145_image002.png"
  v:shapes="Picture_x0020_1"></span><![endif]><span style='mso-ignore:vglayout2'>
  <table cellpadding=0 cellspacing=0>
   <tr>
    <td height=18 class=xl6817145 width=55 style='height:13.2pt;width:41pt'>&nbsp;</td>
   </tr>
  </table>
  </span></td>
  <td class=xl6817145 width=41 style='width:31pt'>&nbsp;</td>
  <td class=xl6817145 width=98 style='width:74pt'>&nbsp;</td>
  <td class=xl6817145 width=27 style='width:20pt'>&nbsp;</td>
  <td class=xl6817145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl6817145 width=56 style='width:42pt'>&nbsp;</td>
  <td class=xl6817145 width=34 style='width:26pt'>&nbsp;</td>
  <td class=xl6817145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl6817145 width=42 style='width:31pt'>&nbsp;</td>
  <td class=xl6817145 width=13 style='width:10pt'>&nbsp;</td>
  <td class=xl6817145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl6817145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl6917145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr height=37 style='mso-height-source:userset;height:28.05pt'>
  <td height=37 class=xl7017145 style='height:28.05pt'>&nbsp;</td>
  <td colspan=15 class=xl14617145 style='border-right:.5pt solid black'>YUPOONG
  VIETNAM CO., LTD.</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:13.05pt'>
  <td colspan=16 height=17 class=xl15517145 style='border-right:.5pt solid black;
  height:13.05pt'>Lot A 2,3,4, Loteco EPZ, Long Binh Wrd, Bien Hoa City, Dong
  Nai Province</td>
 </tr>
 <tr height=21 style='height:15.6pt'>
  <td height=21 class=xl7017145 style='height:15.6pt'>&nbsp;</td>
  <td colspan=14 class=xl8217145>Tel: 02513-684685<span
  style='mso-spacerun:yes'>      </span>Fax: 02513-684685</td>
  <td class=xl7217145>&nbsp;</td>
 </tr>
 <tr height=33 style='mso-height-source:userset;height:25.05pt'>
  <td height=33 class=xl14417145 style='height:25.05pt'>&nbsp;</td>
  <td class=xl6817145>&nbsp;</td>
  <td class=xl12917145>&nbsp;</td>
  <td class=xl12917145>&nbsp;</td>
  <td colspan=7 class=xl14817145>HÓA &#272;&#416;N BÁN HÀNG (<font
  class="font2417145">INVOICE</font><font class="font2317145">)</font></td>
  <td class=xl12917145>&nbsp;</td>
  <td class=xl12917145>&nbsp;</td>
  <td class=xl12917145>&nbsp;</td>
  <td class=xl12917145>&nbsp;</td>
  <td class=xl13017145>&nbsp;</td>
 </tr>
 <tr height=21 style='height:15.6pt'>
  <td height=21 class=xl7017145 style='height:15.6pt'>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td colspan=7 class=xl8217145>(Dùng cho t&#7893; ch&#7913;c, cá nhân trong
  khu phi thu&#7871; quan)</td>
  <td class=xl7117145 colspan=5 style='border-right:.5pt solid black'>M&#7851;u
  s&#7889; (<font class="font1217145">Form</font><font class="font917145">): </font><font
  class="font817145">07KPTQ0/001</font></td>
 </tr>
 <tr height=18 style='height:13.8pt'>
  <td height=18 class=xl7017145 style='height:13.8pt'>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td colspan=7 class=xl17817145>(HÓA &#272;&#416;N CHUY&#7874;N &#272;&#7892;I
  T&#7914; HÓA &#272;&#416;N &#272;I&#7878;N T&#7916;)<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl7117145 colspan=4>Ký hi&#7879;u (<font class="font1217145">Serial</font><font
  class="font917145">):</font><font class="font817145"> YP/19E</font></td>
  <td class=xl7217145>&nbsp;</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td colspan=7 class=xl14917145 width=383 style='width:288pt'>Ngày <font
  class="font1217145">(Date)</font><font class="font1117145"> </font><font
  class="font717145">28 tháng </font><font class="font1217145">(month)</font><font
  class="font717145"> 02 n&#259;m </font><font class="font1217145">(year)</font><font
  class="font717145"> 2019</font></td>
  <td class=xl7117145 colspan=4>S&#7889; (<font class="font1217145">No</font><font
  class="font917145">.):</font><font class="font817145"><span
  style='mso-spacerun:yes'>      </span></font><font class="font1417145"><span
  style='mso-spacerun:yes'> </span></font><font class="font1617145">0000000</font></td>
  <td class=xl7417145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr class=xl6817145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl6717145 style='height:18.0pt'>&nbsp;</td>
  <td class=xl17917145 colspan=10>&#272;&#417;n v&#7883; bán hàng (<font
  class="font1217145">Company name</font><font class="font717145">): </font><font
  class="font517145">CÔNG TY TNHH YUPOONG VI&#7878;T NAM</font></td>
  <td class=xl6817145>&nbsp;</td>
  <td class=xl6817145>&nbsp;</td>
  <td class=xl6817145>&nbsp;</td>
  <td class=xl6817145>&nbsp;</td>
  <td class=xl6917145>&nbsp;</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>
  <td class=xl7517145 colspan=14>&#272;&#7883;a ch&#7881; (<font
  class="font1217145">Address</font><font class="font717145">): </font><font
  class="font2117145">Lô A2,3,4 Khu công nghi&#7879;p Long Bình,
  Ph&#432;&#7901;ng Long Bình, Thành ph&#7889; Biên Hóa, T&#7881;nh
  &#272;&#7891;ng Nai</font></td>
  <td class=xl7217145>&nbsp;</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>
  <td class=xl7517145 colspan=15 style='border-right:.5pt solid black'>S&#7889;
  tài kho&#7843;n (<font class="font1217145">Acc. code</font><font
  class="font717145">): 0121370472272 (USD)/ 0121000472262 (VND) Ngân hàng
  Vietcombank - CN &#272;&#7891;ng Nai</font></td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>
  <td class=xl7517145 colspan=7>&#272;i&#7879;n tho&#7841;i (<font
  class="font1217145">Tel</font><font class="font717145">): 02513-684685 - Fax:
  02513-684685</font></td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl7517145 colspan=7 style='border-right:.5pt solid black'>Mã
  s&#7889; thu&#7871; (<font class="font1217145">Tax code</font><font
  class="font717145">): </font><font class="font1517145">3 6 0 0 5 6 3 4 0 1</font></td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl6717145 style='height:18.0pt'>&nbsp;</td>
  <td class=xl17917145 colspan=5>H&#7885; tên ng&#432;&#7901;i mua hàng (<font
  class="font1217145">Customer's name</font><font class="font717145">):</font></td>
  <td colspan=10 class=xl17617145 style='border-right:.5pt solid black'>&nbsp;</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>
  <td class=xl7517145 colspan=4>Tên &#273;&#417;n v&#7883; (<font
  class="font1217145">Company's name</font><font class="font717145">):</font></td>
  <td colspan=10 class=xl15717145>CÔNG TY TNHH YP LONG AN</td>
  <td class=xl14517145>&nbsp;</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>
  <td class=xl7517145 colspan=13>&#272;&#7883;a ch&#7881; (<font
  class="font1217145">Address</font><font class="font717145">): Lô C3,
  &#272;&#432;&#7901;ng s&#7889; 8, KCN Hòa Bình, Xã Nh&#7883; Thành, H.
  Th&#7911; Th&#7915;a, T&#7881;nh Long An.</font></td>
  <td class=xl7717145>&nbsp;</td>
  <td class=xl7817145>&nbsp;</td>
 </tr>
 <tr height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl7017145 style='height:18.0pt'>&nbsp;</td>
  <td class=xl7517145 colspan=4>S&#7889; tài kho&#7843;n (<font
  class="font1217145">Account code</font><font class="font717145">):</font></td>
  <td class=xl7917145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl7217145>&nbsp;</td>
 </tr>
 <tr height=26 style='mso-height-source:userset;height:19.95pt'>
  <td height=26 class=xl7017145 style='height:19.95pt'>&nbsp;</td>
  <td class=xl7517145 colspan=6>Hình th&#7913;c thanh toán (<font
  class="font1217145">Mod of payment</font><font class="font717145">): </font><font
  class="font517145">TM/CK</font></td>
  <td class=xl7517145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145 colspan=6><font class="font717145">Mã s&#7889; thu&#7871;</font><font
  class="font1017145"> (</font><font class="font1217145">Tax code</font><font
  class="font1017145">): </font><font class="font1517145">1 1 0 1 7 8 6 0 4 6</font></td>
  <td class=xl7217145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=21 style='height:15.6pt'>
  <td colspan=2 height=21 class=xl15017145 style='height:15.6pt'>STT</td>
  <td colspan=5 class=xl15017145 style='border-right:.5pt solid black'>Tên hàng
  hóa, d&#7883;ch v&#7909;</td>
  <td class=xl8117145>&#272;&#417;n v&#7883; tính</td>
  <td colspan=2 class=xl15017145 style='border-right:.5pt solid black'>S&#7889;
  l&#432;&#7907;ng</td>
  <td colspan=3 class=xl8117145>&#272;&#417;n giá</td>
  <td colspan=3 class=xl15017145 style='border-right:.5pt solid black'>Thành
  ti&#7873;n</td>
 </tr>
 <tr class=xl8317145 height=18 style='height:13.2pt'>
  <td colspan=2 height=18 class=xl16317145 style='height:13.2pt'>No.</td>
  <td colspan=5 class=xl15217145 style='border-right:.5pt solid black'>Description
  of goods</td>
  <td class=xl8317145>Unit</td>
  <td colspan=2 class=xl16317145 style='border-right:.5pt solid black'>Quatity</td>
  <td colspan=3 class=xl8317145>Unit price</td>
  <td colspan=3 class=xl16317145 style='border-right:.5pt solid black'>Amount</td>
 </tr>
 <tr class=xl8517145 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl15817145 style='height:15.0pt'>1</td>
  <td colspan=5 height=20 width=291 style='border-right:.5pt solid black;
  height:15.0pt;width:218pt' align=left valign=top><!--[if gte vml 1]><v:shape
   id="Picture_x0020_3" o:spid="_x0000_s4404" type="#_x0000_t75" style='position:absolute;
   margin-left:79.2pt;margin-top:13.2pt;width:346.8pt;height:234.6pt;z-index:3;
   visibility:visible' o:gfxdata="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
">
   <v:imagedata src="HOA%20DON%20_YUPOONG_c_files/HOA%20DON%20_YUPOONG_c_17145_image003.png"
    o:title=""/>
   <x:ClientData ObjectType="Pict">
    <x:SizeWithCells/>
    <x:CF>Bitmap</x:CF>
    <x:AutoPict/>
   </x:ClientData>
  </v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout;
  position:absolute;z-index:3;margin-left:106px;margin-top:18px;width:462px;
  height:313px'><img width=462 height=313
  src="HOA%20DON%20_YUPOONG_c_files/HOA%20DON%20_YUPOONG_c_17145_image004.png"
  v:shapes="Picture_x0020_3"></span><![endif]><span style='mso-ignore:vglayout2'>
  <table cellpadding=0 cellspacing=0>
   <tr>
    <td colspan=5 height=20 class=xl15817145 width=291 style='border-right:
    .5pt solid black;height:15.0pt;width:218pt'>2</td>
   </tr>
  </table>
  </span></td>
  <td class=xl8417145>3</td>
  <td colspan=2 class=xl15817145 style='border-right:.5pt solid black'>4</td>
  <td colspan=3 class=xl8417145>5</td>
  <td colspan=3 class=xl15817145 style='border-right:.5pt solid black'>6</td>
 </tr>
 <tr class=xl9317145 height=22 style='mso-height-source:userset;height:17.1pt'>
  <td colspan=2 height=22 class=xl9217145 style='border-right:.5pt solid black;
  height:17.1pt'>&nbsp;</td>
  <td colspan=5 class=xl16017145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>Nguyên v&#7853;t li&#7879;u</td>
  <td class=xl8617145 style='border-left:none'>&nbsp;</td>
  <td class=xl8717145 style='border-left:none'>&nbsp;</td>
  <td class=xl8817145><span style='mso-spacerun:yes'>   </span></td>
  <td class=xl8917145 style='border-left:none'><span
  style='mso-spacerun:yes'>   </span></td>
  <td class=xl9017145>&nbsp;</td>
  <td class=xl9117145>&nbsp;</td>
  <td class=xl9217145 style='border-left:none'>&nbsp;</td>
  <td class=xl9317145>&nbsp;</td>
  <td class=xl9117145><span style='mso-spacerun:yes'>   </span></td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>1</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>NEDDLE</td>
  <td class=xl9417145 style='border-left:none'>PCS</td>
  <td class=xl9517145 style='border-left:none'>500</td>
  <td class=xl9617145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>0.0838</td>
  <td class=xl9717145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>41.90</td>
  <td class=xl9717145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>2</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>NEDDLE</td>
  <td class=xl9417145 style='border-left:none'>PCS</td>
  <td class=xl9517145 style='border-left:none'>500</td>
  <td class=xl9617145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>0.0838</td>
  <td class=xl9717145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>41.90</td>
  <td class=xl9717145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>&nbsp;</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>&nbsp;</td>
  <td class=xl9417145 style='border-left:none'>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>&nbsp;</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>&nbsp;</td>
  <td class=xl9417145 style='border-left:none'>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>&nbsp;</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>&nbsp;</td>
  <td class=xl9417145 style='border-left:none'>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>&nbsp;</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>&nbsp;</td>
  <td class=xl9417145 style='border-left:none'>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>&nbsp;</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>&nbsp;</td>
  <td class=xl9817145 style='border-left:none'>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>&nbsp;</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>&nbsp;</td>
  <td class=xl9817145 style='border-left:none'>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9717145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>&nbsp;</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>&nbsp;</td>
  <td class=xl9817145 style='border-left:none'>&nbsp;</td>
  <td class=xl10017145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9917145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl12317145 width=6 style='height:18.0pt;width:4pt'>&nbsp;</td>
  <td class=xl12417145 width=33 style='width:25pt'>&nbsp;</td>
  <td class=xl13117145 width=70 style='border-left:none;width:52pt'>&nbsp;</td>
  <td class=xl12517145 width=55 style='width:41pt'>&nbsp;</td>
  <td class=xl12517145 width=41 style='width:31pt'>&nbsp;</td>
  <td class=xl12517145 width=98 style='width:74pt'>&nbsp;</td>
  <td class=xl13217145 width=27 style='width:20pt'>&nbsp;</td>
  <td class=xl9817145 style='border-left:none'>&nbsp;</td>
  <td class=xl10017145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9917145>&nbsp;</td>
  <td class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl12617145>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
 </tr>
 <tr class=xl8217145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>&nbsp;</td>
  <td colspan=5 class=xl13117145 width=291 style='border-right:.5pt solid black;
  border-left:none;width:218pt'>&nbsp;</td>
  <td class=xl9817145 style='border-left:none'>&nbsp;</td>
  <td class=xl10017145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9917145>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl9617145>&nbsp;</td>
 </tr>
 <tr class=xl7517145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td colspan=2 height=24 class=xl12317145 width=39 style='border-right:.5pt solid black;
  height:18.0pt;width:29pt'>&nbsp;</td>
  <td class=xl11517145 width=70 style='width:52pt'>&nbsp;</td>
  <td class=xl11517145 width=55 style='width:41pt'>&nbsp;</td>
  <td colspan=3 class=xl12517145 width=166 style='width:125pt'>&nbsp;</td>
  <td class=xl10117145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl10217145 width=56 style='border-left:none;width:42pt'>&nbsp;</td>
  <td class=xl10317145 width=34 style='width:26pt'>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl10417145 width=13 style='width:10pt'>&nbsp;</td>
  <td colspan=2 class=xl9517145 style='border-left:none'>&nbsp;</td>
  <td class=xl10517145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr class=xl7517145 height=29 style='mso-height-source:userset;height:22.05pt'>
  <td height=29 class=xl12717145 width=6 style='height:22.05pt;width:4pt'>&nbsp;</td>
  <td colspan=11 class=xl16917145 width=583 style='width:437pt'>C&#7897;ng
  ti&#7873;n bán hàng hóa, d&#7883;ch v&#7909; (<font class="font1217145">Total
  amount</font><font class="font917145">):</font></td>
  <td class=xl12817145 width=13 style='width:10pt'>&nbsp;</td>
  <td colspan=2 class=xl16717145 style='border-left:none'>83.80</td>
  <td class=xl10717145>&nbsp;</td>
 </tr>
 <tr class=xl7517145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl10617145 style='height:18.0pt;border-top:none'>&nbsp;</td>
  <td colspan=14 class=xl17517145>S&#7889; ti&#7873;n vi&#7871;t b&#7857;ng
  ch&#7919; (<font class="font1217145">In words</font><font class="font717145">):
  </font><font class="font1317145">Tám m&#432;&#417;n ba &#273;ôla tám
  m&#432;&#417;i cent./.</font></td>
  <td class=xl10817145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr height=18 style='mso-height-source:userset;height:13.95pt'>
  <td height=18 class=xl8017145 style='height:13.95pt'>&nbsp;</td>
  <td class=xl10917145>&nbsp;</td>
  <td class=xl10917145>&nbsp;</td>
  <td class=xl10917145>&nbsp;</td>
  <td class=xl11017145 width=41 style='width:31pt'>&nbsp;</td>
  <td class=xl11017145 width=98 style='width:74pt'>&nbsp;</td>
  <td class=xl11017145 width=27 style='width:20pt'>&nbsp;</td>
  <td class=xl11017145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl11017145 width=56 style='width:42pt'>&nbsp;</td>
  <td class=xl11017145 width=34 style='width:26pt'>&nbsp;</td>
  <td class=xl11017145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl11017145 width=42 style='width:31pt'>&nbsp;</td>
  <td class=xl11017145 width=13 style='width:10pt'>&nbsp;</td>
  <td class=xl11017145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl11017145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl11117145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr class=xl7717145 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 class=xl11317145 style='height:18.0pt'>&nbsp;</td>
  <td colspan=4 class=xl17417145 width=199 style='width:149pt'>Converter</td>
  <td colspan=3 class=xl17417145 width=203 style='width:152pt'>Ng&#432;&#7901;i
  mua hàng (Buyer)</td>
  <td colspan=7 class=xl17417145 width=321 style='width:241pt'>Ng&#432;&#7901;i
  mua hàng (Seller)</td>
  <td class=xl11217145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr class=xl7117145 height=18 style='mso-height-source:userset;height:13.95pt'>
  <td height=18 class=xl11417145 style='height:13.95pt'>&nbsp;</td>
  <td colspan=4 class=xl11517145 width=199 style='width:149pt'>(Ký, ghi rõ
  h&#7885; tên)</td>
  <td colspan=3 class=xl11517145 width=203 style='width:152pt'>(Ký, ghi rõ
  h&#7885; tên)</td>
  <td colspan=7 class=xl11517145 width=321 style='width:241pt'>(Ký, &#273;óng
  d&#7845;u, ghi rõ h&#7885; tên)</td>
  <td class=xl11617145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr class=xl11917145 height=20 style='mso-height-source:userset;height:15.45pt'>
  <td height=20 class=xl11717145 style='height:15.45pt'>&nbsp;</td>
  <td colspan=4 class=xl17317145 width=199 style='width:149pt'>(Signature &amp;
  full name)</td>
  <td colspan=3 class=xl17317145 width=203 style='width:152pt'>(Signature &amp;
  full name)</td>
  <td colspan=7 class=xl17317145 width=321 style='width:241pt'>(Signature,
  stamp &amp; full name)</td>
  <td class=xl11817145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.45pt'>
  <td height=20 class=xl7017145 style='height:15.45pt'>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>
  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>
  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>
  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl7317145 width=56 style='width:42pt'>&nbsp;</td>
  <td class=xl7317145 width=34 style='width:26pt'>&nbsp;</td>
  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl7317145 width=42 style='width:31pt'>&nbsp;</td>
  <td class=xl7317145 width=13 style='width:10pt'>&nbsp;</td>
  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl12017145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr height=22 style='mso-height-source:userset;height:16.5pt'>
  <td height=22 class=xl7017145 style='height:16.5pt'>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>
  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>
  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>
  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl13317145 colspan=3>Signature Valid</td>
  <td align=left valign=top><!--[if gte vml 1]><v:shape id="Picture_x0020_8"
   o:spid="_x0000_s4403" type="#_x0000_t75" style='position:absolute;
   margin-left:13.2pt;margin-top:5.4pt;width:61.2pt;height:40.2pt;z-index:2;
   visibility:visible' o:gfxdata="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
">
   <v:imagedata src="HOA%20DON%20_YUPOONG_c_files/HOA%20DON%20_YUPOONG_c_17145_image005.png"
    o:title=""/>
   <x:ClientData ObjectType="Pict">
    <x:SizeWithCells/>
    <x:CF>Bitmap</x:CF>
    <x:AutoPict/>
   </x:ClientData>
  </v:shape><![endif]--><![if !vml]><span style='mso-ignore:vglayout;
  position:absolute;z-index:2;margin-left:18px;margin-top:7px;width:82px;
  height:54px'><img width=82 height=54
  src="HOA%20DON%20_YUPOONG_c_files/HOA%20DON%20_YUPOONG_c_17145_image006.png"
  v:shapes="Picture_x0020_8"></span><![endif]><span style='mso-ignore:vglayout2'>
  <table cellpadding=0 cellspacing=0>
   <tr>
    <td height=22 class=xl13417145 width=42 style='height:16.5pt;width:31pt'>&nbsp;</td>
   </tr>
  </table>
  </span></td>
  <td class=xl13517145>&nbsp;</td>
  <td class=xl13517145>&nbsp;</td>
  <td class=xl13617145>&nbsp;</td>
  <td class=xl13817145>&nbsp;</td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.45pt'>
  <td height=20 class=xl7017145 style='height:15.45pt'>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>
  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>
  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>
  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>
  <td colspan=7 class=xl17017145 style='border-right:.5pt solid black'><font
  class="font1717145">&#272;&#432;&#7907;c ký b&#7903;i:</font><font
  class="font1817145"> </font><font class="font1917145">CÔNG TY TNHH YUPOONG
  VI&#7878;T NAM</font></td>
  <td class=xl13917145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr height=21 style='height:15.6pt'>
  <td height=21 class=xl7017145 style='height:15.6pt'>&nbsp;</td>
  <td class=xl6517145>Ngày <span style='display:none'>chuy&#7875;n
  &#273;&#7893;i :</span></td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl6617145>&nbsp;</td>
  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>
  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>
  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>
  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl14017145 colspan=3>Ngày Ký: <font class="font2017145">01/03/2019</font></td>
  <td class=xl14117145>&nbsp;</td>
  <td class=xl14117145>&nbsp;</td>
  <td class=xl14117145>&nbsp;</td>
  <td class=xl14217145>&nbsp;</td>
  <td class=xl13717145>&nbsp;</td>
 </tr>
 <tr height=18 style='mso-height-source:userset;height:13.95pt'>
  <td height=18 class=xl7017145 style='height:13.95pt'>&nbsp;</td>
  <td class=xl14317145 colspan=3>Mã nh&#7853;n hóa &#273;&#417;n:</td>
  <td class=xl7317145 width=41 style='width:31pt'>&nbsp;</td>
  <td class=xl7317145 width=98 style='width:74pt'>&nbsp;</td>
  <td class=xl7317145 width=27 style='width:20pt'>&nbsp;</td>
  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl7317145 width=56 style='width:42pt'>&nbsp;</td>
  <td class=xl7317145 width=34 style='width:26pt'>&nbsp;</td>
  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl7317145 width=42 style='width:31pt'>&nbsp;</td>
  <td class=xl7317145 width=13 style='width:10pt'>&nbsp;</td>
  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl12017145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr height=18 style='mso-height-source:userset;height:13.95pt'>
  <td height=18 class=xl7017145 style='height:13.95pt'>&nbsp;</td>
  <td class=xl6617145 colspan=7>Tra c&#7913;u t&#7841;i Website: <font
  class="font617145"><span style='mso-spacerun:yes'> </span></font><font
  class="font2217145">http://genuwinsolution.com/e-invoice</font></td>
  <td class=xl7317145 width=56 style='width:42pt'>&nbsp;</td>
  <td class=xl7317145 width=34 style='width:26pt'>&nbsp;</td>
  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl7317145 width=42 style='width:31pt'>&nbsp;</td>
  <td class=xl7317145 width=13 style='width:10pt'>&nbsp;</td>
  <td class=xl7317145 width=49 style='width:37pt'>&nbsp;</td>
  <td class=xl7317145 width=78 style='width:58pt'>&nbsp;</td>
  <td class=xl12017145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=xl7017145 style='height:12.0pt'>&nbsp;</td>
  <td colspan=14 class=xl16617145>(C&#7847;n ki&#7875;m tra, &#273;&#7889;i
  chi&#7871;u khi l&#7853;p, giao nh&#7853;n hóa &#273;&#417;n)</td>
  <td class=xl12017145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <tr height=16 style='mso-height-source:userset;height:12.0pt'>
  <td height=16 class=xl8017145 style='height:12.0pt'>&nbsp;</td>
  <td colspan=14 class=xl15317145>(In t&#7841;i ph&#7847;n m&#7873;m Genuwin
  E-INVOICE c&#7911;a Công ty TNHH MTV GENUWIN SOLUTION - MST: 1201496252)</td>
  <td class=xl12217145 width=13 style='width:10pt'>&nbsp;</td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=6 style='width:4pt'></td>
  <td width=33 style='width:25pt'></td>
  <td width=70 style='width:52pt'></td>
  <td width=55 style='width:41pt'></td>
  <td width=41 style='width:31pt'></td>
  <td width=98 style='width:74pt'></td>
  <td width=27 style='width:20pt'></td>
  <td width=78 style='width:58pt'></td>
  <td width=56 style='width:42pt'></td>
  <td width=34 style='width:26pt'></td>
  <td width=49 style='width:37pt'></td>
  <td width=42 style='width:31pt'></td>
  <td width=13 style='width:10pt'></td>
  <td width=49 style='width:37pt'></td>
  <td width=78 style='width:58pt'></td>
  <td width=13 style='width:10pt'></td>
 </tr>
 <![endif]>
</table>

</div>


<!----------------------------->
<!--END OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD-->
<!----------------------------->
</body>

</html>
