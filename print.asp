
<HTML>
<HEAD>

<title>Printing Web pages and Page Break Styles </title>
<meta name="keywords" content="printing web pages, page breaks, client side web page printing with page break">
<meta name="description" content="Printing web pages with page breaks.">
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">

<STYLE>P.page { page-break-after: always }</STYLE>

</HEAD>

<body onbeforeprint="beforeprint()"
onafterprint="afterprint()"
  bgcolor="infobackground">
<p><input disabled name="idPrint" type="button"  value="Print this page" onclick="print()"> </p>
</body>
<font face='verdana' size=2>
<script defer>
function window.onload() {
  idPrint.disabled = false;
}

var originalTitle;
function beforeprint() {
  idPrint.disabled = true;
  originalTitle = document.title;
  document.title = originalTitle + " - Printing Web Pages";
}

function afterprint() {
  document.title = originalTitle;
  idPrint.disabled = false;
}
</script>

<p class=page>
<b><font size=3>Printing web pages with page breaks Using style tags and IEs onbeforeprint and onafterprint</b>
Because this is a printing demo there is no other elements on this page.
Click the "print This Page Button" to demonstrate how this document would print
using page breaks. This paragraph will print as page one.
</p>

<P class=page>

And by using any of these tags :BLOCKQUOTE, BODY, CENTER, DD, DIR, DIV, DL, DT, FIELDSET, FORM, Hn, LI, LISTING, 
MARQUEE, MENU, OL, P, PLAINTEXT, PRE, UL, and XMP you can implement a page break and have this paragraph
print on page two. View the source to see this code, which is entirely Client Side Scripting
</p>
       

</BODY>
</HTML>