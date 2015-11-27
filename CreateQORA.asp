<%@Language = VBscript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<html>

<script language="javascript">
function Populate(val)
{
  document.all.T1.value = document.all.T1.value + val
}

// function to ask about the confirmation of the file.
function ConfirmChoice() 
{ 
       answer = confirm("Are you sure, do you want to save ?")
		if (answer == true) 
		{ 
           return ;
		} 
		else
		 { 
		 return (false);
		}  
}
// function to reload the form to add the new entries
function FillDetails()
{
 document.Menu.submit();

}
</script>
<%
Dim connFac
Dim rsFillFac
Dim strSQL

'Database Connectivity Code 
  set connFac = Server.CreateObject("ADODB.Connection")
  connFac.open constr
 
 ' setting up the recordset
 
   strSQL ="Select * from tblFaculty"
   set rsFillFac = Server.CreateObject("ADODB.Recordset")
   rsFillFac.Open strSQL, connFac, 3, 3
%>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Search Risk Assesment</title>

<base target="_self">

</head>

<body link="#000000" vlink="#000000" alink="#000000">
    <!--#include file="HeaderMenu.asp" -->
<table border="0" width="750" style="border-collapse: collapse" id="table3" height="51">
	<tr>
		<td width="20" height="30">
		<img border="0" src="uts_logo.gif"></td>
		<td height="30" width="368">Online Risk Management Database 
		</td>
		<td width="356" height="30">
		<p align="right"><b>&nbsp;&nbsp;&nbsp; Help
		 </b></td>
	</tr>
	<tr>
		<td width="20" height="21">
		&nbsp;</td>
		<td width="368" height="21">
		<p align="left">&nbsp;</td>
		<td width="356" height="21">
		&nbsp;</td>
	</tr>
</table>

	<table border="1" cellspacing="1" width="100%" id="AutoNumber1" bordercolor="#FFFFFF" height="28">
      <tr>
        <td width="100%" colspan="6" height="16">
        <font color="#660033" face="Tahoma" size="2"><b>C</b>reate
        <b>Q</b>uick and <b>O</b>bvious <b>R</b>isk <b>A</b>ssessment <b>F</b>orm</span></td>
      </tr>
      <tr>
        <form method="POST" action="index.asp">
          
          
          <td width="6%" height="11">
          
          &nbsp;</td>
          <td width="16%" height="11">
          
          &nbsp;</td>
          
                 <td width="5%" height="11">
          &nbsp;</td>
          
          <td width="23%" height="11">
          &nbsp;</td>
          
        </form>
 <%'**************************** end of form 1*************************************************************%>       
 <%'**************************** Start of form 2*************************************************************%>       

 <form method="Post" action="index.asp">
   
          
          <td width="10%" height="11">
          &nbsp;</td>
          
          
          Room / Name</span></td>
          <td width="14%" height="11">
          &nbsp;</td>
       
      </tr>
</table>
</form>
<table border="1" cellspacing="1" bordercolor="#FFFFFF" width="100%" id="AutoNumber2">
  <tr>
    <td width="18%" bordercolor="#FFFFFF" bgcolor="#C0C0C0">
    Supervisor Name :</span></td>
    <td width="21%" bordercolor="#FFFFFF" bgcolor="#C0C0C0">
   
      &nbsp;</td>
    <td width="13%" bordercolor="#FFFFFF" bgcolor="#C0C0C0">
   
      &nbsp;Faculty / Unit:</span></td>
    <td width="48%" bordercolor="#FFFFFF" bgcolor="#C0C0C0">
   
      &nbsp;</td>
  </tr>
  <tr>
    <td width="18%" bordercolor="#FFFFFF" bgcolor="#C0C0C0">
    Assessor</span></td>
    <td width="82%" bordercolor="#FFFFFF" colspan="3">
   
      <input type="text" name="T1" size="20">
  
&nbsp;</td>
  </tr>
  <tr>
    <td width="18%" bordercolor="#FFFFFF" bgcolor="#C0C0C0">
    Task Description</span></td>
    <td width="82%" bordercolor="#FFFFFF" colspan="3">
    
      <input type="text" name="T1" size="84">

&nbsp;</td>
  </tr>
</table>


<form method="POST" action="CreateQORA.asp" webbot-action="--WEBBOT-SELF--">
  <!--webbot bot="SaveResults" u-file="../_private/form_results.csv" s-format="TEXT/CSV" s-label-fields="TRUE" startspan --><input TYPE="hidden" NAME="VTI-GROUP" VALUE="0"><!--webbot bot="SaveResults" i-checksum="43374" endspan --><table border="1" cellspacing="1" width="100%" id="AutoNumber3" bordercolor="#FFFFFF" height="325">
  <tr>
    <td width="112%" height="27" colspan="9"><b>
    Select Hazards from the 
    table below</span></b></td>
  </tr>
  <tr>
    <td width="5%" height="27" align="center"><b>
    Sr No</span></b></td>
    <td width="13%" height="27" align="center"><b>
    Biological</span></b></td>
    <td width="14%" height="27" align="center"><b>
    Plant</b></span></td>
    <td width="14%" height="27" align="center"><b>
    Working Environment</b></span></td>
    <td width="12%" height="27" align="center"><b>
    
    Ergonomic/Manual Handling</span></b></td>
    <td width="12%" height="27" align="center"><b>
    Chemical</b></span></td>
    <td width="8%" height="27" align="center"><b>
    Electrical</b></span></td>
    <td width="7%" height="27" align="center"><b>
    Radiation</b></span></td>
    <td width="34%" rowspan="13" height="320">
        
        <p align="left"><b>&nbsp;&nbsp;&nbsp; 
        Edit Text</span></b><br>
        
        <textarea rows="10" name="T1" cols="10"></textarea>

      &nbsp;</td>
  </tr>
  <tr>
    <td>
    1</span></td>
    <td width="13%" height="24">
    <a href="#" onclick="Populate('Imported Biomaterials\r\n')">Imported Biomaterials </a></td>
    <td width="14%" height="24">
    Noise</span></td>
    <td width="14%" height="24">
    Extremes in Temperature</td>
    <td width="12%" height="24">
    Repetitive Movements</td>
    <td width="12%" height="24">
    Hazardous Substances</td>
    <td width="8%" height="24">
    Plug-In Equipment</td>
    <td width="7%" height="24">
    Ionizing Radiation</td>
  </tr>
  <tr>
    <td>
    2</span></td>
    <td width="13%" height="19">
    Cytotoxins</td>
    <td width="14%" height="19">
    Vibration</td>
    <td width="14%" height="19">
    Confined Space</td>
    <td width="12%" height="19">
    Lifting Awkwardly</td>
    <td width="12%" height="19">
    Dangerous Goods</td>
    <td width="8%" height="19">
    Exposed Conductors</td>
    <td width="7%" height="19">
    Non-Ionising Radiation</td>
  </tr>
  <tr>
    <td>
    3</span></td>
    <td width="13%" height="19">
    Pathogens</td>
    <td width="14%" height="19">
    Moving Parts (Crushing, Friction, Stab, Cut, Shear)</td>
    <td width="14%" height="19">
    Height</td>
    <td width="12%" height="19">
    Lifting Heavy Objects</td>
    <td width="12%" height="19">
    Hazardous Waste</td>
    <td width="8%" height="19">
    High Voltage</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td>
    4</span></td>
    <td width="13%" height="19">
    Genetically Modified Organisms</td>
    <td width="14%" height="19">
    Pressure Vessels and Boilers</td>
    <td width="14%" height="19">
    Isolation</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="12%" height="19">
    Fumes</td>
    <td width="8%" height="19">
    Electrical Wiring</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td>
    5</span></td>
    <td width="13%" height="19">
    Communicable Diseases</td>
    <td width="14%" height="19">
    Compressed Gas</td>
    <td width="14%" height="19">
    Slip and Trip Hazards</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="12%" height="19">
    Dust</td>
    <td width="8%" height="19">&nbsp;</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td>
    6</span></td>
    <td width="13%" height="19">
    Animals bites and scatches</td>
    <td width="14%" height="19">
    Lifts</td>
    <td width="14%" height="19">
    Fieldwork</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="12%" height="19">
    Vapours</td>
    <td width="8%" height="19">&nbsp;</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td>
    7</span></td>
    <td width="13%" height="19">
    Allergies to Animal Bedding, Dander and Fluids</td>
    <td width="14%" height="19">
    Hoists</td>
    <td width="14%" height="19">
    Working in Remote Locations</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="12%" height="19">
    Gases</td>
    <td width="8%" height="19">&nbsp;</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td>
    8</span></td>
    <td width="13%" height="19">
    Working With Insects</td>
    <td width="14%" height="19">
    Cranes</td>
    <td width="14%" height="19">
    Working Outdoors</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="12%" height="19">
    Fire</td>
    <td width="8%" height="19">&nbsp;</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td>
    9</span></td>
    <td width="13%" height="19">
    Working With Fungi</td>
    <td width="14%" height="19">
    Sharps</td>
    <td width="14%" height="19">&nbsp;</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="12%" height="19">
    Explosion</td>
    <td width="8%" height="19">&nbsp;</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td>
    10</span></td>
    <td width="13%" height="19">
    Working With Bacteria</td>
    <td width="14%" height="19">
    Needles</td>
    <td width="14%" height="19">&nbsp;</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="8%" height="19">&nbsp;</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td>
    11</span></td>
    <td width="13%" height="19">
    Working With Viruses</td>
    <td width="14%" height="19">&nbsp;</td>
    <td width="14%" height="19">&nbsp;</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="8%" height="19">&nbsp;</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  <tr>
    <td>
    12</span></td>
    <td width="13%" height="19">
    Infectious Materials</td>
    <td width="14%" height="19">&nbsp;</td>
    <td width="14%" height="19">&nbsp;</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="12%" height="19">&nbsp;</td>
    <td width="8%" height="19">&nbsp;</td>
    <td width="7%" height="19">&nbsp;</td>
  </tr>
  </table>

  <p>&nbsp;</p>
</form>
</body>
</html>