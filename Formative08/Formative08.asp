<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>MIS342Formative08</title>
    <style type="text/css">
        .auto-style3 {
            width: 1114px;
            height: 642px;
        }
        .auto-style4 {
            color: #0000FF;
        }
    </style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW"> 
<table width="100%" border="8" align="center" cellpadding="8" cellspacing="8" bordercolor="#FFFFFF" bgcolor="#0000CC">
  <tr>
    <td bgcolor="#0000FF"><h1 align="center"><font color="#FFFFFF"><em><strong>MIS
              342&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Formative</strong></em></font></h1></td>
  </tr>
</table>


  <table width="100%" border="1" cellspacing="1" cellpadding="1">
    <tr bgcolor="#00FFFF">
      <td width="36%">Email Address:
      
        <!-- Enter your winona.edu email address in the VALUE field below   GROK --> 
        <input type="text" name="email" id="email" size="30" maxlength="50" value=""/><br />
      a valid  email address is required</td>
       
      <td width="40%">First Name:
      
        <!-- Enter your first name in the VALUE field below   GROK --> 
      <input type="text" name="FirstName" size="30" maxlength="50" value="EnterYour FirstName"/></td>
      
      <td width="24%">Last Name:
      
        <!-- Enter your last name in the VALUE field below   GROK --> 
      <input type="text" name="LastName" size="25" maxlength="50" value="EnterYourLastName"/></td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td width="36%">Semester:
        <input type="text" name="Semester" value=<% response.write(semester)%> />
      </td>
      <td>Class:
        <input type="text" name="Class" value=<% response.write(xClass)%> />
      </td>
      <td><input name="InstID" type="hidden" id="InstID" value="00617282" />
      
        StarID:
        
        <!-- Enter your StarID in the VALUE field below   GROK --> 
      <input type="text" name="PIN" value="EnterYourStarID" /></td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td width="36%">Section:  <input type="text" name="Section" value="01"/>
      </td>
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Formative08"/>
      </td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr bgcolor="#FF0000">
      <td colspan="3"><div align="center">
          <input type="submit" name="Submit" value="Submit" />
        </div></td>
    </tr>
  </table>


<p><font color="#FF0000"><b>300 points</b></font></p>
    <p>&nbsp;</p>
    <p>Note: 
  Save this  assignment to your website.<br />
    </p>
    <p>Complete your personal information in the above section.<br />
    </p>
 <h2 align="center">    <font color="#FF0000"><b><em>Access
        and Linked Tables! </em></b></font></h2>
 <p>Answer the following questions by copying the contents of the &quot;LinkedTableMadnessProject&quot; folder which is located on the network drive at:&nbsp;&nbsp; </p>
    <%Response.write("\\store\classes\" & YearTerm & courseid & "\ReadOnly\" ) %>  to a location on your computer.<br />
    <p>
    Refer to the PowerPoint &quot;<a href="../../Presentations/Chapter08.ppt">DataIntegrationTips</a>&quot; which is also on the course website 
  </p>
    <p>&nbsp;</p>
    <p>&nbsp;</p>
    <p>You are provided with an Access Database, Linked.accdb, and several files.</p>
    <br />Open Linked.accdb, then open the first table in the list, CATEGORI.
    <br />
    <br />You will be presented with the following error message<br /> 
    <img alt="link error" class="auto-style3" src="LinkedTableError.JPG" /><br />
         <br />
         <br />
 
 <p><font color="#0000FF">R<b>ight click the table CATEGORI, select &quot;Linked Table Manager&quot; from the menu, check CATEGORI, click OK, then&nbsp; find the current location of the file.<br />
      Once you have found the correct file select it, click 'Open' and you will be presented with this message:<br />
     <img src="LinkedTableFound.JPG" /><br />
     
     
     &nbsp;
 </b> </font></p>
    <p><span class="auto-style4">Hint:&nbsp; look carefully at the dialog box settings to find the file type.&nbsp; Or right click the CATEGORI table in the navigation pane and examine the table properties. </span></p>
    <p><font color="#0000FF"><b> 1. (50) Provide the folder and file name of the linked table &quot;CATEGORI&quot;<br />
 </b> </font><b><font color="#000000">
    <input name="q1" type="text" id="q1" size="120" maxlength="120" value="names:" />
<br />
            </font></b></p>
    <p><b>Now use the linked table manager to link the remaining tables.</b></p>
 <div align="left">

  <p><strong> <font color="#0000FF">2. (50)</font></strong> <font color="#0000FF"><b>What is the location of the table &quot;OrderInfo&quot;?</b></font><b><font color="#000000"><br />
    <input name="q2" type="text" id="q2" size="120" maxlength="120" value="table location:" />
    </font></b></p>
     <hr />
  <p><strong> <font color="#0000FF">3. (50)</font></strong> <font color="#0000FF"><b>What is one major difference in how Access stores a table versus how dBase stores a table?</b></font><b><font color="#000000"><br />
    <input name="q3" type="text" id="q3" size="120" maxlength="120" value="Access vs. dBase:" />
    </font></b></p>


     <hr />
     <p>4. (<strong>150</strong>) Once you have successfully relinked all tables, upload your Linked.accdb file to the D2L Assignment Folder &quot;Formative08_Link&quot;.</p>
  </div> 

    <table width="100%" border="1" cellpadding="1" cellspacing="1">
  <tr bgcolor="#FF0000">
    <td colspan="3"><div align="center">
      <input type="submit" name="Submit2" value="Submit" />
    </div></td>
  </tr>
</table>
<hr />
<hr />

</form>
<p>&nbsp;</p>
</body>
</html>
