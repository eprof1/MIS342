<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>MIS342Formative01</title>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="frmHW" method="post" action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp"> 
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
      <td width="36%">Semester: <input type="text" name="Semester" value=<% response.write(semester)%> />  </td>
      <td>Class:  <input type="text" name="Class" value=<% response.write(xClass)%> />    </td>
      <td><input name="InstID" type="hidden" id="InstID" value="00617282" />       StarID:
        
        <!-- Enter your StarID in the VALUE field below   GROK --> 
      <input type="text" name="PIN" value="EnterYourStarID" /></td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td width="36%">Section:  <input type="text" name="Section" value="01"/>
      </td>
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Formative01"/>
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


<p>(<strong>150 points</strong>)</p>

<p>Complete your personal information in the above section.<br>
  Complete this problem by placing your answer in the space provided  below.<br>
  Press the submit button.<br>
  Save the confirmation page as an .htm file. </p>
 <h2 align="center">    <FONT color="#FF0000"><B><em>Formative01-FlatFiles, or Why Not Excel?</em></B></FONT></h2>
 <p><br>
  </p>
 <p><b><font color="#000000"> Make a copy of FlatFile.xls, which is at:  <%Response.write("\\store\classes\" & YearTerm & courseid & "\ReadOnly\" ) %>  <br>
    and save it to: ../Documents/MIS342/Formative/Formative01<br>
    Examine the file.
  </font></b></p>
 <p><b><font color="#000000">Answer the following questions: </font></b></p>
 <p><b><font color="#000000"><br>
    Examine the entries that have been made into the 'CompanyName' field (column). <br>
      <br>
      <font color="#0000FF">(25) 1. </font></font></b><font color="#0000FF"><b> How many records can you find for the Bottom-Dollar Markets? Be careful and watch out for the spelling. </b></font><b><font color="#000000"><br>
    <input name="q1" type="text" id="q1" size="120" maxlength="120" value="number of records:" />
  </font></b></p>
 <div align="left">
    <p><strong> <font color="#0000FF">(25)</font></strong> <font color="#0000FF"><b>2. How many different spellings are there for the Bottqlm-Dollar Markets? </b></font><b><font color="#000000"><br>
    <input name="q2" type="text" id="q2" size="120" maxlength="120" value="number of spelling variations:" />
    </font></b></p>
    <p><strong>    <font color="#0000FF">(25) 3. Which spelling of the Bottolm-Dollar Market is correct? Justify your answer. </font></strong><br>
    <input name="q3" type="text" id="q3" size="120" maxlength="120" value="correct spelling:" />
      <br>
      <br>
      <strong><font color="#0000FF">(75) 4. What problems do you think a business using this Excel spreadsheet to store order data would encounter? </font></strong><br>
    <input name="q4" type="text" id="q4" size="120" maxlength="120" value="problems:" />
  </p>
    <p>&nbsp;  </p>
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
</body>
</html>
