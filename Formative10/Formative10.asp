<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>MIS342Formative10</title>
</head>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>MIS342ICE10</title>
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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Formative10"/>
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


<p><font color="#FF0000"><b>150 points</b></font></p>
    <p>Complete your personal information in the above section.<br />
  Complete this problem by placing your answer in the space provided  below.<br />
Press the submit button.<br />
Save the confirmation page as an .htm file. </p>
<h2 align="center">    <font color="#FF0000"><b><em>VBA Coding</em></b></font></h2>
 <p>Answer the following questions using the Northwind.mdb database file located
    at: </p>

    <%Response.write("\\store\classes\" & YearTerm & courseid & "\ReadOnly\Northwind.mdb" ) %>


     <br />

    <p>In Microsoft Access remember to press &#39;Alt + F11&#39; in order to bring up the VBA IDE.<br />
  </p>
 <p><b><font color="#0000FF">You will use  the Orders table as a data
      source.<br />
  Create a function called 'GermanFreight' that will return the  average value
        from the 'Freight' field of the Orders table, where the value of the
        'Ship
        Country'
        field
  is 'Germany'.</font></b></p>
 <p><b><font color="#0000FF">Hint: use Davg( )</font></b> </p>
 <p><font color="#0000FF"><b>Find the value of the 'GermanFreight' function by
<br />
        Creating a form that contains a text box.<br />
        Set the text box data source to the 
      function 'GermanFreight'<br />
    View the form in datasheet mode.</b></font></p>
 <p><font color="#0000FF"><b>OR by using the Immediate Window.</b></font><b><font color="#000000"><br />
      <br />
      <font color="#0000FF">1. (50) Paste the code for the 'GermanFreight' function in this text box:
    </font><br />
    <input name="q1" type="text" id="q1" size="120" maxlength="120" value="code:" />
  </font></b></p>
 <div align="left">
    <p><strong> <font color="#0000FF">2.&nbsp; (50)</font></strong> <font color="#0000FF"><b>Paste or type the value
          of the 'GermanFreight' function in this text box. <br />
        Note:  You will get no credit for this part if you have not answered
        the first part.</b></font><b><font color="#000000"><br />
    <input name="q2" type="text" id="q2" size="120" maxlength="120" value="value:" />
  </font></b></p>
     <hr />
    <p><b><font color="#0000FF">3. (50) Evaluate the following expression using
        Access:</font></b></p>
    <p><font color="#0000FF"><strong>log(2.7182889) + (17 * 7^4.8) \ 4 + log(34) - 15 mod 6</strong></font></p>
    <p><font color="#0000FF">Paste your answer below.<br />
    Make sure your answer is to 4 decimal places. <br />
    Remember the Immediate Window
    .</font></p>
    <p><img src="ImmediateWindow.png" width="363" height="200" alt="Immediate Window" /></p>
    <p>&nbsp;</p>
    <p><b><font color="#000000">
    <input name="q3" type="text" id="q3" size="120" maxlength="120" value="result:" />
        <hr />
    </font></b></p>
  </div> 

        <table width="100%" border="1" cellpadding="1" cellspacing="1">
  <tr bgcolor="#FF0000">
    <td colspan="3"><div align="center">
      <input type="submit" name="Submit2" value="Submit" />
    </div></td>
  </tr>
</table>

</form>

</body>
</html>
