<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>MIS342Formative06</title>
    <style type="text/css">
        .auto-style1 {
            width: 1039px;
            height: 560px;
        }
        .auto-style2 {
            text-decoration: underline;
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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Formative06"/>
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


<p><font color="#FF0000"><b>250 points</b></font></p>
    <p>&nbsp;</p>
    <p>Note: 
  Save this  assignment to your website.<br />
    </p>
    <p>Complete your personal information in the above section.<br />
  Complete this problem by placing your answers in the spaces provided  below.<br />
Press the submit button.<br />
Save the confirmation page as an .htm file. </p>
 <h2 align="center">    <font color="#FF0000"><b><em>Access
        CrossTab-Tables/Queries/Crosstabs/Reports! </em></b></font></h2>
 <p>Answer the following questions using the GraphingProblem_Data.accdb database file located at:     <%Response.write("\\store\classes\" & YearTerm & courseid & "\ReadOnly\" ) %> <br />

   
    Refer to the PowerPoint CrosstabGraphR2.pptx which is also at:  <%Response.write("\\store\classes\" & YearTerm & courseid & "\ReadOnly\" ) %>   <br />
  </p>
 <p><b><font color="#0000FF">1. (50) Paste the Crosstab </font></b><font color="#0000FF"><b> SQL
            statement in this text box:<br />
 </b></font><b><font color="#000000">
    <input name="q1" type="text" id="q1" size="120" maxlength="120" value="sql:" />
<br />
            </font></b></p>
 <div align="left">
  <p><strong> <font color="#0000FF">2. (50)</font></strong> <font color="#0000FF"><b>What is the value for Rogedesild May sales? Type the value in this text box:</b></font><b><font color="#000000"><br />
    <input name="q2" type="text" id="q2" size="120" maxlength="120" value="May sales dollars:" />
    </font></b></p>
     <hr />

    <p>3.&nbsp; Print out <span class="auto-style2">all</span> pages (about 77) of your completed form in landscape orientation, saving it as <strong>CrossTabGraph.pdf</strong></p>
     <p>Make sure:</p>
     <ul>


         <li>Your name is in the title</li>
         <li>Y-axis has a minimum value of zero and a maximum value of $25,000</li>
         <li>each data point contains a label with the value</li>
         <li>the legend is on the bottom of the page, and lines up fairly well with the columns</li>


     </ul>
     <p>&nbsp;</p>
     <p>Your output will look something like this:<img alt="chart output" class="auto-style1" src="ChartOutputAsPdf.JPG" /></p>
     <p>&nbsp;</p>
     <p>(<strong>150</strong>) Upload your <strong>CrosstabGraph.pdf</strong> file&nbsp; to the D2L Assignment Folder &quot;Formative06_Crosstab&quot;.</p>
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
