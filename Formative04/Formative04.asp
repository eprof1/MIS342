<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>MIS342Formative04</title>

    <style type="text/css">
        .auto-style3 {
            color: #0000FF;
            font-weight: bold;
        }
        .auto-style5 {
            width: 100%;
        }
        .auto-style6 {
            width: 814px;
        }
        .auto-style7 {
            width: 601px;
            height: 778px;
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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Formative04"/>
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


<p><font color="#FF0000"><b>450 points</b></font></p>
    <p>&nbsp;</p>
    <p>Note: 
  Save this  assignment to your website.<br />
    </p>
    <p>Complete your personal information in the above section.<br />
  Complete this problem by placing your answers in the spaces provided  below, and providing requested screen shots and .pdf files.<br />
Press the submit button.<br />
        You are advised to save the confirmation page as a .pdf file. </p>
 <h1 align="center">    <font color="#FF0000"><b>Basic Access Forms and Reports </b></font></h1>
 <p>Answer the following questions using the Northwind.mdb database file located on the class network drive at:  <%Response.write("\\store\classes\" & YearTerm & courseid & "\ReadOnly\" ) %> <br />
    <hr /> <br />
    Forms-<br />
    <br />
    Use your knowledge of Access to create a form / subform based on the Order Details and Orders tables.<br />
    When done, save the form as frmOrders.<br />
    <br />
    <p><span class="auto-style3">1. (50) Which table is the data source for the subform?</span><font color="#0000FF"></b></font></p>
    <p><b><font color="#000000">
        <input id="q1" name="q1" type="text" maxlength="120" class="auto-style6" style="background-color: #00FFFF"  /><hr />

    <p class="auto-style3">2. (50) Which table is the data source for the form?</p>
        </font>
        <input id="q2" name="q2" type="text" maxlength="120" class="auto-style6" style="background-color: #00FFFF"  /><hr />
        
<br />
            <br />
    3. (100)&nbsp; When done, print the form and save as frmOrders.pdf<br />
           </b>
    <table class="auto-style5" border="2">
        <tr>
            <td>frmOrders.pdf</td>
        </tr>
    </table>
        <hr /> <hr />


     <p> Reports-</p>
    <p> Use the Northwind.mdb file and your knowledge of Access to create a report with the following features:</p>
    
   <ol>
       <li>Customers are located in Region OR or BC</li>
       <li>RequiredDate is between 1/1/97 and 12/31/97   </li>

   </ol>
<p>
    Your report must output only the following fields
</p>

        <ul>
            <li>Customer Region, sorted in Ascending order</li>
            <li>RequiredDate, sorted in Ascending order</li>
            <li>ProductID or ProductName from the Order Details</li>
        </ul>
        
        <p>Hint: first create a query, which will return 54 records using the above criteria.<br />
            Be very careful in applying the criteria, especially with respect to 'AND' and 'OR'
        </p>
    <p class="auto-style3">
        4. (50) What are the three tables needed for the query or solution?</p>
    <input id="q4" name="q4" class="auto-style6" maxlength="120"  type="text"  /><hr />
    <br />
        <hr />
    <p>
        Create a report based on the query to properly display the results.</p>
    <p>
        Hint: feel free to experiment and use the Report Wizard.&nbsp; One possible report looks like this:<br />&nbsp;
        <img alt="1997 orders report" class="auto-style7" src="rpt1997OrdersBCOR.JPG" /></p>
    <p>
        &nbsp;</p >
    <b>5. (100)&nbsp; When done, print the report and save as rpt1997OrdersBC_OR.pdf<br />
            </font></b>
    <table class="auto-style5" border="2">
        <tr>
            <td>rpt1997OrdersBC_OR.pdf</td>
        </tr>
    </table>
        <br />
        <hr />
     <hr />

      <article class="h3">
        <div class="col-lg-12"><a href="#Exercise05" class="btn btn-info" data-toggle="collapse">Combine Files</a></div>
      </article>
      
    <div id="Exercise05" class="collapse"> 
      <article class="row">

        <div class="col-lg-12"> 
        <p>To proceed you must have Windows 10 which includes &#39;Print to PDF&#39; installed. You know it is installed when one of the print dialog box choices is &quot;Microsoft Print to PDF&quot;<br />
            <br>Refer to this url and follow these steps.  <a href="https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/" target="_blank">https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/</a>  <br />
            Create one .pdf (portable document format) file from the form and report .pdf files that you have created for this assignment.</p>
        
            <p>
                Uuse PDFill to Merge the form and report .pdf files, save it as Formative04.pdf</p>
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 6. Upload your file 'Formative04.pdf' to the D2L Assignment Folder 'Formative04'</p>

        </div>
        </div> 
      </article>
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
