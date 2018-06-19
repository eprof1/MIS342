<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

<title>MIS342Formative07 </title>

    <style type="text/css">
        .auto-style3 {
            color: #0000FF;
            font-weight: bold;
        }
        .auto-style6 {
            height: 22px;
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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Formative07"/>
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


<p><font color="#FF0000"><b>500 points</b></font></p>
    <p>&nbsp;</p>
    <p>Note: 
  Save this  assignment to your website.<br />
    </p>
    <p>Complete your personal information in the above section.<br />
  Complete this problem by placing your answers in the spaces provided  below.<br />
Press the submit button.<br />
        It is a good idea to save the confirmation page as an .htm or .pdf file. </p>
 <h1 align="center">    <font color="#FF0000"><b>Advanced Reports</b></font></h1>
 <p>Refer to the <a href="../../Presentations/Chapter07.ppt">Advanced Report Tips</a> presentation and create the following report using the <strong>Northwind.mdb</strong> database file located at: <br />
     <%Response.write("\\store\classes\" & YearTerm & courseid & "\ReadOnly\" ) %>
     <br />

     </p><hr />
    <p>Create a report that meets the following requirements and contains the following fields, criteria, sort order

    </p>
    <table style="width: 100%;" border="2">
        <tr>
            <th>field</th>
            <th>criteria or sort</th>
            <th>Notes</th>
        </tr>
        <tr>
            <td class="auto-style6">CustomerID</td>
            <td class="auto-style6"></td>
            <td class="auto-style6"></td>
        </tr>
        <tr>
            <td>OrderID</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
        </tr>
        <tr>
            <td>ProductID</td>
            <td>Ascending</td>
            <td>&nbsp;</td>
        </tr>
        <tr>
            <td>OrderDate</td>
            <td>between 1/1/1998 and 12/31/1998</td>
            <td>ordered in 1998, but not output to report</td>
        </tr>
        <tr>
            <td>UnitPrice</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
        </tr>
        <tr>
            <td>Quantity</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
        </tr>
        <tr>
            <td>ExtendedPrice</td>
            <td>UnitPrice * Quantity</td>
            <td>calculated field</td>
        </tr>
    </table>

    <p>
        Make use of the Report Wizard
        (note it will do most, but not all of the work for you)</p>
    <p>
        Group by CustomerID and OrderID</p>
    <p>
        &nbsp;Sort by ProductID
        </p>
    <p>
        Use the Summary Options to create ExtendedPrice sub-totals for each ProductID
    </p>

    <p><span class="auto-style3">1. (50) Explain why is it very good practice to make a SQL statement, instead of a query, the data source for any report.</span></p>
    <img src="ReportRecordSource.JPG" />
    <br />
    <input name="q1" type="text" id="q1" size="120" maxlength="120" value="sql vs. query:" />
     <hr />

    <strong>Potential Problem Areas</strong><br />
           
    <br />

    <hr />
    When the followng dialog box appears in the Report Wizard make sure to choose the followning settings-choose Landscape orientation, and uncheck "Adjust the field..." or suffer the consequences.<br />
    <img src="ReportWizardFormatSettings.JPG" width="400" />

    <hr />
    Keep in mind that the Report Wizard can save a lot of time, but rarely produces the final product.  After the Wizard is done you will need to make adjustments in the Report Design View.
    <br />
    Plan on your report output being on Legal Paper (8 1/2 x 14 inches) with .25 inch margins.&nbsp;
    <br />
    This will provide a 14 - 1/2 or 13 1/2 inches for workspace in the Report Design View.<br />
    You are advised to adjust the left and width properties of any object that extends beyond the 13.5 inch ruler mark.<br />
    You can traverse all the objects on the report in design view by selecting the object from the dropdown list at the top of the Property Sheet and arrowing down.
    <br />
    Each will be highlighted and violators will be easy to spot.<br />
    <img src="ReportDesignViewIssues.JPG" width="1000" />

    <hr />
    Beware of formatting errors such as <strong>text boxes too narrow</strong> and lead to truncated fields.<br />
    <img src="FieldWidthProblems.JPG" width="800" />
    <hr />
    <br />
    <br />
    <p> View of totals, and Grand Total at end of report-make them look nice by setting number format to Currency, and decimal places to 0 (Zero)</p>
    <img src="GrandTotal.JPG" width="1000" />
    <hr />
    <strong>Realize that it may take several attempts using the Report Wizard to get a satisfactory report.
    </strong>
    <br />

    <hr />

     <hr />

    <div id="Exercise05" class="collapse"> 
      <article class="row">

        <div class="col-lg-12"> 
        <p>To proceed you must have Windows 10 which includes &#39;Print to PDF&#39; installed. You know it is installed when one of the print dialog box choices is &quot;Microsoft Print to PDF&quot;<br />
            <img alt="print to pdf" src="../Formative05/F00_PrintToPDF.JPG" /><br><br />

            When done, print your Report as a .pdf file named &quot;<strong>Formative07.pdf</strong>&quot;</p>
            <p>Use Legal size paper.</p>
            <p>Make sure that the report is in Landscape mode.</p>
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>450</strong>) 2. Upload your file 'Formative07.pdf' to the D2L Assignment Folder 'Formative07'</p>

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
