<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<head>
<title>MIS342Summative04</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<style type="text/css">
<!--
.Question {
	color: #00F;
	font-weight: bold;
}
    .auto-style2 {
        width: 1343px;
        height: 35px;
    }
    .grd *,.grd *:before,.grd *:after{box-sizing:border-box}
    .auto-style3 {
        font-weight: 700;
        font-family: "Segoe UI", "Segoe UI Web", "Segoe UI Symbol", "Helvetica Neue", "BBAlpha Sans", "S60 Sans", Arial, sans-serif;
    }
    .auto-style4 {
        width: 620px;
        height: 513px;
    }
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW"> 
<h2 align="center"><i><font color="#cc00FF">MIS 342 
      <% response.write(semester) %> 
    Summative04-Form and Report Basics</font></i></h2>

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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Summative04"/>
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

 <p><br />
     <font color="#FF0000"><b>550 points</b></font></p>
 <h1 align="center">Microsoft Access <br />
     4-Forms and Reports</h1>
<p><br />
  <br />
  Data Files needed for the Tutorial Project:<br />
  &bull; <a href="LuckyTire.accdb">LuckyTire.accdb</a></p>
<p><br />
  Len Phan wants to work with and display data about Lucky Tire&rsquo;s customers and product orders. He asks<br />
  you to create forms and reports for him.<br />
  <br />
  1. Start Access and open the LuckyTire database. (If the Security Warning appears, click the close<br />
  message bar button to close the warning.)<br />
  <br />
  2. Use the Form Wizard to create a form based on the <strong>Customers</strong> table. Select <strong>all</strong> the fields for the form.<br />
  Select the <strong>Columnar</strong> layout, and choose any style. Specify the title Customer Form for the form.<br />
  <br />
  3. Use the Customer Form to update the Customers table as follows:<br />
  &bull; Use the Find Command to search for the record that contains &ldquo;Prius&rdquo;. Display the record, and<br />
  change the Address to 1872 Juniper Ave.<br />
  (<strong>25</strong>) Make a screen shot of the record once you have changed the address and paste it below:<br />
</p>
<table width="95%" border="1" align="center">
  <tr>
    <td>Ex3_1ScreenShot</td>
  </tr>
</table>
<p><br />
  <br />
  &bull; Add a new record with the following values:<br />
  CustID: 50723<br />
  First: Stewart<br />
  Last: Gordon<br />
  Address: 1655 W E St #227<br />
  City: Upland<br />
  State: CA<br />
  Zip: 91762<br />
  Phone: 909-555-1276<br />
  Make: Subaru<br />
  Model: WRX<br />
  Year: 2007<br />
  &bull; Use the Find Command to search for the record that contains &ldquo;Suburu&rdquo;. Display the record, and<br />
  change the Make to Subaru.<br />
  (<strong>25</strong>) Make a screen shot of the record once you have changed the Make field to Subaru.<br />
</p>
<table width="95%" border="1" align="center">
  <tr>
    <td><p>Ex3_2ScreenShot</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p></td>
  </tr>
</table>
<p><br />
  4. Change the theme or design of the form to something of your choosing. Change the font color of the form title to any shade of brown.<br />
  Save and close the form.<br />
  <br />
  5. Use the Form Wizard to create a new form containing a main form and a subform. Select all the fields<br />
  from the Customers table for the main form, and then select all the fields from the Sales table for the<br />
  subform. Select the Tabular layout. Specify the title Customer Orders for the main<br />
  form, and the title Sales Subform for the subform.<br />
  <br />
  (<strong>50</strong>) Make a screen shot of the form/subform in Form view and paste it below:</p>
<table width="95%" border="1" align="center">
  <tr>
    <td><p>Ex5ScreenShot</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p></td>
  </tr>
</table>
<p><br />
  6. Use the appropriate wildcard character to find all records for model Year 1999 or earlier.</p>
    <p>Hint:&nbsp; see table below</p>
    <p>
        <table id="tblID0EBBDAAA" border="2" cellpadding="2" cellspacing="2" class="banded flipColors">
            <thead>
                <tr>
                    <td>
                        <p>
                            <b class="auto-style3">Character</b>
                        </p>
                    </td>
                    <td>
                        <p>
                            <b class="auto-style3">Description</b>
                        </p>
                    </td>
                    <td>
                        <p>
                            <b class="auto-style3">Example</b>
                        </p>
                    </td>
                </tr>
            </thead>
            <tr>
                <td>
                    <p>
                        *</p>
                </td>
                <td>
                    <p>
                        Matches any number of characters. You can use the asterisk (<b class="auto-style3">*</b>) anywhere in a character string.</p>
                </td>
                <td>
                    <p>
                        <b class="auto-style3">wh*</b> finds what, white, and why, but not awhile or watch.
                    </p>
                </td>
            </tr>
            <tr>
                <td>
                    <p>
                        ?</p>
                </td>
                <td>
                    <p>
                        Matches any single alphabetic character.</p>
                </td>
                <td>
                    <p>
                        <b class="auto-style3">B?ll</b> finds ball, bell, and bill.</p>
                </td>
            </tr>
            <tr>
                <td>
                    <p>
                        [ ]</p>
                </td>
                <td>
                    <p>
                        Matches any single character within the brackets.</p>
                </td>
                <td>
                    <p>
                        <b class="auto-style3">B[ae]ll</b> finds ball and bell, but not bill.</p>
                </td>
            </tr>
            <tr>
                <td>
                    <p>
                        !</p>
                </td>
                <td>
                    <p>
                        Matches any character not in the brackets.</p>
                </td>
                <td>
                    <p>
                        <b class="auto-style3">b[!ae]ll</b> finds bill and bull, but not ball or bell.</p>
                </td>
            </tr>
            <tr>
                <td>
                    <p>
                        -</p>
                </td>
                <td>
                    <p>
                        Matches any one of a range of characters.
                    </p>
                    <p>
                        You must specify the range in ascending order (A to Z, not Z to A).</p>
                </td>
                <td>
                    <p>
                        <b class="auto-style3">b[a-c]d</b> finds bad, bbd, and bcd.</p>
                </td>
            </tr>
            <tr>
                <td>
                    <p>
                        #</p>
                </td>
                <td>
                    <p>
                        Matches any single numeric character.</p>
                </td>
                <td>
                    <p>
                        <b class="auto-style3">1#3</b> finds 103, 113, and 123.</p>
                </td>
            </tr>
        </table>
    </p>
<p>  <strong>(<span class="Question">50) 6. How many records are returned?</span></strong><br />
  <label>
    <input name="q6" type="text" id="q6" class="auto-style2" maxlength="120"  size="120" style="background-color: #00FFFF" />
  </label>
  <br />
  Change the  Year field in record containing the Honda Civic to 1997. Save and close the form.<br />
  <br />
  7. Use the Report Wizard to create a report based on the primary Tires table and the related Sales table.<br />
  Select PartID, Model Name, and Manufacturer from the Tires table, and select Invoice, InvoiceDate,<br />
  and Quantity from the Sales table. <br />
  Select Manufacturer as an additional grouping level. <br />
  Sort the  records in ascending level by InvoiceDate. Select Block layout, Portrait orientation. <br />
  Specify the name Tire Orders for the report. If necessary, resize columns using Layout view.</p>
<p>(<strong>50</strong>) Make a screen shot of the report in Report View and paste below:<br />
</p>
<table width="95%" border="1">
  <tr>
    <td><p>&nbsp;</p>
      <p>Ex7ScreenShot</p>
      <p>&nbsp;</p></td>
  </tr>
</table>
<p><br />
  8. EXPLORE: Refer to Access Help, and search for &quot;Create a grouped or summary report&quot;&nbsp; then scroll down and read &quot;<strong>Add or modify grouping and sorting in an existing report&quot;</strong>&nbsp; <br />
    <div>

        <img alt="help screen for report sums" border="2"  src="NewHelp_ReportSums.PNG" /></div>
    <br />
  <br />
  &nbsp;Add a sum to the Quantity field.</p>
    <p>Hint: In Layout View, click on &quot;Group &amp; Sort&quot; in the &quot;
  Grouping &amp; Totals&quot; Group on the Report Design Tools &quot;Design&quot; tab to open the &quot;Group, Sort and Total&quot; footer. Then click &quot;More&quot; under &quot;Group on Manufacturer&quot; to find the appropriate total settings.
</p>
    <p>Close the Access Help window.<br />
  <br />
  9. Apply conditional formatting to the PartID totals so that any Quantity Total greater than 3 are<br />
  formatted using bold green font. (Hint: Select the first Quantity Total field.)<br />
  <br />
  10. Preview the report to verify that it is formatted correctly. If necessary, return to Layout View and make<br />
  any changes so the report prints within the margins of page. Save the report.<br />
  (<strong>100</strong>) Make a screen shot of the first page of  the report and paste it below. Then  close the report.<br />
</p>
<table width="95%" border="1">
  <tr>
    <td>
      <p>Ex10ScreenShot</p>
      <p>&nbsp;</p></td>
  </tr>
</table>
<p>  <br />
  <br />
  11. Change the Navigation Pane so that it displays all objects grouped by object type.<br />
  (<strong>25</strong>)  Make a screen shot of the Navigation Pane and paste it below.<br />
</p>
<table width="95%" border="1">
  <tr>
    <td>
      <p>Ex11ScreenShot</p>
      </td>
  </tr>
</table>
<p>  <br />
  <br />
  12. Close any open tables or queries. Compact and repair the LuckyTire database. Close it, and then close<br />
  Access.</p>
<p> <span class="Question">(25) What is the purpose of compacting and repairing a database?</span><br />
  <label>

       <input id="q12" type="text" class="auto-style2" maxlength="120" name="q12" size="120" 

           value="Value of compact and repair:  
           
           " align="top" style="background-color: #00FFFF" />

  </label>
</p>
<hr />
<div>
    <article class="h3">
        <div class="col-lg-12">&nbsp;<a href="#Exercise04" class="btn btn-info" data-toggle="collapse">Assignment as pdf</a></div>
      </article>
      
    <div id="Exercise04" class="collapse">
      <article class="row">
        <div class="col-lg-12"><p>For this Exercise, please read all of these instructions.<br>
        Make sure that you have completed all the previous exercises by 
filling in your answers, and publishing them to your website.<br>
        Then browse to your http://classes.winona.edu webiste, ensure all 
the Exercises are complete, then save your assignment as a .pdf file in the same
 location that you saved all of your screen shots for this assignment.<br>
        </p>
        </div>
        <div>          
            <div class="col-lg-12 w3-label w3-text-blue">
            (<strong>100</strong>) 13. Assignment .pdf file creation.<br />
                <br>
            </div>  
        </div>     
      </article>
     </div>
     
      
      <article class="h3">
        <div class="col-lg-12"><a href="#Exercise05" class="btn btn-info" data-toggle="collapse">Combine Assignment and Screen Shots</a></div>
      </article>
      
    <div id="Exercise05" class="collapse"> 
      <article class="row">

        <div class="col-lg-12"> 
        <p>To proceed you must have Windows 10 which includes &#39;Print to PDF&#39; installed. You know it is installed when one of the print dialog box choices is &quot;Microsoft Print to PDF&quot;<br />

            <img alt="print to pdf"  src="F00_PrintToPDF.JPG" />

            
            <br><br />

            Refer to this url and follow these steps.  <a href="https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/" target="_blank">https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/</a>  <br />
            You will be creating one .pdf (portable document format) file from the screen shots that you have taken, and combine it with your completed assignment from your website.</p>
        
        <ol>
            
              
          <li> Open File Explorer, make sure the screen shot files are properly named(Ex1.. , Ex2.. , Ex3.. ).&nbsp; </li>
            <li> Select the screen shots to be combined. </li>         
          <li>Right Click and select &#39;Print&#39; from the pop-up menu.      </li>    
          <li>See the howtogeek article for more options.</li>
            <li>Click on &#39;Print&#39; and save the file with a name such as &quot;xxx...ScreenShots.pdf&quot; in an appropriate location where you can easily find it again, such as OneDrive-MNSCU/...            </li>
            <li>Remember to use PDFill to Merge the screen shot .pdf file with your completed assignment .pdf file, save it as Summative.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 14. Upload your file 'Summative04.pdf' to the D2L Assignment Folder 'Summative04&#39;</p>

        </div>
        </div> 
      </article>
    </div>
       
<div>When all the Exercises are completed, while viewing this assignment
 on your https://classes.winona.edu/... website, press the Submit button
 below. </div>

     </div>
    <hr />

<p>.</p>
<h1 align="center">Concepts</h1>
<p>Sample focus question: How do you best summarize and display data?</p>
<p> Sample concept list (not exclusive-refer to the chapter introduction for additional concepts):</p>
<ul>
  <li>Wizards</li>
  <li>Form Design</li>
  <li>Themes</li>
  <li>Navigation</li>
  <li>Finding data</li>
  <li>Updating data</li>
  <li>Printing data</li>
  <li>Subforms</li>
  <li>Conditional formatting</li>
</ul>
<hr />


 <hr />
<table width="100%" border="1" cellspacing="1" cellpadding="1"> 
	<tr bgcolor="#FF0000"><td><div align="center">
		<input type="submit" name="Submit" value="Submit" />
</div></td></tr> 
</table>
    <br />
	
</form>
</body>
</html>
