<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<head>
<title>MIS342Summative07</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<style type="text/css">
<!--
.HWOutput {
	color: #00F;
}
hwquestion {
	color: #00F;
}
hwquestion {
	color: #00F;
}
#HWquestion {
	color: #00F;
}
.Question {color: #00F;
	font-weight: bold;
}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW"> 
<h2 align="center"><i><font color="#cc00FF">MIS 342 
      <% response.write(semester) %> 
    Summative07-Custom Reports</font></i></h2>

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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Summative07"/>
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
     <font color="#FF0000"><b>450 points</b></font></p>
 <h1 align="center">Microsoft Access</h1>
 <p><em>You  will need  to download the <strong>LuckyTire</strong> database that was  created in Tutorial 6.<br />
It is availabe at:<br />
\\store\classes\20185001248\ReadOnl<strong>y\LuckyTire06.accdb</strong></em></p>
 <p>Len Phan is  the owner of Lucky Tire, a chain of tire and auto service centers. Len would  like to create a of couple of reports. You&rsquo;ll do so by completing the following:<br />
  &nbsp; </p>
 <ol start="1" type="1">
   <li>Open the <strong>LuckyTire06 </strong>database.</li>
 </ol>
 <ol start="2" type="1">
   <li>Use the figure below and the       following steps to modify the <strong>rptTiresReport </strong>report:</li>
   <ol start="1" type="a">
     <li>Resize the columns as        necessary.</li>
     <li>Insert        a space before &ldquo;ID&rdquo; in the <strong>PartID </strong>column        heading. </li>
     <li>Apply        the alternate background color setting, using the Alternate Row color        (row 1, column 8) in the Access Theme Colors palette. </li>
     <li>Group        the report by Manufacturer, and then display the average unit price for        each manufacturer. </li>
     <li>Add a conditional formatting        rule for the UnitPrice field to display the price in green, bold, italic        font when the unit price is greater than $60.</li>
   </ol>
 </ol>
 <p>&nbsp;</p>
 <p><img src="Homework07_clip_image002.jpg" alt="screen shot 2" width="576" height="165" /><br />
   <br />
   (<strong>50</strong>) Print and save the report as a .pdf file, Ex2rptTiresReport.pdf.<br />
 </p>
 <table  width="100%" border="2">
   <tr>
     <td align="left" >rptTiresReport-Ex2rptTiresReport.pdf
       </td>
   </tr>
 </table>
 <p>&nbsp; </p>
 <ol start="3" type="1">
   <li>Create a query that displays the       CustID, First and Last name fields from the tblCustomer table, <br />
     the PartID       and Quantity fields from the tblSales table, <br />
     and the ModelName,       Manufacturer, TireType, and UnitPrice fields from the tblTires table. <br />
     Sort       in ascending order by the TireType, Manufacturer, and Last name, and then       save the query as <strong>qryTireType</strong>.       <br />
     Save and close the query.<br />
     <br />
     <span class="HWOutput"><span id="HWquestion"><strong>(50) Provide the sql for the query in the text box</strong>:<br />
     </span></span>
     <label>
    <input name="q3" type="text" id="q3" size="120" maxlength="120" value="sql:" />
     </label>
   </li>
 </ol>
 <p>&nbsp;</p>
 <ol start="4" type="1">
   <li>Create a custom report based on       the qryTireType query. The following figure shows a sample of the last       page of the completed report. <br />
    Refer to the figure as you create the       report. </li>
   <ol start="1" type="a">
     <li>Save the report as <strong>rptTireType</strong>.</li>
     <li>Add the fields in the following        order: TireType, Manufacturer, ModelName, UnitPrice, Last, First, and        Quantity.</li>
     <li>TireType and Manufacturer        fields are grouping fields.</li>
     <li>UnitPrice and Last fields are sort        fields.</li>
     <li>Show the total quantity in the        group footer for both the TireType and Manufacturer groups.</li>
     <li>Add a horizontal line in the        Page Header, below the column labels.</li>
     <li>Add the title <strong>Quantity of Tires by Type</strong> and        insert the date in the upper right.</li>
     <li>Add the page number, using the        page N of M format, to the Page Footer, and right-align the number.</li>
   </ol>
 </ol>
 <p><img src="rptTireType.png" alt="tire type" width="748" height="438" border="2" /><br />
   <br />
   (<strong>50</strong>) Print and save the report as a .pdf file, Ex4rptTireType.pdf.<br />
 </p>
 <table width="100%" border="2">
   <tr>
     <td align="left" > rptTiretype-Ex4rptTireType.pdf
       </td>
   </tr>
 </table>
<p>&nbsp; </p>
 <ol start="5" type="1">
   <li>Create a mailing label report       according to the following instructions:</li>
   <ol start="1" type="a">
     <li>Use the tblCustomer table as        the record source.</li>
     <li>Use Avery C2160 labels, use a        larger font size and a heavier font weight, and use the other default        font and color options.</li>
     <li>For the prototype label, place        CustID on the first line, First, a space, and Last on the second line;        Address on the third line; and City, a comma, State, a space, and Zip on        the fourth line.</li>
     <li>Sort by Zip, and then by Last,        and then enter the report name <strong>rptCustomerMailingLabels</strong>.</li>
     <li>Make sure the mailing label        layout displays across and then down.<br />
       <br />
       <p>(<strong>50</strong>) Print and save the report as a .pdf file, Ex5rptCustomerMailingLabels.pdf.<br />
       </p>
       <table width="100%"  border="2">
         <tr>
           <td align="left" valign="top">rptCustomerMailingLabels-Ex5rptCustomerMailingLabels.pdf </td>
         </tr>
       </table>
</li>
   </ol>
 </ol>
 <p>&nbsp;</p>
 <ol start="6" type="1">
   <li>Compact and repair the LuckyTire       database, close the database, and then exit Access.</li>
   <li><span class="Question">(50) Briefly explain why the ability to create database reports (not just in  Access) is a valuable skill to learn in advancing your career.</span><br />
     <label>
    <input name="q7" type="text" id="q7" size="120" maxlength="120" value="report skill valuable:" />
     </label>
   </li>
  </ol>
 <p>&nbsp;</p>
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
            (<strong>100</strong>) 8. Assignment .pdf file creation.<br />
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
            <img alt="print to pdf" class="auto-style1" src="F00_PrintToPDF.JPG" /><br><br />

            Refer to this url and follow these steps.  <a href="https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/" target="_blank">https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/</a>  <br />
            You will be creating one .pdf (portable document format) file from the screen shots that you have taken, and combine it with your completed assignment from your website.</p>
        
        <ol>
            
              
          <li> Open File Explorer, make sure the screen shot files are properly named(Ex1.. , Ex2.. , Ex3.. ) and sorted in Ascending order.&nbsp; Note: there are no screen shots in this assignment, only saved pdf files</li>
            <li> Select the screen shots to be combined. </li>         
          <li>Right Click and select &#39;Print&#39; from the pop-up menu.      </li>    
          <li>See the howtogeek article for more options.</li>
            <li>Click on &#39;Print&#39; and save the file with a name such as &quot;xxx...ScreenShots.pdf&quot; in an appropriate location where you can easily find it again, such as OneDrive-MNSCU/...            </li>
            <li>Remember to use PDFill to Merge the saved .pdf files with your completed assignment .pdf file, save it as Summative07.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 9. Upload your file 'Summative07.pdf' to the D2L Assignment Folder 'Summative07'</p>

        </div>
        </div> 
      </article>
    </div>
       
<div>When all the Exercises are completed, while viewing this assignment
 on your https://classes.winona.edu/... website, press the Submit button
 below. </div>

     </div><hr />

 <h1 align="center">Concepts</h1>
 <p> Sample focus question: What are the benefits of creating custom reports?</p>
 <p> Sample concept list (not exclusive-refer to the chapter introduction for additional concepts):</p>
 <ul>
   <li>Report design</li>
   <li>Data source-query</li>
   <li>Data source-sql statement</li>
   <li>Sorting data</li>
   <li>Grouping data</li>
   <li>Formatting</li>
   <li>Duplicate values</li>
   <li>Dates</li>
   <li>Page numbers</li>
   <li>Titles</li>
   <li>Mailing labels</li>
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
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW0" id="frmHW0"> 
	<table border="1" cellpadding="1" cellspacing="1" width="98%">
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>

</form>
</body>
</html>
