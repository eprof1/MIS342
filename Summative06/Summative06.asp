<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<head>
<title>MIS342Assignment06</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<style type="text/css">
<!--
.HWOutput {
	color: #F00;
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
    Summative06-Advanced Forms</font></i></h2>

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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Summative06"/>
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
     <font color="#FF0000"><b>500 points</b></font></p>
 <h1 align="center">Microsoft Access<br />
     </h1>
 <p><a name="_Toc53964503" id="_Toc53964503">Tutorial  Project</a> <br />
  <em>You  will need  to download the <strong>LuckyTire</strong> database that was  created in Tutorial 5.<br />
  It is availabe at:<br />
  \\store\classes\20185001248\ReadOnl<strong>y\LuckyTire05.accdb</strong></em></p>
 <p>Len Phan is  the owner of Lucky Tire, a chain of tire and auto service centers. Len would  like to create a couple forms. You&rsquo;ll do so by completing the following:</p>
 <ol start="1" type="1">
   <li>Open the <strong>LuckyTire </strong>database from Tutorial 5. If the Security       Warning appears, enable the content. </li>
   
   <li>Remove the lookup feature from       the PartID field in the <strong>tblSales </strong>table,       and then resize the Part ID column to its best fit.</li>
   
   <li>Define a one-to-many       relationship between the primary tblTires table and the related tblSales       table. Select the referential integrity option and the cascade updates       option for this relationship.</li>
   
   <li>(<strong>50</strong>) Use the Documenter to document       the tblSales table. Select all table options; use the Names, Data Types,       and Sizes option for fields; and use the Names and Fields option for       indexes.<br />
       Export the report produced by the Documenter as an HTML file. Place it in the same folder on your website as this assignment. <br />
       The default name will be doc_rptObjects.html, which is acceptable. <br />
     Do not select an HTML template, and use default encoding.<br />
     <img src="DocumenterHTMLReport.png" width="384" height="194" alt="output options" /><br />
     <strong>Provide a link to 
     the HTML report below</strong><br />
     Note that there will be several pages. <br />
       <br />
             
     <table width="95%"  border="1">
       <tr>
         <td align="left" valign="top"><p>Hyperlink to Documenter-Ex4 Documenter HTML file, page 1</p>
          <p>&nbsp;</p></td>
        </tr>
      </table>


       <br />
    <br />
   </li>
   
   <li>Use the Multiple Items tool to       create a form based on the qryCustomersWithoutSales query, change the       title to <strong>Customers Who Have Not       Bought Tires</strong>, and then save the form as <strong>frmCustomersWithoutSalesMultipleItems</strong>.<br />
     (<strong>25</strong>) Provide a screen shot of the form in Form View below.<br />
     <table width="95%" height="73" border="1">
       <tr>
         <td align="left" valign="top"><p>frmCustomersWithoutSalesMultipleItems-Ex5ScreenShot</p>
          <p>&nbsp;</p></td>
        </tr>
      </table>
     <br />
     <br />
   </li>
 </ol>
 <ol start="6" type="1">
   <li>Use the <strong>Split Form tool</strong> to       create a split form based on the tblTires table, and then make the       following changes to the form in Layout view.</li>
   <ol start="1" type="a">
     <li>Reduce the widths of all five        text boxes to a reasonable size. (Hint: Use the longest Model Name as a        guide.) </li>
     <li>Remove the ModelName and        TireType controls and their labels from the stacked layout, move these        two controls to the right and then to the top of the form. Add these two        controls to a new stacked layout and then anchor them to the top right.</li>
     <li>Remove the UnitPrice control        from the stacked layout, and then anchor the control to the bottom left.</li>
     <li>Make sure the tab order is        left-to-right, top-to-bottom for the main form text boxes.</li>
     <li>Change the title to <strong>Tire Data</strong>, and then save the        modified form as <strong>frmTiresSplitForm</strong>.<br />
       <br />
       (<strong>50</strong>) Provide a screen shot of the form in Form View below.<br />
       <table width="95%" height="73" border="1">
         <tr>
           <td align="left" valign="top"><p>frmTiresSplitForm-Ex6ScreenShot&nbsp;</p>
            <p>&nbsp;</p></td>
         </tr>
       </table>
       <br />
       <br />
     </li>
    </ol>
   
   <li>Use the figure below and the       following steps to create a custom form named <strong>frmTireOrders </strong>based on the tblSales table.</li>
 </ol>
<p><img src="Homework06_clip_image002.jpg" alt="" width="576" height="330" /></p>
 <ol start="7" type="1">
   <ol start="1" type="a">
     <li>Place the following fields from        the tblCustomer table at the top of the Detail section: First, Last,        Phone, Make, Model, and Year.</li>
     <li>Move the fields into two        columns in the Detail section, and apply the Stacked layout to both        columns. Resize and align&nbsp; controls        as necessary.</li>
     <li>Add the title <strong>Tire Order</strong> in the Form Header        section, and bold the title.</li>
     <li>Make sure the form&rsquo;s Record        Source property is set to tblCustomer, and then add a combo box to the        Form Header section to find the field values for First, Last, Phone,        Make, Model, and Year. In the wizard steps, select the CustID, Last, and        First fields, and do not hide the key column. Label the control <strong>Select Customer</strong>. Resize and        position the control. </li>
     <li>Create a simple query using all        the fields from the tblSales table and the UnitPrice fields from the        tblTires table. Name this query <strong>qryInvoiceTotal</strong>.        Use the Builder to add a new field that displays the Quantity field        multiplied by the UnitPrice fields and name this field <strong>InvAmt</strong>. </li>
     <li>Add a subform based on the <strong>qryInvoiceTotal</strong> query, name the        subform <strong>frmInvoiceTotalSubform</strong>.        Change the label of InvoicePaid to <strong>Paid?</strong> And the label of InvAmt to <strong>Total</strong>.        Delete the subform label, resize the columns in the subform for their        best fit, and resize and position the subform.</li>
     <li>Add a calculated control that        displays the total of the InvAmt field in the subform. Set the label&rsquo;s        Caption Property to <strong>Total of        Invoices:</strong>. Set the calculated control&rsquo;s Name to <strong>txtInvAmtSum</strong>, the Tab Stop property to No, the ControlTip        Text property to <strong>Calculated total        of invoices</strong>. Format the calculated field as currency.</li>
     <li>Add a rectangle around the        Total of Invoices control in the Detail section and set the line thickness        to 2 pts.</li>
     <li>Use the Label tool to add your        name below the subform at the bottom of the Detail section.</li>
     <li>For the labels in the Details        section, except for label displaying your name, use the Dark Label Text        font color (row 1, column 3 in the Access Theme Colors palette) or equivalent.</li>
     <li>For the title, Select Customer        label, and the label displaying your name, use the Maroon font color (row        1, column 6 in the Standard Colors palette) or equivalent, and bold the font.</li>
     <li>For the background color of the        Form Header section, use the Light Header Background color (row 1, column        6 in the Access Theme Colors palette) or equivalent.</li>
     <li>Make sure the tab order is        top-to-bottom, left-to-right for the main form text boxes.<br />
       <br />
       <br />
        Provide a screen shot of the form in Form View below.<br />
       (<strong>25</strong>) Select Customer: combo box <br />
       (<strong>50</strong>) Total of Invoices: correct calculation <br />
       <table width="95%" height="73" border="1">
         <tr>
           <td align="left" valign="top"><p>frmTireOrders-Ex7ScreenShot</p>
             <p>&nbsp;</p>
             <p>&nbsp;</p></td>
         </tr>
       </table>
     </li>
   </ol>
 </ol>
 <p>&nbsp;</p>
 <ol start="8" type="1">
   <li>Close the LuckyTire database       without exiting Access, open the <strong>LuckyTire </strong>database, compact and repair the database, close the database, and       then exit Access.</li>
   <li><span class="Question">(75) Briefly explain why forms are not the best Access database object to use when you want to create printed output.</span><br />
     <label>
    <input name="q9" type="text" id="q9" size="120" maxlength="120" value="forms not for printing:" />
     </label>
     <br />
     <br />
   </li>
   <li><span class="Question">(25) For form/subform created in step 7, describe what type of relationship must exist between the tblCustomer and tblSales. <br />
     If your answer is more than about 6 words it is too long.
   </span><br />
     <label>
    <input name="q10" type="text" id="q10" size="120" maxlength="120" value="relationship:" />
     </label>
   </li>
  </ol>
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
            (<strong>100</strong>) 11. Assignment .pdf file creation.<br />
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
            
              
          <li> Open File Explorer, make sure the screen shot files are properly named(Ex1.. , Ex2.. , Ex3.. ) and sorted in Ascending order.&nbsp; </li>
            <li> Select the screen shots to be combined. </li>         
          <li>Right Click and select &#39;Print&#39; from the pop-up menu.      </li>    
          <li>See the howtogeek article for more options.</li>
            <li>Click on &#39;Print&#39; and save the file with a name such as &quot;xxx...ScreenShots.pdf&quot; in an appropriate location where you can easily find it again, such as OneDrive-MNSCU/...            </li>
            <li>Remember to use PDFill to Merge the screen shot .pdf file with your completed assignment .pdf file, save it as Summative06.pdf</li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 12. Upload your file 'Summative06.pdf' to the D2L Assignment Folder 'Summative06'</p>

        </div>
        </div> 
      </article>
    </div>
       
<div>When all the Exercises are completed, while viewing this assignment
 on your https://classes.winona.edu/... website, press the Submit button
 below. </div>

     </div>    <hr />

 <h1 align="center">Concepts</h1>
 <p> Sample focus question: In what ways can a database system be enhanced using advanced form functionality?</p>
 <p> Sample concept list (not exclusive-refer to the chapter introduction for additional concepts):</p>
 <ul>
   <li>Form design</li>
   <li>Lookup fields</li>
   <li>Database documenter</li>
   <li>Form types</li>
   <li>Form tools</li>
   <li>Multiple items </li>
   <li>Split form</li>
   <li>Field placement</li>
   <li>Form controls</li>
   <li>Headers and footers</li>
   <li>Subforms</li>
   <li>Calculated controls</li>
   <li>Tab order</li>
 </ul>
<hr />

 <hr />
<table width="100%" border="1" cellspacing="1" cellpadding="1"> 
	<tr bgcolor="#FF0000">
      <td>
	  <div align="center">
		<input type="submit" name="Submit" value="Submit" />
      </div>        
      </td>
	</tr>
    </table>

</form>
</body>
</html>
