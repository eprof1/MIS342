<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>MIS342Formative05

</title>
    <style type="text/css">
        .auto-style3 {
            color: #0000FF;
            font-weight: bold;
        }
        .auto-style5 {
            width: 100%;
        }
        .auto-style6 {
            width: 914px;
            height: 59px;
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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Formative05"/>
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
  Complete this problem by placing your answers in the spaces provided  below.<br />
Press the submit button.<br />
Save the confirmation page as an .htm file. </p>
 <h1 align="center">    <font color="#FF0000"><b>Advanced Access
        Queries </b></font></h1>
 <p>Answer the following questions using the Northwind.mdb database file located at:  <%Response.write("\\store\classes\" & YearTerm & courseid & "\ReadOnly\" ) %> <br />
    <hr /> <br />
 <p><font color="#0000FF">C<b>reate a &quot;List of Values&quot; query.</b></font></p>
    <p><span class="auto-style3">Use the Orders table </span></p>
    <p><span class="auto-style3">Use the 'IN' operator to locate all Customer Company Names that shipped an order to Austria or Sweden&nbsp; </span></p>
    <p><span class="auto-style3">Use the Totals row to &quot;Group By&quot; </span></p>
    <p><span class="auto-style3">1. (50) Paste the SQL statement below</span></p>
 </b> </font><b><font color="#000000">

<input id="q1" name="q1" type="text" maxlength="120" class="auto-style6" style="background-color: #00FFFF"  />

     <hr />

<br />
            </font></b>Parameter Query</p>
    <p>
        Locate all Customer Names that had a shipment between any two dates</p>
    <p>
        Sort the results ascending by date</p>
    <p>
        Run the query, using July 23, 1996 and September 5, 1996 as the start and end dates.</p>
    <p>
        (<strong>50</strong>) 2. Print the query results as a .pdf file named Parameter.pdf</p>
    <table class="auto-style5" border="2">
        <tr>
            <td>Parameter Query-Ex2</td>
        </tr>
    </table>
     <hr />

    <p>
        Find Duplicates Query</p>
    <p>
        Use the Order Details table</p>
    <p>
        Use the Find Duplicates Query Wizard to find duplicates of the Product/ProductID field</p>
    <p>
        Think about other ways to accomplish this task.</p>
    <p>
        (<strong>50</strong>) 3. Print the query results as a .pdf file named Duplicates.pdf</p>
    <table class="auto-style5" border="2">
        <tr>
            <td>Duplicates Query-Ex3</td>
        </tr>
    </table>
     <hr />

 <div align="left">
  <p> Find Unmatched Records Query </p>
     <p> Use the Order Details and Orders tables</p>
     <p> Output all Order Details records</p>
    <p class="auto-style3">4. (50) Briefly explain why no records are returned.&nbsp; Hint: no records is the correct answer.</p>
 </b> </font><b><font color="#000000">

     <input id="q4" name="q4" type="text" maxlength="120" class="auto-style6" style="background-color: #00FFFF"  />

            </font></b>
        <hr />

     <hr />

      <article class="h3">
        <div class="col-lg-12">
            <br />
            <a href="#Exercise05" class="btn btn-info" data-toggle="collapse">Combine Assignment and Screen Shots</a></div>
      </article>
      
    <div id="Exercise05" class="collapse"> 
      <article class="row">

        <div class="col-lg-12"> 
        <p>To proceed you must have Windows 10 which includes &#39;Print to PDF&#39; installed. You know it is installed when one of the print dialog box choices is &quot;Microsoft Print to PDF&quot;</p>
            <p>After publishing this file to your website, print a copy as a .pdf, save it as Formative05.pdf.<br />
            <img alt="print to pdf"  src="F00_PrintToPDF.JPG" />
            
            <br><br />

            </p>
            <p>Refer to this url and follow these steps.  <a href="https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/" target="_blank">https://www.howtogeek.com/248462/how-to-combine-images-into-one-pdf-file-in-windows/</a>  <br />
            You will be creating one .pdf (portable document format) file from the .pdf files that you have created, and combine it with your completed assignment from your website.</p>
        
        <ol>
            
              
          <li> Open File Explorer, make sure the .pdf files are properly named( Ex2.. , Ex3.., Formative05 ).&nbsp; </li>
            <li>Remember to use PDFill to Merge the&nbsp; .pdf files, save it as <strong>Formative05_Combined.pdf</strong></li>
          </ol>  
        <div class="col-lg-12 w3-label w3-text-blue">


        <p>(<strong>100</strong>) 5. Upload your file 'Formative05_Combined.pdf' to the D2L Assignment Folder 'Formative05'</p>

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
