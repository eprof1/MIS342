<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/MIS342/ClassConstants.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>MIS342Formative03</title>

  
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
      <td colspan="2"> Assignment: <input  name="Assignment" type="text" value="Formative03"/>
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


<p><font color="#FF0000"><b>700 points</b></font></p>
    <p>&nbsp;</p>
    <p>Note: 
  Save this  assignment to your website.<br />
    </p>
    <p>Complete your personal information in the above section.<br />
  Complete this problem by placing your answers in the spaces provided  below, and providing the requested .pdf file, which is made by combining the results of the queries.<br />
Press the submit button.<br />
        You are advised to save the confirmation page as a .pdf file. </p>
 <h1 align="center">    <font color="#FF0000"><b>Queries! </b></font></h1>
 <p>Answer the following questions using a copy of the <strong>Northwind.mdb</strong> database file located on the class network drive at:  <%Response.write("\\store\classes\" & YearTerm & courseid & "\ReadOnly\" ) %> <br />
     You are advised to save a copy of this .mdb file to a folder located at .../Documents/MIS342/Formative03/
    <hr /> <br />
    Basic Queries-<br />
    <br />
    Use your knowledge of Access to create a query that does the following:<br />
    When done, save the query as qryQ1<br />
    <strong>Which orders were fulfilled in calendar year 1996?</strong><br />
    Sort the results in Ascending date order, then sort by customer.<p><span class="auto-style3">1. (50) How many orders were shipped in 1996?<br />
        Run the query, then print a copy as EX1.pdf</span><font color="#0000FF">   </font></p>
    <p><b><font color="#000000">
        <input id="q1" name="q1" type="text" maxlength="120" size="120" style="background-color: #00FFFF"  /></font></b>
    
        <hr />
    Use your knowledge of Access to create a query that does the following:<br />
    When done, save the query as qryQ2<br />
    Hint: examine the 'Like' operator


    by referring to the Help menu, because you must use it in your answer.<br />
    Ref: <a href="https://support.office.com/en-us/article/Like-Operator-B2F7EF03-9085-4FFB-9829-EEF18358E931">Like Operator </a>
</p>

    <p class="auto-style3">2. (50) Which CustomerIDs that begin with FRAN contain the next letter 'R' or 'S'?<br />
        Run the query, then print a copy as EX2.pdf
    </p>
        </font>
        <input id="q2" name="q2" type="text" maxlength="120" size="120" style="background-color: #00FFFF"  /><hr />
        

<p>
    Use your knowledge of Access to create a set of queries that does the following:<br />
    When done, save the queries as qryQ3, qryQ4, qryQ5...<br />
    Ref: <a href="https://support.office.com/en-us/article/Learn-to-build-an-expression-20c385ee-accd-4306-bc7b-adf11f26948a#__toc355966595" target="_blank">Calculated Field in a Query </a>
</p>

    <p class="auto-style3">3. (50) Calculate the ExtendedPrice, which is&nbsp; (UnitPrice * Quantity) for each order in the Order Details table.<br />
        Remove the line breaks and provide the SQL statement below. <br />
        Run the query, then print a copy as EX3.pdf<br />
        </font>
        <input id="q3" name="q3" type="text" maxlength="120" size="120" style="background-color: #00FFFF"  />
</p>

   <p class="auto-style3">4. (50) What is the total ExtendedPrice (Price * Quantity) for order 11077?<br />
       Run the query, then print a copy as EX4.pdf<br />
       Hint: use the grouping functions.<br />
        </font>
        <input id="q4" name="q4" type="text" maxlength="120" size="120" style="background-color: #00FFFF"  />
   </p>

   <p class="auto-style3">5. (50) Now modify the query so that it shows not only OrderID and total of ExtendedPrice, but also a count of the number of items on each order.<br />
       Remove line breaks and provide the SQL statement below.<br />
       Run the query, then print a copy as EX5.pdf<br />
       Hint: use the grouping function again.<br />
        </font>
        <input id="q5" name="q5" type="text" maxlength="120" size="120" style="background-color: #00FFFF"  />
   </p>



        <hr />

        <p>
    Use your knowledge of Access, and the Help menu, to answer the following:<br />
    Hint: examine the 'Like' operator
</p>

    <p class="auto-style3">6. (50) What two strings will the query&nbsp;&nbsp;&nbsp; <em>Like "st[ao]re" </em>&nbsp;&nbsp; return?<br />
        Provide the answer below.<br />
                </font>
        <input id="q6" name="q6" type="text" maxlength="120" size="120" style="background-color: #00FFFF"  />
</p>

        <hr />

            <br />
    <hr />


     <p> Working with dates-</p>
    <p> Research the DateDiff() function, then create a calculated field in a query that answers the following questions.<br />
        Ref: <a href="https://support.office.com/en-us/article/DateDiff-Function-E6DD7EE6-3D01-4531-905C-E24FC238F85F" target="_blank">DateDiff()</a> 
    </p>
    <p> &nbsp; <br />
        Save your query as qryQ7</p>
    <p class="auto-style3">
        7. (50) What are the ages in year of all employees?<br />
        Remove the line breaks and provide the SQL statement below.<br />
        Run the query, then print a copy as EX7.pdf<br />
    </p>
    <input id="q7" name="q7" class="auto-style6" maxlength="120"  type="text"  />
        <hr />
    <p> &nbsp; <br />
        Save your query as qryQ8</p>
    <p class="auto-style3">
        8. (50) What is the seniority, in weeks, of each employee?<br />
        Remove the line breaks, and provide the SQL statement below.<br />
        Run the query, then print a copy as EX8.pdf<br />
    </p>
    <input id="q8" name="q8" class="auto-style6" maxlength="120"  type="text"  /><br />

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
                Use PDFill to Merge the Query Datasheet View .pdf files, save it as Formative03.pdf</p>

        <div class="col-lg-12 w3-label w3-text-blue">
        <p>(<strong>100</strong>) 9. Upload your file 'Formative03.pdf' to the D2L Assignment Folder 'Formative03_PDF&#39;</p>
            <hr />
        </div>

        <div class="col-lg-12 w3-label w3-text-blue">
        <p>(<strong>100</strong>) 10. <a href="https://support.office.com/en-us/article/Compact-and-repair-a-database-6EE60F16-AED0-40AC-BF22-85FA9F4005B2">Compact &amp; Repair</a>, and then upload a copy of your Northwind.mdb file, containing all your completed queries,  to the D2L Assignment Folder 'Formative03_MDB'.<br />
            Note: failing to compact and repair will result in a loss of points, because you will be uploading a larger .mdb file than necessary.  Your file should be about 2.5MB.
        </p>
            <hr />
        </div>


        <div class="col-lg-12 w3-label w3-text-blue">
        <p>(<strong>100</strong>) 11. Using a web browser, view your completed assignment on your website (http://classes.winona.edu/<%Response.write(YearTerm & courseid) %>...)<br />
            Once everything is complete, press the Submit button (above or below) to submit a copy of your answers to the course web database.</p>
            <hr />
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
