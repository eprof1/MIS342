<!doctype html>
<html lang="en">

<head>
    <meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Bootstrap -->
	<link href="../css/bootstrap.css" rel="stylesheet">

    <!-- constant file to update email, firstname, lastname, starid, section on students' pages -->
    <script src="../js/constants.js"></script>

    <!-- model MIS342 Formative file, version 1 released 2018/07/01 PgP -->

    <title>MIS342Formative08</title>
</head>

<body onload="setVariables()"> 
     <form class="container-fluid" action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmAssignment" id="frmAssignment">
    <input name="InstID" id="InstID" value="00617282" type="hidden">

    <h1 class=" text-center">Access and Linked Tables!</h1>

<!-- -start-Student, class and assignment information. Assignment hard-coded, other values from constants.js file-start -->
    <div class="row col-lg-12">
      <div class="col-lg-3">
        <label class="text-primary" for="FirstName">First Name:</label>
        <input class="form-control" name="FirstName" id="FirstName"  type="text" value="default first name" readonly="readonly">
      </div>
      <div class="col-lg-3">
        <label class="text-primary" for="LastName">Last Name:</label>
        <input class="form-control" name="LastName" id="LastName"  type="text" value="default last name" readonly="readonly">
      </div>
      <div class="col-lg-3">
        <label class="text-primary" for="pin">StarID:</label>
        <input class="form-control" name="pin" id="pin"  type="text" value="default StarID" readonly="readonly">
      </div>
      <div class="col-lg-3">
        <label class="text-primary" for="email">Email:</label>
        <input class="form-control" name="email" id="email"  type="email" value="default email address" readonly="readonly">
      </div>
    </div>

    <div class="row col-lg-12">
      <div class="col-lg-3">
        <label class="text-primary" for="semester">Semester:</label>
        <input class="form-control" name="semester" id="semester"  type="text" value="default" readonly="readonly">
     </div>
      <div class="col-lg-3">
        <label class="text-primary" for="class">Class:</label>
        <input class="form-control" name="class" id="class" value="xClass" type="text" readonly="readonly">
      </div>
      <div class="col-lg-3">
        <label class="text-primary" for="section">Section:</label>
        <input class="form-control" name="section" id="section"  value="default" type="text"  readonly="readonly">
      </div>
      <div class="col-lg-3"> 
        <label class="text-primary" for="assignment">Assignment:</label>
        <input class="form-control" name="assignment" id="assignment" value="Formative08" type="text" readonly="readonly">
      </div>    
    </div>
<!-- -end-Student, class and assignment information-end -->



<!-- Calculate the correct path for saving the files -->
<script>Path = startPath + xStarID + "&#92;OneDrive - MNSCU&#92;" + xClass + "&#92;" + document.getElementById("assignment").value + "&#92;"; </script> 


        <div class="col-lg-12">
            <br /><input class="text-danger btn-block btn-lg" name="Submit" value="Submit" type="submit"> <br />
        </div>
  
        <div class="col-lg-12 h3 text-danger">
            300 Points <hr />
        </div>

         <p>Note: 
  Save this  assignment to your website.<br />
    </p>
    <p>Complete your personal information in the above section.<br />
    </p>
    
         <!-- -start-objectives-start -->
    <div class="col-lg-12"><a href="#Objectives" class="btn btn-info" data-toggle="collapse">Learning Objectives</a></div>      
      <div id="Objectives" class=" collapse in col-lg-12">

          </div>
         <!--End of Objectives-->

         <div class=" col-lg-12"><a href="#Exercise01" class="btn btn-info" data-toggle="collapse">Exercise 1 - Access and Linked Tables </a></div>        
      
       <div id="Exercise01" class="collapse in  col-lg-12"> <br />

            <p>Answer the following questions by copying the contents of the "LinkedTableMadnessProject" folder which is located on the network drive at: </p>
    ("\\store\classes\" & YearTerm & courseid & "\ReadOnly\" ) to a location on your computer.<br />
    <p>
    Refer to the PowerPoint "<a href="../../Presentations/Chapter08.ppt">DataIntegrationTips</a>" which is also on the course website 
  </p>


            <p>You are provided with an Access Database, Linked.accdb, and several files.</p>
    <br />Open Linked.accdb, then open the first table in the list, CATEGORI.
    <br />

             <br />You will be presented with the following error message<br /> 
    <img alt="link error" class="auto-style3" src="LinkedTableError.JPG" /><br />
         <br />
         <br />

         <p><strong>Right click the table CATEGORI, select "Linked Table Manager" from the menu, check CATEGORI, click OK, then find the current location of the file.<br />
      Once you have found the correct file select it, click 'Open' and you will be presented with this message:  </strong><br />
     <img alt = "link found" src="LinkedTableFound.JPG" /><br />

          <p><strong> Hint link found" src="LinkedTableFound.JPG" /><br />

          <p><strong> Hint "120" value="table location:" />
    </font></b></p>
     <hr />
  <p><strong> <font color="#0000FF">3. (50)</font></strong> <font color="#0000FF"><b>What is one major difference in how Access stores a table versus how dBase stores a table?</b></font><b><font color="#000000"><br />
    <input name="q3" type="text" id="q3" size="120" maxlength="120" value="Access vs. dBase:" />
    </font></b></p>


     <hr />
     <p>4. (<strong>150</strong>) Once you have successfully relinked all tables, upload your Linked.accdb file to the D2L Assignment Folder &quot;Formative08_Link&quot;.</p>
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
