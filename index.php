<?php
/*
Dev note
Template - https://www.sampletemplates.com/sample-profile/simple-company-profile-template.html
Docx Replacer - https://github.com/igorrebega/docx-replacer
Simple xlsx - https://github.com/shuchkin/simplexlsx

Samples
https://stackoverflow.com/questions/4914750/how-to-zip-a-whole-folder-using-php
*/

include_once 'lib/global.func.php';
include_once 'dreplacer.class.php';

if(isset($_FILES['templateFile']) && isset($_FILES['dataFile'])){
    $upl_dir = 'temp/';
    if(!file_exists($upl_dir)) mkdir($upl_dir);

    $t_tmp_name = $_FILES["templateFile"]["tmp_name"];
    $t_name = basename($_FILES["templateFile"]["name"]);
    $t_path = "$upl_dir/$t_name";

    $d_tmp_name = $_FILES["dataFile"]["tmp_name"];
    $d_name = basename($_FILES["dataFile"]["name"]);
    $d_path = "$upl_dir/$d_name";
    
    if(move_uploaded_file($t_tmp_name, $t_path) && move_uploaded_file($d_tmp_name, $d_path)){
        $dreplacer = new Dreplacer();
        $dreplacer->process($t_path, $d_path);
    }
}

?>
<!DOCTYPE html>
<html>
<body>

<h2>Dreplacer</h2>

<form method="post" enctype="multipart/form-data">
  <label for="templateFile">Template file:</label>
  <input type="file" id="templateFile" name="templateFile" accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document"/><br><br>
  <label for="dataFile">Data file:</label>
  <input type="file" id="dataFile" name="dataFile" accept=".xlsx, .xls"/><br><br><br>
  <input type="submit" value="Submit"/>
</form>

<br/><br/><br/><br/>
<h4>Sample Files</h4>
<a href="sample_files/template_sample.docx" download>template_sample.docx</a><br><br>
<a href="sample_files/data_sample.docx" download>data_sample.docx</a>

</body>
</html>