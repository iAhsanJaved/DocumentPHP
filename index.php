<?php

	require_once 'vendor/autoload.php';

	if (isset($_POST['submit'])) {
	
		$inputFile = "file";
	    $target_dir = __DIR__."/uploaded_files/";
	    $newfilename = date('dmYHis').time().str_replace(" ", "", basename($_FILES[$inputFile]["name"]));
	    $target_file = $target_dir . $newfilename;

	    if (move_uploaded_file($_FILES[$inputFile]["tmp_name"], $target_file)) {  
			if ($_POST['From'] == 'DOC' && $_POST['To'] == 'DOCX') {
				DocToDocx($newfilename);
			} else if ($_POST['From'] == 'DOC' && $_POST['To'] == 'ODT') {
				DocToODT($newfilename);
			} else if ($_POST['From'] == 'DOC' && $_POST['To'] == 'RTF') {
				DocToRTF($newfilename);
			}

			
	  	}

	}

	function DocToDocx($filename)
	{
		$filePath = __DIR__."/uploaded_files/".$filename;
		$object = new \PhpOffice\PhpWord\PhpWord();
		$phpWord = \PhpOffice\PhpWord\IOFactory::load($filePath, 'MsDoc');
		$xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
		$resultFileName = "Result.docx";
		$xmlWriter->save($resultFileName);
		unlink($filePath); // Delete uploaded file

		header('Content-Description: File Transfer');
		header('Content-Type: application/octet-stream');
		header('Content-Disposition: attachment; filename='.$resultFileName);
		header('Content-Transfer-Encoding: binary');
		header('Expires: 0');
		header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		header('Pragma: public');
		header('Content-Length: ' . filesize($resultFileName));
		flush();
		readfile($resultFileName);
		unlink($resultFileName); // deletes the temporary file
		exit;

	}

	function DocToODT($filename)
	{
		$filePath = __DIR__."/uploaded_files/".$filename;
		$object = new \PhpOffice\PhpWord\PhpWord();
		$phpWord = \PhpOffice\PhpWord\IOFactory::load($filePath, 'MsDoc');
		$xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'ODText');
		$resultFileName = "Result.odt";
		$xmlWriter->save($resultFileName);
		unlink($filePath); // Delete uploaded file

		header('Content-Description: File Transfer');
		header('Content-Type: application/octet-stream');
		header('Content-Disposition: attachment; filename='.$resultFileName);
		header('Content-Transfer-Encoding: binary');
		header('Expires: 0');
		header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		header('Pragma: public');
		header('Content-Length: ' . filesize($resultFileName));
		flush();
		readfile($resultFileName);
		unlink($resultFileName); // deletes the temporary file
		exit;

	}

	function DocToRTF($filename)
	{
		$filePath = __DIR__."/uploaded_files/".$filename;
		$object = new \PhpOffice\PhpWord\PhpWord();
		$phpWord = \PhpOffice\PhpWord\IOFactory::load($filePath, 'MsDoc');
		$xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'RTF');
		$resultFileName = "Result.rtf";
		$xmlWriter->save($resultFileName);
		unlink($filePath); // Delete uploaded file

		header('Content-Description: File Transfer');
		header('Content-Type: application/octet-stream');
		header('Content-Disposition: attachment; filename='.$resultFileName);
		header('Content-Transfer-Encoding: binary');
		header('Expires: 0');
		header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		header('Pragma: public');
		header('Content-Length: ' . filesize($resultFileName));
		flush();
		readfile($resultFileName);
		unlink($resultFileName); // deletes the temporary file
		exit;

	}

	
	

?>

<!DOCTYPE html>
<html>
<head>
	<title>Document PHP</title>
	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
	<style type="text/css">
body {
    padding: 50px;
    background: #f2f2f2;
}
h1 {
    color: #808080d6;
    letter-spacing: 0.2em;
}
h4 {
	font-size: 1.4rem;
}
input{
  padding: 10px 25px;
}
.container {
    text-align: center;
}
.container {
    box-shadow: 5px 5px 25px 0px rgba(46,61,73,0.2);
    padding: 20px;
    background: #fff;
    border-radius: 0.375rem;
}
.btn {
    display: block;
    margin-left: auto;
    margin-right: auto;
    padding: 15px;
    box-shadow: 4px 4px 4px 0 rgba(46, 61, 73, 0.13);
    border-radius: 0.25rem;
    letter-spacing: 0.09375rem;
    background: #fff;
}
form {
    border: 2px solid #eee;
    padding: 20px;
}
	</style> 		
</head>
<body>
	<div class="container">
		<h1>Document PHP</h1>
		<form action="" method="POST" enctype="multipart/form-data">
			<div class="row justify-content-center">
				<div class="form-group col-md-2">
				    <label>From</label>
				    <select name="From" class="form-control">
				      <option>DOC</option>
				    </select>
				</div>	
				<div class="form-group col-md-2">
				    <label>To</label>
				    <select name="To" class="form-control">
				      <option>DOCX</option>
				      <option>ODT</option>
				      <option>RTF</option>
				    </select>
				</div>
			</div>
			
			<div class="row justify-content-center">
				<div class="col-md-4"> 
					<input type="file" class="form-control" name="file">
				</div>
			</div>
			<div class="row mt-2">
				<div class="col"> 
					<input type="submit" class="btn" value="Convert" name="submit">
				</div>
			</div>
			
		</form>
		
		</div>
	</div>
	<script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
	<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
	<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
</body>
</html>

