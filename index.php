<?php

	require_once 'vendor/autoload.php';

	if (isset($_POST['_DocToDocx'])) {
	
		$inputFile = "file";
	    $target_dir = __DIR__."/uploaded_files/";
	    $newfilename = date('dmYHis').time().str_replace(" ", "", basename($_FILES[$inputFile]["name"]));
	    $target_file = $target_dir . $newfilename;

	    if (move_uploaded_file($_FILES[$inputFile]["tmp_name"], $target_file)) {  
			DocToDocx($newfilename);
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

	
	

?>

<!DOCTYPE html>
<html>
<head>
	<title>Document PHP</title>
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
	</style> 		
</head>
<body>
	<div class="container">
		<h1>Document PHP</h1>
		<hr>
		<h4>Doc to Docx</h4>
		<form action="" method="POST" enctype="multipart/form-data">
			<input type="file" name="file">
			<input type="submit" class="btn" value="Convert" name="_DocToDocx">
		</form>
		<hr>
		
		</div>
	</div>
</body>
</html>

