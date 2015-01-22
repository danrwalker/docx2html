<?php

	require_once sprintf('%s/../src/docx2html.class.php',dirname(__FILE__));

	//Example Docx file path
	$dirDocxFilePath = sprintf('%s/example.docx',dirname(__FILE__));

	//Get an instance of the docx2html class
	$resMyDoc = new docx2html();

	//Load in the demo docx file
	$resMyDoc->load($dirDocxFilePath);

	echo 'DocX XML Content:<br><pre>'.$resMyDoc->xml().'</pre><hr>';

	echo 'DocX Plain Content:<br><pre>'.$resMyDoc->plain().'</pre><hr>';