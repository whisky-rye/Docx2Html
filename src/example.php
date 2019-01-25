<?php
include('Docx2Html.php');

$Docx2Html = new Docx2Html('test1');  // test2 test3
$html = $Docx2Html->transferToHtml();
// print_r($html);
$url = $Docx2Html->writeToHtml($html);
print_r($url);

$Docx2Html->deleteTemp();





