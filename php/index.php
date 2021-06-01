<?php

require "vendor/autoload.php";
require "SdecParcer.php";

$r = new SdecParcer("../SC9DK270G3-DBLK3448.xls");
//$r = new SdecParcer("vendor/phpoffice/phpspreadsheet/samples/templates/27template.xls");
//$r->readAll();
//print_r(
//$r->readSheet("D111C")
//);
$r->readImg("D111C");