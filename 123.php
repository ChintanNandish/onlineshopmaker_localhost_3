<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$str = file_get_contents('user_folders/chintan_admin/product_data.json');
$json = json_decode($str, TRUE);
$shopname = array_keys($json)[0];
$product = $json[$shopname];
$count = 0;
$keys = array_keys($product[array_keys($product)[0]]);
unset($keys[array_search('product_image',$keys)]);
foreach($keys as $k){
	$sheet->setCellValue(chr(ord('A')+$count).'1', $k);
	$count++;
}
$count=0;
$number = 2;

foreach($product as $p){
	unset($p['product_image']);
	foreach($p as $values){
		if($values != ''){
			$sheet->setCellValue(chr(ord('A')+$count).$number, $values);
		}
		else{
			$sheet->setCellValue(chr(ord('A')+$count).$number, 'N.A');
		}
		$count++;
	}
	$count = 0;
	$number++;
}


$writer = new Xlsx($spreadsheet);
$writer->save('user_folders/chintan_admin/'.$shopname.'_product_list.xlsx');


?>