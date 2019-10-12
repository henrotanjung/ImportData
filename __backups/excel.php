<?php
require_once ('PHPExcel.php');
require_once ('PHPExcel.php');
require_once ('PHPExcel/Writer/Excel2007.php');
require_once ('PHPExcel/IOFactory.php');

$con = mysqli_connect("192.168.0.24","luki.handoko","MysticRiver20140310","his-tpa");

// Check connection
if (mysqli_connect_errno())
  {
  echo "Failed to connect to MySQL: " . mysqli_connect_error();
  };

$file=date('Ymd');
$inputFileName = 'Y:/ADMIN/EXTERNAL USE/UPLOAD_FILE/'.$file.'.xlsx';
//date_default_timezone_set('Asia/Jakarta');

//  Read excel workbook
try {
    $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load($inputFileName);
} catch(Exception $e) {
    die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
}

//  Get worksheet dimensions
$sheet = $objPHPExcel->getSheet(0); 
$highestRow = $sheet->getHighestRow(); 
$highestColumn = $sheet->getHighestColumn();
$headingsArray = $sheet->rangeToArray('A1:'.$highestColumn.'1',null, true, true, true);
$headers = $headingsArray[1];
echo "Berikut data case yang berhasil diupdate : <br/>";
for ($row = 2; $row <= $highestRow;++$row){
	$dataRow = $sheet->rangeToArray('A'.$row.':'.$highestColumn.$row,null, true, true, false);
	if (is_array($dataRow)) {
		foreach ($dataRow as $key => $data) {
			
			$update = "UPDATE `his-tpa`.`case` SET
						`his-tpa`.`case`.userlevel = -1,
						`his-tpa`.`case`.currency_rate_01_to_idr = 1,
						`his-tpa`.`case`.currency_rate_idr_to_02 = 1,
						`his-tpa`.`case`.bill_no = '".$data[1]."',
						`his-tpa`.`case`.bill_issue_date = NOW(),
						`his-tpa`.`case`.bill_due_date = IF(`his-tpa`.`case`.category = 0, DATE_ADD(NOW(),INTERVAL 7 DAY), DATE_ADD(NOW(),INTERVAL 14 DAY))
						WHERE
						`his-tpa`.`case`.status in (11,12,19,20) AND
						`his-tpa`.`case`.id = ".$data[0]."";
			$result = mysqli_query($con,$update) or die(mysqli_error($con));
			if ($result) {
				echo $data[0]."<br/>";
			}
		}	
	}
}