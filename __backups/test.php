<?php
require_once ('PHPExcel.php');
require_once ('PHPExcel.php');
require_once ('PHPExcel/Writer/Excel2007.php');
require_once ('PHPExcel/IOFactory.php');

$con = mysqli_connect("192.168.0.24","luki.handoko","MysticRiver20140310","his-tpa-dev");

// Check connection
if (mysqli_connect_errno())
  {
  echo "Failed to connect to MySQL: " . mysqli_connect_error();
  };

$date =date('Ymd');
$file='verify_case_'.$date.'';
$inputFileName = 'C:/xampp/htdocs/insecure/intsys/finance_module/file/'.$file.'.xlsx';
$file2='verify_case_result_'.$date.'.xlsx';


$Excel = new PHPExcel();
$Excel->setActiveSheetIndex(0);
$Excel->getActiveSheet()->setCellValue('A1', 'CASE ID');
$Excel->getActiveSheet()->setCellValue('B1', 'PATIENT');
$Excel->getActiveSheet()->setCellValue('C1', 'HOSPITAL');
$Excel->getActiveSheet()->setCellValue('D1', 'CATEGORY');
$Excel->getActiveSheet()->setCellValue('E1', 'TYPE');
$Excel->getActiveSheet()->setCellValue('F1', 'COVER');
$Excel->getActiveSheet()->setCellValue('G1', 'STATUS');
$Excel->getActiveSheet()->setCellValue('H1', 'OBV DATE');
$Excel->getActiveSheet()->setCellValue('I1', 'OBV BY');
$Excel->getActiveSheet()->setCellValue('J1', 'REMARKS');

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
for ($row = 2; $row <= $highestRow;$row++){
	$dataRow = $sheet->rangeToArray('A'.$row.':'.$highestColumn.$row,null, true, true, false);
	if (is_array($dataRow)) {
		foreach ($dataRow as $key => $data) {
      $case_id = $data[0];
      $patient = $data[1];
      $hospital = $data[2];
      $category = $data[3];
      $type = $data[4];
      $cover = $data[5];
      $obv_date = $data[7];
      $obv_by = $data[8];
      if (empty($case_id) || is_null($case_id)) {
        $remarks = "Case ID cannot be null";
        $status = $data[6];
      } else {
        $v_case = "SELECT `case`.id, `case`.`status` FROM `case` WHERE `case`.id =  ".$case_id."";
        $r_case = mysqli_query($con, $v_case);
        $count = mysqli_num_rows($r_case);
        if ($count != 1) {
          $remarks = "Case ID not found";
          $status = $data[6];
        } else {
          $v_status = mysqli_query($con, "SELECT
                                      `case`.id,
                                      `case`.`status`,
                                      case_status.`name`
                                    FROM
                                      `case`
                                    LEFT JOIN case_status ON `case`.`status` = case_status.`status`
                                    WHERE
                                      case_status.userlevels = - 1
                                    AND `case`.id = ".$case_id."");
          while ($data2 = mysqli_fetch_row($v_status)) {
            if ($data2[1] <> 15 && $data2[1] <> 26) {
              $remarks = "Please check case status";
              $status = $data2[2];
            } else {
              $update1 = "UPDATE `case` SET `case`.edited_by = '".$obv_by."', `case`.edit_date = NOW() WHERE `case`.id = ".$case_id.";";
              $r_update1 = mysqli_query($con, $update1);
              if ($r_update1) {
                $update2 = "UPDATE `case` SET `case`.original_bill_verified = 4 WHERE `case`.id = ".$case_id.";";
                $r_update2 = mysqli_query($con, $update2);
                if ($r_update2) {
                  $remarks = "Success";
                  $get_status = "SELECT
                                    `case`.id,
                                    case_status.`name`
                                  FROM
                                    `case`
                                  LEFT JOIN case_status ON `case`.`status` = case_status.`status`
                                  WHERE
                                    case_status.userlevels = - 1
                                  AND `case`.id = ".$case_id."";
                  $r_status = mysqli_query($con, $get_status);
                  while ($data1 = mysqli_fetch_row($r_status)) {
                    $status = $data1[1];
                  }
                }
              }
            }
          }
        }
      }
      $Excel->getActiveSheet()->setCellValue('A'.$row, $case_id);
      $Excel->getActiveSheet()->setCellValue('B'.$row, $patient);
      $Excel->getActiveSheet()->setCellValue('C'.$row, $hospital);
      $Excel->getActiveSheet()->setCellValue('D'.$row, $category);
      $Excel->getActiveSheet()->setCellValue('E'.$row, $type);
      $Excel->getActiveSheet()->setCellValue('F'.$row, $cover);
      $Excel->getActiveSheet()->setCellValue('G'.$row, $status);
      $Excel->getActiveSheet()->setCellValue('H'.$row, $obv_date);
      $Excel->getActiveSheet()->setCellValue('I'.$row, $obv_by);
      $Excel->getActiveSheet()->setCellValue('J'.$row, $remarks);
    }
  }
}
$Excel->getActiveSheet()->setTitle('Result');
$objWriter = PHPExcel_IOFactory::createWriter($Excel, 'Excel2007');
$objWriter->save('C:/xampp/htdocs/insecure/intsys/finance_module/file/'.$file2);
