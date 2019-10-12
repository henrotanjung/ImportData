<?php
require_once ('PHPExcel.php');
require_once ('PHPExcel.php');
require_once ('PHPExcel/Writer/Excel2007.php');
require_once ('PHPExcel/IOFactory.php');

$con = mysqli_connect("192.168.0.24","luki.handoko","MysticRiver20140310","his-tpa");
$con2 = mysqli_connect("192.168.0.24","luki.handoko","MysticRiver20140310","her-tpa");

// Check connection
if (mysqli_connect_errno())
  {
  echo "Failed to connect to MySQL: " . mysqli_connect_error();
  };

$r_batch = mysqli_query($con2, "SELECT MAX(log_upload_module.batch) + 1 FROM log_upload_module");
while ($data2 = mysqli_fetch_row($r_batch)) {
  $batch = $data2[0];
}

$date =date('Ymd');
$file='upload_payment_'.$date.'';
$inputFileName = '//192.168.0.20/Z Folder/Central Billing/INTERNAL ONLY/FINANCE CBD/UPLOAD_CASE/'.$file.'.xlsx';
$file2='upload_payment_result_'.$date.'_'.$batch.'.xlsx';


$Excel = new PHPExcel();
$Excel->setActiveSheetIndex(0);
$Excel->getActiveSheet()->setCellValue('A1', 'CASE ID');
$Excel->getActiveSheet()->setCellValue('B1', 'PATIENT');
$Excel->getActiveSheet()->setCellValue('C1', 'PAYMENT DATE');
$Excel->getActiveSheet()->setCellValue('D1', 'PAYMENT BATCHING');
$Excel->getActiveSheet()->setCellValue('E1', 'PAID TO PROVIDER');
$Excel->getActiveSheet()->setCellValue('F1', 'HOSPITAL');
$Excel->getActiveSheet()->setCellValue('G1', 'STATUS');
$Excel->getActiveSheet()->setCellValue('H1', 'BILL NO');
$Excel->getActiveSheet()->setCellValue('I1', 'PAYMENT APPROVAL DATE');
$Excel->getActiveSheet()->setCellValue('J1', 'INVOICE NO');
$Excel->getActiveSheet()->setCellValue('K1', 'CLIENT');
$Excel->getActiveSheet()->setCellValue('L1', 'ATTACHMENT');
$Excel->getActiveSheet()->setCellValue('M1', 'PAID BY');
$Excel->getActiveSheet()->setCellValue('N1', 'REMARKS');

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
      //
      $case_id = $data[0];
      $patient = $data[1];
      $payment_date = $data[2];
      $payment_batching = $data[3];
      $paid_to_provider = $data[4];
      $hospital = $data[5];
      $bill_no = $data[7];
      if (is_null($data[8])) {
        $payment_approval_date = 'NULL';
      } else {
        $payment_approval_date = "'$data[8]'";
      }
      $payment_approval_datex = $data[8];
      $invoice_no = $data[9];
      $client = $data[10];
      $attachment = $data[11];
      $paid_by = $data[12];
      $code = "FINANCE/PAYMENT/".$date."";

      if (empty($case_id) || is_null($case_id)) {
        $remarks = "Error [Case ID cannot NULL]";
        $status = $data[6];
      } else {
        //
        $v_case = "SELECT `case`.id,`case`.`status`,case_status.`name` FROM `case` LEFT JOIN case_status ON `case`.`status` = case_status.`status`WHERE case_status.userlevels = - 1 AND `case`.id = ".$case_id.";";
        $r_case = mysqli_query($con, $v_case);
        $count = mysqli_num_rows($r_case);
        if ($count != 1) {
          $remarks = "Error [Cannot found case id ".$case_id."]";
          $status = $data[6];
        } else {
          while ($data1 = mysqli_fetch_row($r_case)) {
            if ($data1[1] != 16 && $data1[1] != 27) {
              $remarks = "Error [Cannot process payment in status ".$data1[2]."]";
              $status = $data1[2];
            } else {
              $update = "UPDATE `case` SET `case`.edited_by = '".$paid_by."', `case`.edit_date = NOW() WHERE `case`.id = ".$case_id.";";
              $r_update = mysqli_query($con, $update);
              if ($r_update) {
                $remarks = "Success";
                $status = $data[6];
                $update1 = "UPDATE `case` SET `case`.payment_date = '".$payment_date."', `case`.client_approval_for_payment = ".$payment_approval_date.", `case`.upload_proof_of_payment = '".$attachment."' WHERE `case`.id = ".$case_id.";";
                $r_update1 = mysqli_query($con, $update1);
                if (mysqli_affected_rows($con)) {
                  $src = "//192.168.0.20/Z Folder/Central Billing/INTERNAL ONLY/FINANCE CBD/UPLOAD_CASE/proof_of_payment/".$attachment."";
                  $dst = "//10.10.10.27/c$/xampp/htdocs/insecure/intsys/his-tpa/AAint/upload/".$case_id."/".$attachment."";
                  //$dst = "C:/xampp/htdocs/insecure/intsys/finance_module/dest/".$attachment."";
                  copy($src, $dst);
                }
              } else {
                $remarks = "Error [Please check excel format]";
                $status = $data1[2];
              }
            }
          }
        }
      }
      $insert = "INSERT INTO `log_upload_module` (
                        `code`,
                        `case`,
                        `payment_date`,
                        `upload_proof_of_payment`,
                        `client_approval_for_payment`,
                        `paid_by`,
                        `remarks`,
                        `process_date`
                      )
                      VALUES
                        (
                          '".$code."',
                          ".$case_id.",
                          '".$payment_date."',
                          '".$attachment."',
                          ".$payment_approval_date.",
                          '".$paid_by."',
                          '".$remarks."',
                          NOW()
                        );
      ";
      $r_insert = mysqli_query($con2, $insert);
      $Excel->getActiveSheet()->setCellValue('A'.$row, $case_id);
      $Excel->getActiveSheet()->setCellValue('B'.$row, $patient);
      $Excel->getActiveSheet()->setCellValue('C'.$row, $payment_date);
      $Excel->getActiveSheet()->setCellValue('D'.$row, $payment_batching);
      $Excel->getActiveSheet()->setCellValue('E'.$row, $paid_to_provider);
      $Excel->getActiveSheet()->setCellValue('F'.$row, $hospital);
      $Excel->getActiveSheet()->setCellValue('G'.$row, $status);
      $Excel->getActiveSheet()->setCellValue('H'.$row, $bill_no);
      $Excel->getActiveSheet()->setCellValue('I'.$row, $payment_approval_datex);
      $Excel->getActiveSheet()->setCellValue('J'.$row, $invoice_no);
      $Excel->getActiveSheet()->setCellValue('K'.$row, $client);
      $Excel->getActiveSheet()->setCellValue('L'.$row, $attachment);
      $Excel->getActiveSheet()->setCellValue('M'.$row, $paid_by);
      $Excel->getActiveSheet()->setCellValue('N'.$row, $remarks);
    }
  }
}
$Excel->getActiveSheet()->setTitle('Result');
$objWriter = PHPExcel_IOFactory::createWriter($Excel, 'Excel2007');
try {
  $objWriter->save('//192.168.0.20/Z Folder/Central Billing/INTERNAL ONLY/FINANCE CBD/UPLOAD_CASE/result/'.$file2);
  $update_batch = mysqli_query($con2, "UPDATE log_upload_module SET log_upload_module.batch = ".$batch." WHERE log_upload_module.batch IS NULL");
  unlink('//192.168.0.20/Z Folder/Central Billing/INTERNAL ONLY/FINANCE CBD/UPLOAD_CASE/'.$file.'.xlsx');
  echo '<script>alert("Please check result file");window.location.href="index.html";</script>';
} catch(Exception $e) {
  echo '<script>alert("Error while creating result");window.location.href="index.html";</script>';
}
