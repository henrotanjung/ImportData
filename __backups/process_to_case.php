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

$date =date('Ymd');
$file='create_case_result_'.$date.'';
$inputFileName = 'Y:/ADMIN/EXTERNAL USE/UPLOAD_CASE/'.$file.'.xlsx';
$file2='process_to_case_result_'.$date.'.xlsx';


$Excel = new PHPExcel();
$Excel->setActiveSheetIndex(0);
$Excel->getActiveSheet()->setCellValue('A1', 'Id');
$Excel->getActiveSheet()->setCellValue('B1', 'Status');
$Excel->getActiveSheet()->setCellValue('C1', 'Create by');
$Excel->getActiveSheet()->setCellValue('D1', 'Receive Date');
$Excel->getActiveSheet()->setCellValue('E1', 'Category');
$Excel->getActiveSheet()->setCellValue('F1', 'Type');
$Excel->getActiveSheet()->setCellValue('G1', 'Admission Date');
$Excel->getActiveSheet()->setCellValue('H1', 'Admission Time');
$Excel->getActiveSheet()->setCellValue('I1', 'Discharge Date');
$Excel->getActiveSheet()->setCellValue('J1', 'Discharge Doctor');
$Excel->getActiveSheet()->setCellValue('K1', 'Diagnosis');
$Excel->getActiveSheet()->setCellValue('L1', 'Bill No');
$Excel->getActiveSheet()->setCellValue('M1', 'Bill Issue Date');
$Excel->getActiveSheet()->setCellValue('N1', 'Bill Due Date');
$Excel->getActiveSheet()->setCellValue('O1', 'Provider');
$Excel->getActiveSheet()->setCellValue('P1', 'Other Provider');
$Excel->getActiveSheet()->setCellValue('Q1', 'Policy No');
$Excel->getActiveSheet()->setCellValue('R1', 'Dob');
$Excel->getActiveSheet()->setCellValue('S1', 'Client');
$Excel->getActiveSheet()->setCellValue('T1', 'Patient Name');
$Excel->getActiveSheet()->setCellValue('U1', 'Principle');
$Excel->getActiveSheet()->setCellValue('V1', 'Case ID');
$Excel->getActiveSheet()->setCellValue('W1', 'Case Ref');
$Excel->getActiveSheet()->setCellValue('X1', 'Remarks');

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
      $create_byx = $data[2];
      $receive_datex = $data[3];
      $categoryx = $data[4];
      $typex = $data[5];
      $admission_datex = $data[6];
      $admission_timex = $data[7];
      $discharge_datex = $data[8];
      $discharge_doctorx = $data[9];
      $diagnosisx = $data[10];
      $bill_nox = $data[11];
      $bill_issue_datex = $data[12];
      $bill_due_datex = $data[13];
      $providerx = $data[15];
      $other_providerx = $data[16];
      $policy_nox = $data[19];
      $dobx = $data[20];
      $clientx = $data[21];
      $patientx = $data[25];
      $principlex = $data[28];
      $case_id = '';
      $ref = '';
      $status_1 = $data[1];
      if (empty($data[0]) || is_null($data[0])) {
        $msg = "ID cannot null";
      } else {
          $create_case = mysqli_query($con, "SELECT
          r_create_case.id,
          r_create_case.`status`,
          r_create_case.created_by,
          r_create_case.receive_date,
          r_create_case.category,
          r_create_case.type,
          r_create_case.admission_date,
          r_create_case.admission_time,
          r_create_case.discharge_date,
          r_create_case.discharge_doctor,
          r_create_case.diagnosis,
          r_create_case.bill_no,
          r_create_case.bill_issue_date,
          r_create_case.bill_due_date,
          r_create_case.id_provider,
          r_create_case.provider_error_remarks,
          r_create_case.provider,
          r_create_case.other_provider,
          r_create_case.other_provider_city,
          r_create_case.other_provider_country,
          r_create_case.policy_no,
          r_create_case.dob,
          r_create_case.client,
          r_create_case.company,
          r_create_case.branch,
          r_create_case.patient,
          r_create_case.patient_error_remarks,
          r_create_case.member_id,
          r_create_case.member_card,
          r_create_case.principle,
          r_create_case.gender,
          r_create_case.relation,
          r_create_case.policy_status,
          r_create_case.policy_holder,
          r_create_case.policy_effective_date,
          r_create_case.policy_expiry_date,
          r_create_case.policy_issue_date,
          r_create_case.policy_declare_date,
          r_create_case.policy_lapse_date,
          r_create_case.policy_revival_date,
          r_create_case.policy_termination_date,
          r_create_case.policy_suspend_date,
          r_create_case.policy_unsuspend_date,
          r_create_case.program,
          r_create_case.plan,
          r_create_case.plan_attach_date,
          r_create_case.plan_expiry_date,
          r_create_case.rider,
          r_create_case.rider_attach_date,
          r_create_case.rider_expiry_date,
          r_create_case.special_condition,
          r_create_case.exclusion,
          r_create_case.create_date,
          r_create_case.edit_date,
          r_create_case.ref,
          client.full_name as client_name,
          member.member_name as member_name,
          principle.member_name as principle_name,
          program.close_case_option,
          program.doc_send_back_required
          FROM
          r_create_case
          LEFT JOIN client ON r_create_case.client = client.id
          LEFT JOIN member ON r_create_case.patient = member.id
          LEFT JOIN principle ON r_create_case.principle = principle.id
          LEFT JOIN program ON r_create_case.program = program.id
          WHERE
          r_create_case.id = ".$data[0]."") or die(mysqli_error($con));
        $count = mysqli_num_rows($create_case);
        if ($count != 1) {
          $msg = "ID Not Found";
        } else {
          while ($data1 = mysqli_fetch_row($create_case)) {
            $created_by = $data1[2];
            $create_date = date('Y-m-d H:i:s');
            $created_by_01 = $data1[2];
            $create_date_01 = $data1[3];
            $source = 0;
            $status = 19;
            $category = $data1[4];
            $type = $data1[5];
            $patient = $data1[25];
            $provider = $data1[14];
            $admission_date = $data1[6];
            $ref = "0".$category."".$type."/".$patient."/".$provider."/".$admission_date."";
            $dob = $data1[21];
            $gender = $data1[30];
            $member_id = $data1[27];
            $member_card = $data1[28];
            $principle = $data1[29];
            $relation = $data1[31];
            $client = $data1[22];
            if (is_null($data1[24])) {
              $branch = 'NULL';
            } else {
              $branch = $data1[24];
            }
            if (is_null($data1[23])) {
              $company = 'NULL';
            } else {
              $company = $data1[23];
            }
            $policy_status = $data1[32];
            $policy_no = $data1[20];
            $policy_holder = $data1[33];
            $policy_effective_date = $data1[34];
            $policy_expiry_date = $data1[35];
            $policy_issue_date = $data1[36];
            $policy_declare_date = $data1[37];
            if (is_null($data1[38])) {
              $policy_lapse_date = 'NULL';
            } else {
              $policy_lapse_date = '"$data1[38]"';
            }
            if (is_null($data1[39])) {
              $policy_revival_date = 'NULL';
            } else {
              $policy_revival_date = '"$data1[38]"';
            }
            if (is_null($data1[40])) {
              $policy_termination_date = 'NULL';
            } else {
              $policy_termination_date = '"$data1[40]"';
            }
            if (is_null($data1[41])) {
              $policy_suspend_date = 'NULL';
            } else {
              $policy_suspend_date = '"$data1[41]"';
            }
            if (is_null($data1[42])) {
              $policy_unsuspend_date = 'NULL';
            } else {
              $policy_unsuspend_date = '"$data1[42]"';
            }
            $exclusion = $data1[51];
            $special_condition = $data1[50];
            $program = $data1[43];
            $plan = $data1[44];
            $plan_attach_date = $data1[45];
            $plan_expiry_date = $data1[46];
            if (is_null($data1[47])) {
              $rider = 'NULL';
            } else {
              $rider = $data1[47];
            }
            if (is_null($data1[48])) {
              $rider_attach_date = 'NULL';
            } else {
              $rider_attach_date = '"$data1[48]"';
            }
            if (is_null($data1[49])) {
              $rider_expiry_date = 'NULL';
            } else {
              $rider_expiry_date = '"$data1[49]"';
            }
            $other_provider = $data1[17];
            $other_provider_city = $data1[18];
            $other_provider_country = $data1[19];
            $admission_time = $data1[7];
            $discharge_date = $data1[8];
            $discharge_doctor = $data1[9];
            $diagnosis = $data1[10];
            $bill_no = $data1[11];
            $bill_issue_date = $data1[12];
            $bill_due_date = $data1[13];
            $edited_by = $data1[2];
            $edit_date = date('Y-m-d H:i:s');

            $client_name = $data1[55];
            $patient_name = $data1[56];
            $principle_name = $data1[57];

            $provider_cancel = 0;
            $currency_01 = 85;
            $currency_02 = 85;
            $currency_rate = 1;
            $currency_rate_01_to_idr = 0;
            $currency_rate_idr_to_02 = 0;
            $amount_currency_01 = 0;
            $upload_haf_doc_finished = 0;
            $haf_doc_completed = 1;
            $issue_initial_log = 1;
            $bill_received = 0;
            $upload_bill_finished = 0;
            $bill_completed = 1;
            $ws_type = 0;
            $ws_finished = 0;
            $ws_approval = 1;
            $issue_log = 1;
            $issue_log_option = 0;
            $original_bill_checked = 1;
            $close_case_option = $data1[58];
            $original_bill_verified = 1;
            $doc_send_back_to_client_required = $data1[59];
 
            if ($data1[1] == 2) {
              $msg = "ID not verified yet";
              $status_1 = $data1[1];
            } elseif ($data1[1] == 3) {
              $msg = "The case is redundant";
              $status_1 = $data1[1];
            } else {
              $cek_ref = "SELECT `case`.ref FROM `case` WHERE `case`.ref = '".$ref."';";
              $r_cek_ref = mysqli_query($con, $cek_ref) or die(mysqli_error($con));
              $count1 = mysqli_num_rows($r_cek_ref);
              if ($count1 > 0) {
                $msg = "The case is redundant";
                $status_1 = $data1[1];
              } elseif ($policy_lapse_date != 'NULL' && $admission_date > $policy_lapse_date) {
                $msg = "Admission date cannot be newer than policy lapse date";
                $status_1 = $data1[1];
              } elseif ($policy_termination_date != 'NULL' && $admission_date > $policy_termination_date) {
                $msg = "Admission date cannot be newer than policy termination date";
                $status_1 = $data1[1];
              } elseif ($admission_date > $plan_expiry_date) {
                $msg = "Admission date cannot be newer than plan expiry date";
                $status_1 = $data1[1];
              } elseif ($admission_date < $policy_effective_date) {
                $msg = "Admission date cannot be lower than policy effective date";
                $status_1 = $data1[1];
              } else {
                $status_1 = 3;
                $msg = "Success";
                $insert = "INSERT INTO `case` (
                      `created_by`,
                      `create_date`,
                      `created_by_01`,
                      `create_date_01`,
                      `source`,
                      `status`,
                      `category`,
                      `type`,
                      `ref`,
                      `patient`,
                      `dob`,
                      `gender`,
                      `member_id`,
                      `member_card`,
                      `principle`,
                      `relation`,
                      `client`,
                      `branch`,
                      `company`,
                      `policy_status`,
                      `policy_no`,
                      `policy_holder`,
                      `policy_effective_date`,
                      `policy_expiry_date`,
                      `policy_issue_date`,
                      `policy_declare_date`,
                      `policy_lapse_date`,
                      `policy_revival_date`,
                      `policy_termination_date`,
                      `policy_suspend_date`,
                      `policy_unsuspend_date`,
                      `exclusion`,
                      `special_condition`,
                      `program`,
                      `plan`,
                      `plan_attach_date`,
                      `plan_expiry_date`,
                      `rider`,
                      `rider_attach_date`,
                      `rider_expiry_date`,
                      `provider`,
                      `other_provider`,
                      `other_provider_city`,
                      `other_provider_country`,
                      `admission_date`,
                      `admission_time`,
                      `discharge_date`,
                      `discharge_doctor`,
                      `diagnosis`,
                      `bill_no`,
                      `bill_issue_date`,
                      `bill_due_date`,
                      `provider_cancel`,
                      `currency_01`,
                      `currency_02`,
                      `currency_rate`,
                      `currency_rate_01_to_idr`,
                      `currency_rate_idr_to_02`,
                      `amount_currency_01`,
                      `upload_haf_doc_finished`,
                      `haf_doc_completed`,
                      `issue_initial_log`,
                      `bill_received`,
                      `upload_bill_finished`,
                      `bill_completed`,
                      `ws_type`,
                      `ws_finished`,
                      `ws_approval`,
                      `issue_log`,
                      `issue_log_option`,
                      `original_bill_checked`,
                      `close_case_option`,
                      `original_bill_verified`,
                      `doc_send_back_to_client_required`,
                      `edited_by`,
                      `edit_date`,
                      `userlevel`
                    )
                    VALUES
                      (
                        '".$created_by."',
                        '".$create_date."',
                        '".$created_by_01."',
                        '".$create_date_01."',
                        ".$source.",
                        ".$status.",
                        ".$category.",
                        ".$type.",
                        '".$ref."',
                        ".$patient.",
                        '".$dob."',
                        ".$gender.",
                        '".$member_id."',
                        '".$member_card."',
                        ".$principle.",
                        ".$relation.",
                        ".$client.",
                        ".$branch.",
                        ".$company.",
                        ".$policy_status.",
                        '".$policy_no."',
                        '".$policy_holder."',
                        '".$policy_effective_date."',
                        '".$policy_expiry_date."',
                        '".$policy_issue_date."',
                        '".$policy_declare_date."',
                        ".$policy_lapse_date.",
                        ".$policy_revival_date.",
                        ".$policy_termination_date.",
                        ".$policy_suspend_date.",
                        ".$policy_unsuspend_date.",
                        '".$exclusion."',
                        '".$special_condition."',
                        ".$program.",
                        ".$plan.",
                        '".$plan_attach_date."',
                        '".$plan_expiry_date."',
                        ".$rider.",
                        ".$rider_attach_date.",
                        ".$rider_expiry_date.",
                        ".$provider.",
                        '".$other_provider."',
                        '".$other_provider_city."',
                        ".$other_provider_country.",
                        '".$admission_date."',
                        '".$admission_time."',
                        '".$discharge_date."',
                        '".$discharge_doctor."',
                        '".$diagnosis."',
                        '".$bill_no."',
                        '".$bill_issue_date."',
                        '".$bill_due_date."',
                        ".$provider_cancel.",
                        ".$currency_01.",
                        ".$currency_02.",
                        ".$currency_rate.",
                        ".$currency_rate_01_to_idr.",
                        ".$currency_rate_idr_to_02.",
                        ".$amount_currency_01.",
                        ".$upload_haf_doc_finished.",
                        ".$haf_doc_completed.",
                        ".$issue_initial_log.",
                        ".$bill_received.",
                        ".$upload_bill_finished.",
                        ".$bill_completed.",
                        ".$ws_type.",
                        ".$ws_finished.",
                        ".$ws_approval.",
                        ".$issue_log.",
                        ".$issue_log_option.",
                        ".$original_bill_checked.",
                        ".$close_case_option.",
                        ".$original_bill_verified.",
                        ".$doc_send_back_to_client_required.",
                        '".$edited_by."',
                        '".$edit_date."',
                        7
                      );";
                  $r_insert = mysqli_query($con, $insert) or die(mysqli_error($con));
                  if ($r_insert) {
                    $case_id = mysqli_insert_id($con);
                    $update = mysqli_query($con, "UPDATE r_create_case SET r_create_case.`status` = 3, r_create_case.ref = '".$ref."', r_create_case.edit_date = NOW() WHERE r_create_case.id = ".$data[0].";") or die(mysqli_error($con));
                  }
              }
            }
          }
        }
        $Excel->getActiveSheet()->setCellValue('A'.$row, $data[0]);
        $Excel->getActiveSheet()->setCellValue('B'.$row, $status_1);
        $Excel->getActiveSheet()->setCellValue('C'.$row, $create_byx);
        $Excel->getActiveSheet()->setCellValue('D'.$row, $receive_datex);
        $Excel->getActiveSheet()->setCellValue('E'.$row, $categoryx);
        $Excel->getActiveSheet()->setCellValue('F'.$row, $typex);
        $Excel->getActiveSheet()->setCellValue('G'.$row, $admission_datex);
        $Excel->getActiveSheet()->setCellValue('H'.$row, $admission_timex);
        $Excel->getActiveSheet()->setCellValue('I'.$row, $discharge_datex);
        $Excel->getActiveSheet()->setCellValue('J'.$row, $discharge_doctorx);
        $Excel->getActiveSheet()->setCellValue('K'.$row, $diagnosisx);
        $Excel->getActiveSheet()->setCellValue('L'.$row, $bill_nox);
        $Excel->getActiveSheet()->setCellValue('M'.$row, $bill_issue_datex);
        $Excel->getActiveSheet()->setCellValue('N'.$row, $bill_due_datex);
        $Excel->getActiveSheet()->setCellValue('O'.$row, $providerx);
        $Excel->getActiveSheet()->setCellValue('P'.$row, $other_providerx);
        $Excel->getActiveSheet()->setCellValueExplicit('Q'.$row, $policy_nox,PHPExcel_Cell_DataType::TYPE_STRING);
        $Excel->getActiveSheet()->setCellValue('R'.$row, $dobx);
        $Excel->getActiveSheet()->setCellValue('S'.$row, $clientx);
        $Excel->getActiveSheet()->setCellValue('T'.$row, $patientx);
        $Excel->getActiveSheet()->setCellValue('U'.$row, $principlex);
        $Excel->getActiveSheet()->setCellValue('V'.$row, $case_id);
        $Excel->getActiveSheet()->setCellValue('W'.$row, $ref);
        $Excel->getActiveSheet()->setCellValue('X'.$row, $msg);
      }
    }
  }
}
$Excel->getActiveSheet()->setTitle('Result');
$objWriter = PHPExcel_IOFactory::createWriter($Excel, 'Excel2007');
$objWriter->save('Y:/ADMIN/EXTERNAL USE/UPLOAD_CASE/'.$file2);
echo '<script>alert("Please check result file");window.location.href="index2.html";</script>';
