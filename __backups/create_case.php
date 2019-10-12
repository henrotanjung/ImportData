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
$file='create_case_'.$date.'';
$inputFileName = 'Y:/ADMIN/EXTERNAL USE/UPLOAD_CASE/'.$file.'.xlsx';
$file2='create_case_result_'.$date.'.xlsx';


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
$Excel->getActiveSheet()->setCellValue('O1', 'ID Provider');
$Excel->getActiveSheet()->setCellValue('P1', 'Provider');
$Excel->getActiveSheet()->setCellValue('Q1', 'Other Provider');
$Excel->getActiveSheet()->setCellValue('R1', 'Other Provider City');
$Excel->getActiveSheet()->setCellValue('S1', 'Other Provider Country');
$Excel->getActiveSheet()->setCellValue('T1', 'Policy No');
$Excel->getActiveSheet()->setCellValue('U1', 'Dob');
$Excel->getActiveSheet()->setCellValue('V1', 'Client');
$Excel->getActiveSheet()->setCellValue('W1', 'Company');
$Excel->getActiveSheet()->setCellValue('X1', 'Branch');
$Excel->getActiveSheet()->setCellValue('Y1', 'Patient');
$Excel->getActiveSheet()->setCellValue('Z1', 'Patient Name');
$Excel->getActiveSheet()->setCellValue('AA1', 'Member Id');
$Excel->getActiveSheet()->setCellValue('AB1', 'Member Card');
$Excel->getActiveSheet()->setCellValue('AC1', 'Principle');
$Excel->getActiveSheet()->setCellValue('AD1', 'Gender');
$Excel->getActiveSheet()->setCellValue('AE1', 'Relation');
$Excel->getActiveSheet()->setCellValue('AF1', 'Policy Status');
$Excel->getActiveSheet()->setCellValue('AG1', 'Policy Holder');
$Excel->getActiveSheet()->setCellValue('AH1', 'Policy Effective Date');
$Excel->getActiveSheet()->setCellValue('AI1', 'Policy Expiry Date');
$Excel->getActiveSheet()->setCellValue('AJ1', 'Policy Issue Date');
$Excel->getActiveSheet()->setCellValue('AK1', 'Policy Declare Date');
$Excel->getActiveSheet()->setCellValue('AL1', 'Policy Lapse Date');
$Excel->getActiveSheet()->setCellValue('AM1', 'Policy Revival Date');
$Excel->getActiveSheet()->setCellValue('AN1', 'Policy Termination Date');
$Excel->getActiveSheet()->setCellValue('AO1', 'Policy Suspend Date');
$Excel->getActiveSheet()->setCellValue('AP1', 'Policy Unsuspend Date');
$Excel->getActiveSheet()->setCellValue('AQ1', 'Program');
$Excel->getActiveSheet()->setCellValue('AR1', 'Plan');
$Excel->getActiveSheet()->setCellValue('AS1', 'Plan Attach Date');
$Excel->getActiveSheet()->setCellValue('AT1', 'Plan Expiry Date');
$Excel->getActiveSheet()->setCellValue('AU1', 'Special Condition');
$Excel->getActiveSheet()->setCellValue('AV1', 'Exclusion');

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
      // IF ID EMPTY
      if (empty($data[0]) || is_null($data[0])) {
        // CEK MEMBER
        if (empty($data[24]) || is_null($data[24])) {
          if (!empty($data[19]) || !empty($data[21])) {
            $client = $data[21];
            $policy_no = $data[19];
            if (empty($data[20]) || is_null($data[20])) {
              $v_member = "SELECT member.id, member.member_dob FROM member WHERE member.client = ".$data[21]." AND member.policy_no = '".$data[19]."';";
              $r_member = mysqli_query($con, $v_member);
              $count = mysqli_num_rows($r_member);
              if ($count <> 1) {
                $patient = 'NULL';
                $patient_error_remarks = "Not Found";
              } else {
                while ($data1 = mysqli_fetch_row($r_member)) {
                  $patient = $data1[0];
                  $patient_error_remarks = NULL;
                  $dob = "'$data1[1]'";
                  $dobx = $data1[1];
                }
              }
            } else {
              $dob = "'$data[20]'";
              $dobx = $data[20];
              $v_member2 = "SELECT member.id FROM member WHERE member.client = ".$data[21]." AND member.policy_no = '".$data[19]."' AND member.member_dob = '".$data[20]."';";
              $r_member2 = mysqli_query($con, $v_member2);
              $count2 = mysqli_num_rows($r_member2);
              if ($count2 <> 1) {
                $patient = 'NULL';
                $patient_error_remarks = "Not Found";
              } else {
                while ($data2 = mysqli_fetch_row($r_member2)) {
                  $patient = $data2[0];
                  $patient_error_remarks = NULL;
                }
              }
            }
          } else {
            $patient = 'NULL';
            $client = NULL;
            $policy_no = NULL;
            $dob = 'NULL';
            $dobx = NULL;
            $patient_error_remarks = "Not Found";
          }
        } else {
          $v_member3 = "SELECT member.id, member.client, member.member_dob, member.policy_no from member WHERE member.id = ".$data[24].";";
          $r_member3 = mysqli_query($con, $v_member3);
          $count3 = mysqli_num_rows($r_member3) or die(mysqli_error($con));
          if ($count3 <> 1) {
            $patient = 'NULL';
            $patient_error_remarks = "Not Found";
            $dob = 'NULL';
            $dobx = NULL;
            $client = 'NULL';
            $policy_no = NULL;
          } else {
            while ($data5 = mysqli_fetch_row($r_member3)) {
              $patient = $data5[0];
              $dob = "'$data5[2]'";
              $dobx = $data5[2];
              $client = $data5[1];
              $policy_no = $data5[3];
              $patient_error_remarks = NULL;
            }
          }
        }

        if ($patient == 'NULL') {
          $patient_name = 'NULL';
          $member_id = 'NULL';
          $member_card = 'NULL';
          $principle = 'NULL';
          $gender = 'NULL';
          $relation = 'NULL';
          $policy_status = 'NULL';
          $policy_holder = 'NULL';
          $policy_effective_date = 'NULL';
          $policy_expiry_date = 'NULL';
          $policy_issue_date = 'NULL';
          $policy_declare_date = 'NULL';
          $policy_lapse_date = 'NULL';
          $policy_revival_date = 'NULL';
          $policy_termination_date = 'NULL';
          $policy_suspend_date = 'NULL';
          $policy_unsuspend_date = 'NULL';
          $program = 'NULL';
          $plan = 'NULL';
          $plan_attach_date = 'NULL';
          $plan_expiry_date = 'NULL';
          $special_condition = 'NULL';
          $exclusion = 'NULL';
          $company = 'NULL';
          $branch = 'NULL';
          $rider = 'NULL';
          $rider_attach_date = 'NULL';
          $rider_expiry_date = 'NULL';

          $patient_namex = NULL;
          $member_idx = NULL;
          $member_cardx = NULL;
          $principlex = NULL;
          $genderx = NULL;
          $relationx = NULL;
          $policy_statusx = NULL;
          $policy_holderx = NULL;
          $policy_effective_datex = NULL;
          $policy_expiry_datex = NULL;
          $policy_issue_datex = NULL;
          $policy_declare_datex = NULL;
          $policy_lapse_datex = NULL;
          $policy_revival_datex = NULL;
          $policy_termination_datex = NULL;
          $policy_suspend_datex = NULL;
          $policy_unsuspend_datex = NULL;
          $programx = NULL;
          $planx = NULL;
          $plan_attach_datex = NULL;
          $plan_expiry_datex = NULL;
          $special_conditionx = NULL;
          $exclusionx = NULL;
        } else {
          $member_data = mysqli_query($con, "SELECT
                                          member.id,
                                          member.member_name,
                                          member.member_id,
                                          member.member_card,
                                          member.member_principle,
                                          member.member_gender,
                                          member.member_relation,
                                          member.policy_status,
                                          member.policy_holder,
                                          member.policy_effective_date,
                                          member.policy_expiry_date,
                                          member.policy_issue_date,
                                          member.policy_declare_date,
                                          member.policy_lapse_date,
                                          member.policy_revival_date,
                                          member.policy_termination_date,
                                          member.policy_suspend_date,
                                          member.policy_unsuspend_date,
                                          member.program,
                                          member.plan,
                                          member.plan_attach_date,
                                          member.plan_expiry_date,
                                          member.special_condition,
                                          member.exclusion,
                                          principle.member_name AS principle_name,
                                          program.`name` AS program_name,
                                          plan.`name` AS plan_name,
                                          member.company,
                                          company.`name` as company_name,
                                          member.branch,
                                          branch.`name` as branch_name,
                                          member.rider,
                                          member.rider_attach_date,
                                          member.rider_expiry_date
                                        FROM
                                          member
                                        LEFT JOIN principle ON member.member_principle = principle.id
                                        LEFT JOIN program ON member.program = program.id
                                        LEFT JOIN plan ON member.plan = plan.id
                                        LEFT JOIN company ON member.company = company.id
                                        LEFT JOIN branch ON member.branch = branch.id
                                        WHERE
                                          member.id = ".$patient.";") or die(mysqli_error($con));
          while ($data3 = mysqli_fetch_row($member_data)) {
            $patient_name = $data3[1];
            $member_id = "'$data3[2]'";
            $member_card = "'$data3[3]'";
            $principle = $data3[4];
            $gender = $data3[5];
            $relation = $data3[6];
            $policy_status = $data3[7];
            $policy_holder = "'$data3[8]'";
            $policy_effective_date = "'$data3[9]'";
            $policy_expiry_date = "'$data3[10]'";
            $policy_issue_date = "'$data3[11]'";
            $policy_declare_date = "'$data3[12]'";          
            if (is_null($data3[13])) {
              $policy_lapse_date = 'NULL';
              $policy_lapse_datex = NULL;
              } else {
              $policy_lapse_date = "'$data3[13]'";
              $policy_lapse_datex = $data3[13];
            }        
            if (is_null($data3[14])) {
              $policy_revival_date = 'NULL';
              $policy_revival_datex = NULL;
            } else {
              $policy_revival_date = "'$data3[14]'";
              $policy_revival_datex = $data3[14];
            }          
            if (is_null($data3[15])) {
              $policy_termination_date = 'NULL';
              $policy_termination_datex = NULL;
            } else {
              $policy_termination_date = "'$data3[15]'";
              $policy_termination_datex = $data3[15];
            }
            if (is_null($data3[16])) {
              $policy_suspend_date = 'NULL';
              $policy_suspend_datex = NULL;
            } else {
              $policy_suspend_date = "'$data3[16]'";
              $policy_suspend_datex = $data3[16];
            }
            if (is_null($data3[17])) {
              $policy_unsuspend_date = 'NULL';
              $policy_unsuspend_datex = NULL;
            } else {
              $policy_unsuspend_date = "'$data3[17]'";
              $policy_unsuspend_datex = $data3[17];
            }
            $program = $data3[18];
            $plan = $data3[19];
            $plan_attach_date = "'$data3[20]'";
            $plan_expiry_date = "'$data3[21]'";
            $special_condition = "'$data3[22]'";
            $exclusion = "'$data3[23]'";
            if (is_null($data3[27])) {
              $company = 'NULL';
            } else {
              $company = $data3[27];
            }
            if (is_null($data3[29])) {
              $branch = 'NULL';
            } else {
              $branch = $data3[29];
            }
            if (is_null($data3[31])) {
              $rider = 'NULL';
            } else {
              $rider = $data3[31];
            }
            if (is_null($data3[32])) {
              $rider_attach_date = 'NULL';
            } else {
              $rider_attach_date = "'$data3[32]'";
            }
            if (is_null($data3[33])) {
              $rider_expiry_date = 'NULL';
            } else {
              $rider_expiry_date = "'$data3[33]'";
            }

            $patient_namex = $data3[1];
            $member_idx = $data3[2];
            $member_cardx = $data3[3];
            $principlex = $data3[24];
            $genderx = $data3[5];
            $relationx = $data3[6];
            $policy_statusx = $data3[7];
            $policy_holderx = $data3[8];
            $policy_effective_datex = $data3[9];
            $policy_expiry_datex = $data3[10];
            $policy_declare_datex = $data3[12];
            $programx = $data3[25];
            $policy_issue_datex = $data3[11];
            $planx = $data3[26];
            $plan_attach_datex = $data3[20];
            $plan_expiry_datex = $data3[21];
            $special_conditionx = $data3[22];
            $exclusionx = $data3[23]; 
            $companyx = $data3[28];
            $branchx = $data3[30];
          }
        }
        // CEK PROVIDER
        if (empty($data[14]) || is_null($data[14])) {
          if (!empty($data[15])) {
            $verify_rs = "SELECT provider.id, provider.full_name FROM provider WHERE provider.full_name LIKE '%".$data[15]."%';";
            $r_verify_rs = mysqli_query($con, $verify_rs) or die(mysqli_error($con));
            $count = mysqli_num_rows($r_verify_rs);
            if ($count > 0) {
              while ($data4 = mysqli_fetch_row($r_verify_rs)) {
                $id_provider = $data4[0];
                $provider_name = $data4[1];
                $provider_error_remarks = $data[15];
              }
            } else {
              $id_provider = 310;
              $provider_name = "NON PARTICIPATING HOSPITAL";
              $provider_error_remarks = "Not Found";
            }
          } else {
            $id_provider = 'NULL';
            $provider_name = $data[15];
            $provider_error_remarks = "Not Found";
          }
        } else {
          $verify_rs2 = "SELECT provider.id, provider.full_name FROM provider WHERE provider.id = ".$data[14].";";
          $r_verify_rs2 = mysqli_query($con, $verify_rs2) or die(mysqli_error($con));
          $count = mysqli_num_rows($r_verify_rs2);
          if ($count > 0) {
            while ($data6 = mysqli_fetch_row($r_verify_rs2)) {
              $id_provider = $data6[0];
              $provider_name = $data6[1];
              $provider_error_remarks = NULL;
            }
          } else {
            $id_provider = $data[14];
            $provider_name = 'NULL';
            $provider_error_remarks = "Not Found";
          }
        }

        if ($patient == 'NULL' || $id_provider == 'NULL') {
          $status = 2;
        } else {
          $status = 1;
        }

        $create_by = $data[2];
        $receive_date = $data[3];
        $category = $data[4];
        $type = $data[5];
        $admission_date = $data[6];
        $admission_time = $data[7];
        $discharge_date = $data[8];
        $discharge_doctor = $data[9];
        $diagnosis = $data[10];
        $bill_no = $data[11];
        $bill_issue_date = date('Y-m-d');
        $bill_due_date = date('Y-m-d', strtotime("+7 days"));
        //$provider = $data[15];
        $other_provider = $data[16];
        $other_provider_city = $data[17];
        $other_provider_country = $data[18];
        //$policy_no = $data[19];
        //$client = $data[21];

        $insert = "INSERT INTO `r_create_case` (
                            `status`,
                            `created_by`,
                            `receive_date`,
                            `category`,
                            `type`,
                            `admission_date`,
                            `admission_time`,
                            `discharge_date`,
                            `discharge_doctor`,
                            `diagnosis`,
                            `bill_no`,
                            `bill_issue_date`,
                            `bill_due_date`,
                            `id_provider`,
                            `provider_error_remarks`,
                            `provider`,
                            `other_provider`,
                            `other_provider_city`,
                            `other_provider_country`,
                            `policy_no`,
                            `dob`,
                            `client`,
                            `company`,
                            `branch`,
                            `patient_error_remarks`,
                            `patient`,
                            `member_id`,
                            `member_card`,
                            `principle`,
                            `gender`,
                            `relation`,
                            `policy_status`,
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
                            `program`,
                            `plan`,
                            `plan_attach_date`,
                            `plan_expiry_date`,
                            `rider`,
                            `rider_attach_date`,
                            `rider_expiry_date`,
                            `special_condition`,
                            `exclusion`,
                            `create_date`,
                            `edit_date`
                          )
                          VALUES
                            (
                              ".$status.",
                              '".$create_by."',
                              '".$receive_date."',
                              ".$category.",
                              ".$type.",
                              '".$admission_date."',
                              '".$admission_time."',
                              '".$discharge_date."',
                              '".$discharge_doctor."',
                              '".$diagnosis."',
                              '".$bill_no."',
                              '".$bill_issue_date."',
                              '".$bill_due_date."',
                              ".$id_provider.",
                              '".$provider_error_remarks."',
                              '".$provider_name."',
                              '".$other_provider."',
                              '".$other_provider_city."',
                              ".$other_provider_country.",
                              '".$policy_no."',
                              ".$dob.",
                              ".$client.",
                              ".$company.",
                              ".$branch.",
                              '".$patient_error_remarks."',
                              ".$patient.",
                              ".$member_id.",
                              ".$member_card.",
                              ".$principle.",
                              ".$gender.",
                              ".$relation.",
                              ".$policy_status.",
                              ".$policy_holder.",
                              ".$policy_effective_date.",
                              ".$policy_expiry_date.",
                              ".$policy_issue_date.",
                              ".$policy_declare_date.",
                              ".$policy_lapse_date.",
                              ".$policy_revival_date.",
                              ".$policy_termination_date.",
                              ".$policy_suspend_date.",
                              ".$policy_unsuspend_date.",
                              ".$program.",
                              ".$plan.",
                              ".$plan_attach_date.",
                              ".$plan_expiry_date.",
                              ".$rider.",
                              ".$rider_attach_date.",
                              ".$rider_expiry_date.",
                              ".$special_condition.",
                              ".$exclusion.",
                              NOW(),
                              NULL
                            );";
      $result = mysqli_query($con, $insert) or die(mysqli_error($con));
      if ($result) {
        $id = mysqli_insert_id($con);
        $Excel->getActiveSheet()->setCellValue('A'.$row, $id);
        $Excel->getActiveSheet()->setCellValue('B'.$row, $status);
        $Excel->getActiveSheet()->setCellValue('C'.$row, $create_by);
        $Excel->getActiveSheet()->setCellValue('D'.$row, $receive_date);
        $Excel->getActiveSheet()->setCellValue('E'.$row, $category);
        $Excel->getActiveSheet()->setCellValue('F'.$row, $type);    
        $Excel->getActiveSheet()->setCellValue('G'.$row, $admission_date);
        $Excel->getActiveSheet()->setCellValue('H'.$row, $admission_time);
        $Excel->getActiveSheet()->setCellValue('I'.$row, $discharge_date);
        $Excel->getActiveSheet()->setCellValue('J'.$row, $discharge_doctor);
        $Excel->getActiveSheet()->setCellValue('K'.$row, $diagnosis);
        $Excel->getActiveSheet()->setCellValue('L'.$row, $bill_no);
        $Excel->getActiveSheet()->setCellValue('M'.$row, $bill_issue_date);
        $Excel->getActiveSheet()->setCellValue('N'.$row, $bill_due_date);
        if ($id_provider == 'NULL') {
          $Excel->getActiveSheet()->setCellValue('O'.$row, $provider_error_remarks);
        } else {
          $Excel->getActiveSheet()->setCellValue('O'.$row, $id_provider);
        }
        $Excel->getActiveSheet()->setCellValue('P'.$row, $provider_name);
        $Excel->getActiveSheet()->setCellValue('Q'.$row, $other_provider);
        $Excel->getActiveSheet()->setCellValue('R'.$row, $other_provider_city);
        $Excel->getActiveSheet()->setCellValue('S'.$row, $other_provider_country);
        $Excel->getActiveSheet()->setCellValueExplicit('T'.$row, $policy_no,PHPExcel_Cell_DataType::TYPE_STRING);
        $Excel->getActiveSheet()->setCellValue('U'.$row, $dobx);
        $Excel->getActiveSheet()->setCellValue('V'.$row, $client);
        $Excel->getActiveSheet()->setCellValue('W'.$row, $companyx);
        $Excel->getActiveSheet()->setCellValue('X'.$row, $branchx);
        if ($patient == 'NULL') {
          $Excel->getActiveSheet()->setCellValue('Y'.$row, $patient_error_remarks);
        } else {
          $Excel->getActiveSheet()->setCellValue('Y'.$row, $patient);      
        }
        $Excel->getActiveSheet()->setCellValue('Z'.$row, $patient_namex);
        $Excel->getActiveSheet()->setCellValueExplicit('AA'.$row, $member_idx,PHPExcel_Cell_DataType::TYPE_STRING);
        $Excel->getActiveSheet()->setCellValueExplicit('AB'.$row, $member_cardx,PHPExcel_Cell_DataType::TYPE_STRING);
        $Excel->getActiveSheet()->setCellValue('AC'.$row, $principlex);
        $Excel->getActiveSheet()->setCellValue('AD'.$row, $genderx);
        $Excel->getActiveSheet()->setCellValue('AE'.$row, $relationx);
        $Excel->getActiveSheet()->setCellValue('AF'.$row, $policy_statusx);
        $Excel->getActiveSheet()->setCellValue('AG'.$row, $policy_holderx);
        $Excel->getActiveSheet()->setCellValue('AH'.$row, $policy_effective_datex);
        $Excel->getActiveSheet()->setCellValue('AI'.$row, $policy_expiry_datex);
        $Excel->getActiveSheet()->setCellValue('AJ'.$row, $policy_issue_datex);
        $Excel->getActiveSheet()->setCellValue('AK'.$row, $policy_declare_datex);
        $Excel->getActiveSheet()->setCellValue('AL'.$row, $policy_lapse_datex);
        $Excel->getActiveSheet()->setCellValue('AM'.$row, $policy_revival_datex);
        $Excel->getActiveSheet()->setCellValue('AN'.$row, $policy_termination_datex);
        $Excel->getActiveSheet()->setCellValue('AO'.$row, $policy_suspend_datex);
        $Excel->getActiveSheet()->setCellValue('AP'.$row, $policy_unsuspend_datex);
        $Excel->getActiveSheet()->setCellValue('AQ'.$row, $programx);
        $Excel->getActiveSheet()->setCellValue('AR'.$row, $planx);
        $Excel->getActiveSheet()->setCellValue('AS'.$row, $plan_attach_datex);
        $Excel->getActiveSheet()->setCellValue('AT'.$row, $plan_expiry_datex);
        $Excel->getActiveSheet()->setCellValue('AU'.$row, $special_conditionx);
        $Excel->getActiveSheet()->setCellValue('AV'.$row, $exclusionx);                                                                                                                                                                                                                                                                                 
        }
        // END

      } else {
        $cek_id = mysqli_query($con, "SELECT r_create_case.id FROM r_create_case WHERE r_create_case.id = ".$data[0].";") or die(mysqli_error($con));
        $count = mysqli_num_rows($cek_id);
        if ($count != 1) {
          // IF ID NOT EXISTS
            // CEK MEMBER
          if (empty($data[24]) || is_null($data[24])) {
            if (!empty($data[19]) || !empty($data[21])) {
              $client = $data[21];
              $policy_no = $data[19];
              if (empty($data[20]) || is_null($data[20])) {
                $v_member = "SELECT member.id, member.member_dob FROM member WHERE member.client = ".$data[21]." AND member.policy_no = '".$data[19]."';";
                $r_member = mysqli_query($con, $v_member);
                $count = mysqli_num_rows($r_member);
                if ($count <> 1) {
                  $patient = 'NULL';
                  $patient_error_remarks = "Not Found";
                } else {
                  while ($data1 = mysqli_fetch_row($r_member)) {
                    $patient = $data1[0];
                    $patient_error_remarks = NULL;
                    $dob = "'$data1[1]'";
                    $dobx = $data1[1];
                  }
                }
              } else {
                $dob = "'$data[20]'";
                $dobx = $data[20];
                $v_member2 = "SELECT member.id FROM member WHERE member.client = ".$data[21]." AND member.policy_no = '".$data[19]."' AND member.member_dob = '".$data[20]."';";
                $r_member2 = mysqli_query($con, $v_member2);
                $count2 = mysqli_num_rows($r_member2);
                if ($count2 <> 1) {
                  $patient = 'NULL';
                  $patient_error_remarks = "Not Found";
                } else {
                  while ($data2 = mysqli_fetch_row($r_member2)) {
                    $patient = $data2[0];
                    $patient_error_remarks = NULL;
                  }
                }
              }
            } else {
              $patient = 'NULL';
              $client = NULL;
              $policy_no = NULL;
              $dob = 'NULL';
              $dobx = NULL;
              $patient_error_remarks = "Not Found";
            }
          } else {
            $v_member3 = "SELECT member.id, member.client, member.member_dob, member.policy_no from member WHERE member.id = ".$data[24].";";
            $r_member3 = mysqli_query($con, $v_member3);
            $count3 = mysqli_num_rows($r_member3) or die(mysqli_error($con));
            if ($count3 <> 1) {
              $patient = 'NULL';
              $patient_error_remarks = "Not Found";
              $dob = 'NULL';
              $dobx = NULL;
              $client = 'NULL';
              $policy_no = NULL;
            } else {
              while ($data5 = mysqli_fetch_row($r_member3)) {
                $patient = $data5[0];
                $dob = "'$data5[2]'";
                $dobx = $data5[2];
                $client = $data5[1];
                $policy_no = $data5[3];
                $patient_error_remarks = NULL;
              }
            }
          }

          if ($patient == 'NULL') {
            $patient_name = 'NULL';
            $member_id = 'NULL';
            $member_card = 'NULL';
            $principle = 'NULL';
            $gender = 'NULL';
            $relation = 'NULL';
            $policy_status = 'NULL';
            $policy_holder = 'NULL';
            $policy_effective_date = 'NULL';
            $policy_expiry_date = 'NULL';
            $policy_issue_date = 'NULL';
            $policy_declare_date = 'NULL';
            $policy_lapse_date = 'NULL';
            $policy_revival_date = 'NULL';
            $policy_termination_date = 'NULL';
            $policy_suspend_date = 'NULL';
            $policy_unsuspend_date = 'NULL';
            $program = 'NULL';
            $plan = 'NULL';
            $plan_attach_date = 'NULL';
            $plan_expiry_date = 'NULL';
            $special_condition = 'NULL';
            $exclusion = 'NULL';
            $company = 'NULL';
            $branch = 'NULL';
            $rider = 'NULL';
            $rider_attach_date = 'NULL';
            $rider_expiry_date = 'NULL';

            $patient_namex = NULL;
            $member_idx = NULL;
            $member_cardx = NULL;
            $principlex = NULL;
            $genderx = NULL;
            $relationx = NULL;
            $policy_statusx = NULL;
            $policy_holderx = NULL;
            $policy_effective_datex = NULL;
            $policy_expiry_datex = NULL;
            $policy_issue_datex = NULL;
            $policy_declare_datex = NULL;
            $policy_lapse_datex = NULL;
            $policy_revival_datex = NULL;
            $policy_termination_datex = NULL;
            $policy_suspend_datex = NULL;
            $policy_unsuspend_datex = NULL;
            $programx = NULL;
            $planx = NULL;
            $plan_attach_datex = NULL;
            $plan_expiry_datex = NULL;
            $special_conditionx = NULL;
            $exclusionx = NULL;
          } else {
            $member_data = mysqli_query($con, "SELECT
                                          member.id,
                                          member.member_name,
                                          member.member_id,
                                          member.member_card,
                                          member.member_principle,
                                          member.member_gender,
                                          member.member_relation,
                                          member.policy_status,
                                          member.policy_holder,
                                          member.policy_effective_date,
                                          member.policy_expiry_date,
                                          member.policy_issue_date,
                                          member.policy_declare_date,
                                          member.policy_lapse_date,
                                          member.policy_revival_date,
                                          member.policy_termination_date,
                                          member.policy_suspend_date,
                                          member.policy_unsuspend_date,
                                          member.program,
                                          member.plan,
                                          member.plan_attach_date,
                                          member.plan_expiry_date,
                                          member.special_condition,
                                          member.exclusion,
                                          principle.member_name AS principle_name,
                                          program.`name` AS program_name,
                                          plan.`name` AS plan_name,
                                          member.company,
                                          company.`name` as company_name,
                                          member.branch,
                                          branch.`name` as branch_name,
                                          member.rider,
                                          member.rider_attach_date,
                                          member.rider_expiry_date
                                        FROM
                                          member
                                        LEFT JOIN principle ON member.member_principle = principle.id
                                        LEFT JOIN program ON member.program = program.id
                                        LEFT JOIN plan ON member.plan = plan.id
                                        LEFT JOIN company ON member.company = company.id
                                        LEFT JOIN branch ON member.branch = branch.id
                                        WHERE
                                      member.id = ".$patient.";") or die(mysqli_error($con));
            while ($data3 = mysqli_fetch_row($member_data)) {
              $patient_name = $data3[1];
              $member_id = "'$data3[2]'";
              $member_card = "'$data3[3]'";
              $principle = $data3[4];
              $gender = $data3[5];
              $relation = $data3[6];
              $policy_status = $data3[7];
              $policy_holder = "'$data3[8]'";
              $policy_effective_date = "'$data3[9]'";
              $policy_expiry_date = "'$data3[10]'";
              $policy_issue_date = "'$data3[11]'";
              $policy_declare_date = "'$data3[12]'";          
              if (is_null($data3[13])) {
                $policy_lapse_date = 'NULL';
                $policy_lapse_datex = NULL;
                } else {
                $policy_lapse_date = "'$data3[13]'";
                $policy_lapse_datex = $data3[13];
              }        
              if (is_null($data3[14])) {
                $policy_revival_date = 'NULL';
                $policy_revival_datex = NULL;
              } else {
                $policy_revival_date = "'$data3[14]'";
                $policy_revival_datex = $data3[14];
              }          
              if (is_null($data3[15])) {
                $policy_termination_date = 'NULL';
                $policy_termination_datex = NULL;
              } else {
                $policy_termination_date = "'$data3[15]'";
                $policy_termination_datex = $data3[15];
              }
              if (is_null($data3[16])) {
                $policy_suspend_date = 'NULL';
                $policy_suspend_datex = NULL;
              } else {
                $policy_suspend_date = "'$data3[16]'";
                $policy_suspend_datex = $data3[16];
              }
              if (is_null($data3[17])) {
                $policy_unsuspend_date = 'NULL';
                $policy_unsuspend_datex = NULL;
              } else {
                $policy_unsuspend_date = "'$data3[17]'";
                $policy_unsuspend_datex = $data3[17];
              }
              $program = $data3[18];
              $plan = $data3[19];
              $plan_attach_date = "'$data3[20]'";
              $plan_expiry_date = "'$data3[21]'";
              $special_condition = "'$data3[22]'";
              $exclusion = "'$data3[23]'";
              if (is_null($data3[27])) {
                $company = 'NULL';
              } else {
                $company = $data3[27];
              }
              if (is_null($data3[29])) {
                $branch = 'NULL';
              } else {
                $branch = $data3[29];
              }
              if (is_null($data3[31])) {
                $rider = 'NULL';
              } else {
                $rider = $data3[31];
              }
              if (is_null($data3[32])) {
                $rider_attach_date = 'NULL';
              } else {
                $rider_attach_date = "'$data3[32]'";
              }
              if (is_null($data3[33])) {
                $rider_expiry_date = 'NULL';
              } else {
                $rider_expiry_date = "'$data3[33]'";
              }

              $patient_namex = $data3[1];
              $member_idx = $data3[2];
              $member_cardx = $data3[3];
              $principlex = $data3[24];
              $genderx = $data3[5];
              $relationx = $data3[6];
              $policy_statusx = $data3[7];
              $policy_holderx = $data3[8];
              $policy_effective_datex = $data3[9];
              $policy_expiry_datex = $data3[10];
              $policy_declare_datex = $data3[12];
              $programx = $data3[25];
              $policy_issue_datex = $data3[11];
              $planx = $data3[26];
              $plan_attach_datex = $data3[20];
              $plan_expiry_datex = $data3[21];
              $special_conditionx = $data3[22];
              $exclusionx = $data3[23]; 
              $companyx = $data3[28];
              $branchx = $data3[30];
            }
          }
          // CEK PROVIDER
          if (empty($data[14]) || is_null($data[14])) {
            if (!empty($data[15])) {
              $verify_rs = "SELECT provider.id, provider.full_name FROM provider WHERE provider.full_name LIKE '%".$data[15]."%';";
              $r_verify_rs = mysqli_query($con, $verify_rs) or die(mysqli_error($con));
              $count = mysqli_num_rows($r_verify_rs);
              if ($count > 0) {
                while ($data4 = mysqli_fetch_row($r_verify_rs)) {
                  $id_provider = $data4[0];
                  $provider_name = $data4[1];
                  $provider_error_remarks = $data[15];
                }
              } else {
                $id_provider = 310;
                $provider_name = "NON PARTICIPATING HOSPITAL";
                $provider_error_remarks = "Not Found";
              }
            } else {
              $id_provider = 'NULL';
              $provider_name = $data[15];
              $provider_error_remarks = "Not Found";
            }
          } else {
            $verify_rs2 = "SELECT provider.id, provider.full_name FROM provider WHERE provider.id = ".$data[14].";";
            $r_verify_rs2 = mysqli_query($con, $verify_rs2) or die(mysqli_error($con));
            $count = mysqli_num_rows($r_verify_rs2);
            if ($count > 0) {
              while ($data6 = mysqli_fetch_row($r_verify_rs2)) {
                $id_provider = $data6[0];
                $provider_name = $data6[1];
                $provider_error_remarks = NULL;
              }
            } else {
              $id_provider = $data[14];
              $provider_name = 'NULL';
              $provider_error_remarks = "Not Found";
            }
          }

          if ($patient == 'NULL' || $id_provider == 'NULL') {
            $status = 2;
          } else {
            $status = 1;
          }

          $create_by = $data[2];
          $receive_date = $data[3];
          $category = $data[4];
          $type = $data[5];
          $admission_date = $data[6];
          $admission_time = $data[7];
          $discharge_date = $data[8];
          $discharge_doctor = $data[9];
          $diagnosis = $data[10];
          $bill_no = $data[11];
          $bill_issue_date = date('Y-m-d');
          $bill_due_date = date('Y-m-d', strtotime("+7 days"));
          $provider = $data[15];
          $other_provider = $data[16];
          $other_provider_city = $data[17];
          $other_provider_country = $data[18];
          //$policy_no = $data[19];
          //$client = $data[21];

          $insert = "INSERT INTO `r_create_case` (
                            `status`,
                            `created_by`,
                            `receive_date`,
                            `category`,
                            `type`,
                            `admission_date`,
                            `admission_time`,
                            `discharge_date`,
                            `discharge_doctor`,
                            `diagnosis`,
                            `bill_no`,
                            `bill_issue_date`,
                            `bill_due_date`,
                            `id_provider`,
                            `provider_error_remarks`,
                            `provider`,
                            `other_provider`,
                            `other_provider_city`,
                            `other_provider_country`,
                            `policy_no`,
                            `dob`,
                            `client`,
                            `company`,
                            `branch`,
                            `patient_error_remarks`,
                            `patient`,
                            `member_id`,
                            `member_card`,
                            `principle`,
                            `gender`,
                            `relation`,
                            `policy_status`,
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
                            `program`,
                            `plan`,
                            `plan_attach_date`,
                            `plan_expiry_date`,
                            `rider`,
                            `rider_attach_date`,
                            `rider_expiry_date`,
                            `special_condition`,
                            `exclusion`,
                            `create_date`,
                            `edit_date`
                          )
                          VALUES
                            (
                              ".$status.",
                              '".$create_by."',
                              '".$receive_date."',
                              ".$category.",
                              ".$type.",
                              '".$admission_date."',
                              '".$admission_time."',
                              '".$discharge_date."',
                              '".$discharge_doctor."',
                              '".$diagnosis."',
                              '".$bill_no."',
                              '".$bill_issue_date."',
                              '".$bill_due_date."',
                              ".$id_provider.",
                              '".$provider_error_remarks."',
                              '".$provider_name."',
                              '".$other_provider."',
                              '".$other_provider_city."',
                              ".$other_provider_country.",
                              '".$policy_no."',
                              ".$dob.",
                              ".$client.",
                              ".$company.",
                              ".$branch.",
                              '".$patient_error_remarks."',
                              ".$patient.",
                              ".$member_id.",
                              ".$member_card.",
                              ".$principle.",
                              ".$gender.",
                              ".$relation.",
                              ".$policy_status.",
                              ".$policy_holder.",
                              ".$policy_effective_date.",
                              ".$policy_expiry_date.",
                              ".$policy_issue_date.",
                              ".$policy_declare_date.",
                              ".$policy_lapse_date.",
                              ".$policy_revival_date.",
                              ".$policy_termination_date.",
                              ".$policy_suspend_date.",
                              ".$policy_unsuspend_date.",
                              ".$program.",
                              ".$plan.",
                              ".$plan_attach_date.",
                              ".$plan_expiry_date.",
                              ".$rider.",
                              ".$rider_attach_date.",
                              ".$rider_expiry_date.",
                              ".$special_condition.",
                              ".$exclusion.",
                              NOW(),
                              NULL
                              );";
        $result = mysqli_query($con, $insert) or die(mysqli_error($con));
        if ($result) {
          $id = mysqli_insert_id($con);
          $Excel->getActiveSheet()->setCellValue('A'.$row, $id);
          $Excel->getActiveSheet()->setCellValue('B'.$row, $status);
          $Excel->getActiveSheet()->setCellValue('C'.$row, $create_by);
          $Excel->getActiveSheet()->setCellValue('D'.$row, $receive_date);
          $Excel->getActiveSheet()->setCellValue('E'.$row, $category);
          $Excel->getActiveSheet()->setCellValue('F'.$row, $type);    
          $Excel->getActiveSheet()->setCellValue('G'.$row, $admission_date);
          $Excel->getActiveSheet()->setCellValue('H'.$row, $admission_time);
          $Excel->getActiveSheet()->setCellValue('I'.$row, $discharge_date);
          $Excel->getActiveSheet()->setCellValue('J'.$row, $discharge_doctor);
          $Excel->getActiveSheet()->setCellValue('K'.$row, $diagnosis);
          $Excel->getActiveSheet()->setCellValue('L'.$row, $bill_no);
          $Excel->getActiveSheet()->setCellValue('M'.$row, $bill_issue_date);
          $Excel->getActiveSheet()->setCellValue('N'.$row, $bill_due_date);
          if ($id_provider == 'NULL') {
            $Excel->getActiveSheet()->setCellValue('O'.$row, $provider_error_remarks);
          } else {
            $Excel->getActiveSheet()->setCellValue('O'.$row, $id_provider);
          }
          $Excel->getActiveSheet()->setCellValue('P'.$row, $provider_name);
          $Excel->getActiveSheet()->setCellValue('Q'.$row, $other_provider);
          $Excel->getActiveSheet()->setCellValue('R'.$row, $other_provider_city);
          $Excel->getActiveSheet()->setCellValue('S'.$row, $other_provider_country);
          $Excel->getActiveSheet()->setCellValueExplicit('T'.$row, $policy_no,PHPExcel_Cell_DataType::TYPE_STRING);
          $Excel->getActiveSheet()->setCellValue('U'.$row, $dobx);
          $Excel->getActiveSheet()->setCellValue('V'.$row, $client);
          $Excel->getActiveSheet()->setCellValue('W'.$row, $companyx);
          $Excel->getActiveSheet()->setCellValue('X'.$row, $branchx);
          if ($patient == 'NULL') {
            $Excel->getActiveSheet()->setCellValue('Y'.$row, $patient_error_remarks);
          } else {
            $Excel->getActiveSheet()->setCellValue('Y'.$row, $patient);      
          }
          $Excel->getActiveSheet()->setCellValue('Z'.$row, $patient_namex);
          $Excel->getActiveSheet()->setCellValueExplicit('AA'.$row, $member_idx,PHPExcel_Cell_DataType::TYPE_STRING);
          $Excel->getActiveSheet()->setCellValueExplicit('AB'.$row, $member_cardx,PHPExcel_Cell_DataType::TYPE_STRING);
          $Excel->getActiveSheet()->setCellValue('AC'.$row, $principlex);
          $Excel->getActiveSheet()->setCellValue('AD'.$row, $genderx);
          $Excel->getActiveSheet()->setCellValue('AE'.$row, $relationx);
          $Excel->getActiveSheet()->setCellValue('AF'.$row, $policy_statusx);
          $Excel->getActiveSheet()->setCellValue('AG'.$row, $policy_holderx);
          $Excel->getActiveSheet()->setCellValue('AH'.$row, $policy_effective_datex);
          $Excel->getActiveSheet()->setCellValue('AI'.$row, $policy_expiry_datex);
          $Excel->getActiveSheet()->setCellValue('AJ'.$row, $policy_issue_datex);
          $Excel->getActiveSheet()->setCellValue('AK'.$row, $policy_declare_datex);
          $Excel->getActiveSheet()->setCellValue('AL'.$row, $policy_lapse_datex);
          $Excel->getActiveSheet()->setCellValue('AM'.$row, $policy_revival_datex);
          $Excel->getActiveSheet()->setCellValue('AN'.$row, $policy_termination_datex);
          $Excel->getActiveSheet()->setCellValue('AO'.$row, $policy_suspend_datex);
          $Excel->getActiveSheet()->setCellValue('AP'.$row, $policy_unsuspend_datex);
          $Excel->getActiveSheet()->setCellValue('AQ'.$row, $programx);
          $Excel->getActiveSheet()->setCellValue('AR'.$row, $planx);
          $Excel->getActiveSheet()->setCellValue('AS'.$row, $plan_attach_datex);
          $Excel->getActiveSheet()->setCellValue('AT'.$row, $plan_expiry_datex);
          $Excel->getActiveSheet()->setCellValue('AU'.$row, $special_conditionx);
          $Excel->getActiveSheet()->setCellValue('AV'.$row, $exclusionx);                                                                                                                                                                                                                                                                                 
          }
          // END

        } else {
          // IF ID EXISTS
            // CEK MEMBER
          if (empty($data[24]) || is_null($data[24])) {
            if (!empty($data[19]) || !empty($data[21])) {
              $client = $data[21];
              $policy_no = $data[19];
              if (empty($data[20]) || is_null($data[20])) {
                $v_member = "SELECT member.id, member.member_dob FROM member WHERE member.client = ".$data[21]." AND member.policy_no = '".$data[19]."';";
                $r_member = mysqli_query($con, $v_member);
                $count = mysqli_num_rows($r_member);
                if ($count <> 1) {
                  $patient = 'NULL';
                  $patient_error_remarks = "Not Found";
                } else {
                  while ($data1 = mysqli_fetch_row($r_member)) {
                    $patient = $data1[0];
                    $patient_error_remarks = NULL;
                    $dob = "'$data1[1]'";
                    $dobx = $data1[1];
                  }
                }
              } else {
                $dob = "'$data[20]'";
                $dobx = $data[20];
                $v_member2 = "SELECT member.id FROM member WHERE member.client = ".$data[21]." AND member.policy_no = '".$data[19]."' AND member.member_dob = '".$data[20]."';";
                $r_member2 = mysqli_query($con, $v_member2);
                $count2 = mysqli_num_rows($r_member2);
                if ($count2 <> 1) {
                  $patient = 'NULL';
                  $patient_error_remarks = "Not Found";
                } else {
                  while ($data2 = mysqli_fetch_row($r_member2)) {
                    $patient = $data2[0];
                    $patient_error_remarks = NULL;
                  }
                }
              }
            } else {
              $patient = 'NULL';
              $client = NULL;
              $policy_no = NULL;
              $dob = 'NULL';
              $dobx = NULL;
              $patient_error_remarks = "Not Found";
            }
          } else {
            $v_member3 = "SELECT member.id, member.client, member.member_dob, member.policy_no from member WHERE member.id = ".$data[24].";";
            $r_member3 = mysqli_query($con, $v_member3);
            $count3 = mysqli_num_rows($r_member3) or die(mysqli_error($con));
            if ($count3 <> 1) {
              $patient = 'NULL';
              $patient_error_remarks = "Not Found";
              $dob = 'NULL';
              $dobx = NULL;
              $client = 'NULL';
              $policy_no = NULL;
            } else {
              while ($data5 = mysqli_fetch_row($r_member3)) {
                $patient = $data5[0];
                $dob = "'$data5[2]'";
                $dobx = $data5[2];
                $client = $data5[1];
                $policy_no = $data5[3];
                $patient_error_remarks = NULL;
              }
            }
          }

          if ($patient == 'NULL') {
            $patient_name = 'NULL';
            $member_id = 'NULL';
            $member_card = 'NULL';
            $principle = 'NULL';
            $gender = 'NULL';
            $relation = 'NULL';
            $policy_status = 'NULL';
            $policy_holder = 'NULL';
            $policy_effective_date = 'NULL';
            $policy_expiry_date = 'NULL';
            $policy_issue_date = 'NULL';
            $policy_declare_date = 'NULL';
            $policy_lapse_date = 'NULL';
            $policy_revival_date = 'NULL';
            $policy_termination_date = 'NULL';
            $policy_suspend_date = 'NULL';
            $policy_unsuspend_date = 'NULL';
            $program = 'NULL';
            $plan = 'NULL';
            $plan_attach_date = 'NULL';
            $plan_expiry_date = 'NULL';
            $special_condition = 'NULL';
            $exclusion = 'NULL';
            $company = 'NULL';
            $branch = 'NULL';
            $rider = 'NULL';
            $rider_attach_date = 'NULL';
            $rider_expiry_date = 'NULL';

            $patient_namex = NULL;
            $member_idx = NULL;
            $member_cardx = NULL;
            $principlex = NULL;
            $genderx = NULL;
            $relationx = NULL;
            $policy_statusx = NULL;
            $policy_holderx = NULL;
            $policy_effective_datex = NULL;
            $policy_expiry_datex = NULL;
            $policy_issue_datex = NULL;
            $policy_declare_datex = NULL;
            $policy_lapse_datex = NULL;
            $policy_revival_datex = NULL;
            $policy_termination_datex = NULL;
            $policy_suspend_datex = NULL;
            $policy_unsuspend_datex = NULL;
            $programx = NULL;
            $planx = NULL;
            $plan_attach_datex = NULL;
            $plan_expiry_datex = NULL;
            $special_conditionx = NULL;
            $exclusionx = NULL;
          } else {
            $member_data = mysqli_query($con, "SELECT
                                          member.id,
                                          member.member_name,
                                          member.member_id,
                                          member.member_card,
                                          member.member_principle,
                                          member.member_gender,
                                          member.member_relation,
                                          member.policy_status,
                                          member.policy_holder,
                                          member.policy_effective_date,
                                          member.policy_expiry_date,
                                          member.policy_issue_date,
                                          member.policy_declare_date,
                                          member.policy_lapse_date,
                                          member.policy_revival_date,
                                          member.policy_termination_date,
                                          member.policy_suspend_date,
                                          member.policy_unsuspend_date,
                                          member.program,
                                          member.plan,
                                          member.plan_attach_date,
                                          member.plan_expiry_date,
                                          member.special_condition,
                                          member.exclusion,
                                          principle.member_name AS principle_name,
                                          program.`name` AS program_name,
                                          plan.`name` AS plan_name,
                                          member.company,
                                          company.`name` as company_name,
                                          member.branch,
                                          branch.`name` as branch_name,
                                          member.rider,
                                          member.rider_attach_date,
                                          member.rider_expiry_date
                                        FROM
                                          member
                                        LEFT JOIN principle ON member.member_principle = principle.id
                                        LEFT JOIN program ON member.program = program.id
                                        LEFT JOIN plan ON member.plan = plan.id
                                        LEFT JOIN company ON member.company = company.id
                                        LEFT JOIN branch ON member.branch = branch.id
                                        WHERE
                                      member.id = ".$patient.";") or die(mysqli_error($con));
            while ($data3 = mysqli_fetch_row($member_data)) {
              $patient_name = $data3[1];
              $member_id = "'$data3[2]'";
              $member_card = "'$data3[3]'";
              $principle = $data3[4];
              $gender = $data3[5];
              $relation = $data3[6];
              $policy_status = $data3[7];
              $policy_holder = "'$data3[8]'";
              $policy_effective_date = "'$data3[9]'";
              $policy_expiry_date = "'$data3[10]'";
              $policy_issue_date = "'$data3[11]'";
              $policy_declare_date = "'$data3[12]'";          
              if (is_null($data3[13])) {
                $policy_lapse_date = 'NULL';
                $policy_lapse_datex = NULL;
                } else {
                $policy_lapse_date = "'$data3[13]'";
                $policy_lapse_datex = $data3[13];
              }        
              if (is_null($data3[14])) {
                $policy_revival_date = 'NULL';
                $policy_revival_datex = NULL;
              } else {
                $policy_revival_date = "'$data3[14]'";
                $policy_revival_datex = $data3[14];
              }          
              if (is_null($data3[15])) {
                $policy_termination_date = 'NULL';
                $policy_termination_datex = NULL;
              } else {
                $policy_termination_date = "'$data3[15]'";
                $policy_termination_datex = $data3[15];
              }
              if (is_null($data3[16])) {
                $policy_suspend_date = 'NULL';
                $policy_suspend_datex = NULL;
              } else {
                $policy_suspend_date = "'$data3[16]'";
                $policy_suspend_datex = $data3[16];
              }
              if (is_null($data3[17])) {
                $policy_unsuspend_date = 'NULL';
                $policy_unsuspend_datex = NULL;
              } else {
                $policy_unsuspend_date = "'$data3[17]'";
                $policy_unsuspend_datex = $data3[17];
              }
              $program = $data3[18];
              $plan = $data3[19];
              $plan_attach_date = "'$data3[20]'";
              $plan_expiry_date = "'$data3[21]'";
              $special_condition = "'$data3[22]'";
              $exclusion = "'$data3[23]'";
              if (is_null($data3[27])) {
                $company = 'NULL';
              } else {
                $company = $data3[27];
              }
              if (is_null($data3[29])) {
                $branch = 'NULL';
              } else {
                $branch = $data3[29];
              }
              if (is_null($data3[31])) {
                $rider = 'NULL';
              } else {
                $rider = $data3[31];
              }
              if (is_null($data3[32])) {
                $rider_attach_date = 'NULL';
              } else {
                $rider_attach_date = "'$data3[32]'";
              }
              if (is_null($data3[33])) {
                $rider_expiry_date = 'NULL';
              } else {
                $rider_expiry_date = "'$data3[33]'";
              }

              $patient_namex = $data3[1];
              $member_idx = $data3[2];
              $member_cardx = $data3[3];
              $principlex = $data3[24];
              $genderx = $data3[5];
              $relationx = $data3[6];
              $policy_statusx = $data3[7];
              $policy_holderx = $data3[8];
              $policy_effective_datex = $data3[9];
              $policy_expiry_datex = $data3[10];
              $policy_declare_datex = $data3[12];
              $programx = $data3[25];
              $policy_issue_datex = $data3[11];
              $planx = $data3[26];
              $plan_attach_datex = $data3[20];
              $plan_expiry_datex = $data3[21];
              $special_conditionx = $data3[22];
              $exclusionx = $data3[23]; 
              $companyx = $data3[28];
              $branchx = $data3[30];
            }
          }
          // CEK PROVIDER
          if (empty($data[14]) || is_null($data[14])) {
            if (!empty($data[15])) {
              $verify_rs = "SELECT provider.id, provider.full_name FROM provider WHERE provider.full_name LIKE '%".$data[15]."%';";
              $r_verify_rs = mysqli_query($con, $verify_rs) or die(mysqli_error($con));
              $count = mysqli_num_rows($r_verify_rs);
              if ($count > 0) {
                while ($data4 = mysqli_fetch_row($r_verify_rs)) {
                  $id_provider = $data4[0];
                  $provider_name = $data4[1];
                  $provider_error_remarks = $data[15];
                }
              } else {
                $id_provider = 310;
                $provider_name = "NON PARTICIPATING HOSPITAL";
                $provider_error_remarks = "Not Found";
              }
            } else {
              $id_provider = 'NULL';
              $provider_name = $data[15];
              $provider_error_remarks = "Not Found";
            }
          } else {
            $verify_rs2 = "SELECT provider.id, provider.full_name FROM provider WHERE provider.id = ".$data[14].";";
            $r_verify_rs2 = mysqli_query($con, $verify_rs2) or die(mysqli_error($con));
            $count = mysqli_num_rows($r_verify_rs2);
            if ($count > 0) {
              while ($data6 = mysqli_fetch_row($r_verify_rs2)) {
                $id_provider = $data6[0];
                $provider_name = $data6[1];
                $provider_error_remarks = NULL;
              }
            } else {
              $id_provider = $data[14];
              $provider_name = 'NULL';
              $provider_error_remarks = "Not Found";
            }
          }

          if ($patient == 'NULL' || $id_provider == 'NULL') {
            $status = 2;
          } else {
            $status = 1;
          }

          $create_by = $data[2];
          $receive_date = $data[3];
          $category = $data[4];
          $type = $data[5];
          $admission_date = $data[6];
          $admission_time = $data[7];
          $discharge_date = $data[8];
          $discharge_doctor = $data[9];
          $diagnosis = $data[10];
          $bill_no = $data[11];
          $bill_issue_date = date('Y-m-d');
          $bill_due_date = date('Y-m-d', strtotime("+7 days"));
          //$provider = $data[15];
          $other_provider = $data[16];
          $other_provider_city = $data[17];
          $other_provider_country = $data[18];
          //$policy_no = $data[19];
          //$client = $data[21];

          $update = "UPDATE `r_create_case`
                    SET 
                     `status` = ".$status.",
                     `created_by` = '".$create_by."',
                     `receive_date` = '".$receive_date."',
                     `category` = ".$category.",
                     `type` = ".$type.",
                     `admission_date` = '".$admission_date."',
                     `admission_time` = '".$admission_time."',
                     `discharge_date` = '".$discharge_date."',
                     `discharge_doctor` = '".$discharge_doctor."',
                     `diagnosis` = '".$diagnosis."',
                     `bill_no` = '".$bill_no."',
                     `bill_issue_date` = '".$bill_issue_date."',
                     `bill_due_date` = '".$bill_due_date."',
                     `id_provider` = ".$id_provider.",
                     `provider_error_remarks` = '".$provider_error_remarks."',
                     `provider` = '".$provider_name."',
                     `other_provider` = '".$other_provider."',
                     `other_provider_city` = '".$other_provider_city."',
                     `other_provider_country` = ".$other_provider_country.",
                     `policy_no` = '".$policy_no."',
                     `dob` = ".$dob.",
                     `client` = ".$client.",
                     `company` = ".$company.",
                     `branch` = ".$branch.",
                     `patient_error_remarks` = '".$patient_error_remarks."',
                     `patient` = ".$patient.",
                     `member_id` = ".$member_id.",
                     `member_card` = ".$member_card.",
                     `principle` = ".$principle.",
                     `gender` = ".$gender.",
                     `relation` = ".$relation.",
                     `policy_status` = ".$policy_status.",
                     `policy_holder` = ".$policy_holder.",
                     `policy_effective_date` = ".$policy_effective_date.",
                     `policy_expiry_date` = ".$policy_expiry_date.",
                     `policy_issue_date` = ".$policy_issue_date.",
                     `policy_declare_date` = ".$policy_declare_date.",
                     `policy_lapse_date` = ".$policy_lapse_date.",
                     `policy_revival_date` = ".$policy_revival_date.",
                     `policy_termination_date` = ".$policy_termination_date.",
                     `policy_suspend_date` = ".$policy_suspend_date.",
                     `policy_unsuspend_date` = ".$policy_unsuspend_date.",
                     `program` = ".$program.",
                     `plan` = ".$plan.",
                     `plan_attach_date` = ".$plan_attach_date.",
                     `plan_expiry_date` = ".$plan_expiry_date.",
                     `rider` = ".$rider.",
                     `rider_attach_date` = ".$rider_attach_date.",
                     `rider_expiry_date` = ".$rider_expiry_date.",
                     `special_condition` = ".$special_condition.",
                     `exclusion` = ".$exclusion.",
                     `edit_date` = NOW()
                    WHERE
                      (`id` = ".$data[0].");
                    ";
        $update_r = mysqli_query($con, $update) or die(mysqli_error($con));
        if ($update_r) {
          $id = $data[0];
          $Excel->getActiveSheet()->setCellValue('A'.$row, $id);
          $Excel->getActiveSheet()->setCellValue('B'.$row, $status);
          $Excel->getActiveSheet()->setCellValue('C'.$row, $create_by);
          $Excel->getActiveSheet()->setCellValue('D'.$row, $receive_date);
          $Excel->getActiveSheet()->setCellValue('E'.$row, $category);
          $Excel->getActiveSheet()->setCellValue('F'.$row, $type);    
          $Excel->getActiveSheet()->setCellValue('G'.$row, $admission_date);
          $Excel->getActiveSheet()->setCellValue('H'.$row, $admission_time);
          $Excel->getActiveSheet()->setCellValue('I'.$row, $discharge_date);
          $Excel->getActiveSheet()->setCellValue('J'.$row, $discharge_doctor);
          $Excel->getActiveSheet()->setCellValue('K'.$row, $diagnosis);
          $Excel->getActiveSheet()->setCellValue('L'.$row, $bill_no);
          $Excel->getActiveSheet()->setCellValue('M'.$row, $bill_issue_date);
          $Excel->getActiveSheet()->setCellValue('N'.$row, $bill_due_date);
          if ($id_provider == 'NULL') {
            $Excel->getActiveSheet()->setCellValue('O'.$row, $provider_error_remarks);
          } else {
            $Excel->getActiveSheet()->setCellValue('O'.$row, $id_provider);
          }
          $Excel->getActiveSheet()->setCellValue('P'.$row, $provider_name);
          $Excel->getActiveSheet()->setCellValue('Q'.$row, $other_provider);
          $Excel->getActiveSheet()->setCellValue('R'.$row, $other_provider_city);
          $Excel->getActiveSheet()->setCellValue('S'.$row, $other_provider_country);
          $Excel->getActiveSheet()->setCellValueExplicit('T'.$row, $policy_no,PHPExcel_Cell_DataType::TYPE_STRING);
          $Excel->getActiveSheet()->setCellValue('U'.$row, $dobx);
          $Excel->getActiveSheet()->setCellValue('V'.$row, $client);
          $Excel->getActiveSheet()->setCellValue('W'.$row, $companyx);
          $Excel->getActiveSheet()->setCellValue('X'.$row, $branchx);
          if ($patient == 'NULL') {
            $Excel->getActiveSheet()->setCellValue('Y'.$row, $patient_error_remarks);
          } else {
            $Excel->getActiveSheet()->setCellValue('Y'.$row, $patient);      
          }
          $Excel->getActiveSheet()->setCellValue('Z'.$row, $patient_namex);
          $Excel->getActiveSheet()->setCellValueExplicit('AA'.$row, $member_idx,PHPExcel_Cell_DataType::TYPE_STRING);
          $Excel->getActiveSheet()->setCellValueExplicit('AB'.$row, $member_cardx,PHPExcel_Cell_DataType::TYPE_STRING);
          $Excel->getActiveSheet()->setCellValue('AC'.$row, $principlex);
          $Excel->getActiveSheet()->setCellValue('AD'.$row, $genderx);
          $Excel->getActiveSheet()->setCellValue('AE'.$row, $relationx);
          $Excel->getActiveSheet()->setCellValue('AF'.$row, $policy_statusx);
          $Excel->getActiveSheet()->setCellValue('AG'.$row, $policy_holderx);
          $Excel->getActiveSheet()->setCellValue('AH'.$row, $policy_effective_datex);
          $Excel->getActiveSheet()->setCellValue('AI'.$row, $policy_expiry_datex);
          $Excel->getActiveSheet()->setCellValue('AJ'.$row, $policy_issue_datex);
          $Excel->getActiveSheet()->setCellValue('AK'.$row, $policy_declare_datex);
          $Excel->getActiveSheet()->setCellValue('AL'.$row, $policy_lapse_datex);
          $Excel->getActiveSheet()->setCellValue('AM'.$row, $policy_revival_datex);
          $Excel->getActiveSheet()->setCellValue('AN'.$row, $policy_termination_datex);
          $Excel->getActiveSheet()->setCellValue('AO'.$row, $policy_suspend_datex);
          $Excel->getActiveSheet()->setCellValue('AP'.$row, $policy_unsuspend_datex);
          $Excel->getActiveSheet()->setCellValue('AQ'.$row, $programx);
          $Excel->getActiveSheet()->setCellValue('AR'.$row, $planx);
          $Excel->getActiveSheet()->setCellValue('AS'.$row, $plan_attach_datex);
          $Excel->getActiveSheet()->setCellValue('AT'.$row, $plan_expiry_datex);
          $Excel->getActiveSheet()->setCellValue('AU'.$row, $special_conditionx);
          $Excel->getActiveSheet()->setCellValue('AV'.$row, $exclusionx);                                                                                                                                                                                                                                                                                 
          }
          // END

        }
      }
    }
  }
}
$Excel->getActiveSheet()->setTitle('Result');
$objWriter = PHPExcel_IOFactory::createWriter($Excel, 'Excel2007');
$objWriter->save('Y:/ADMIN/EXTERNAL USE/UPLOAD_CASE/'.$file2);
echo '<script>alert("Please check result file");window.location.href="index2.html";</script>';
