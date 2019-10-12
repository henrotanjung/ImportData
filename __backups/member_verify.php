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
$file='create_case_'.$date.'';
$inputFileName = 'C:/xampp/htdocs/insecure/intsys/upload_file/file/'.$file.'.xlsx';
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
//echo "Berikut data case yang berhasil diupdate : <br/>";
for ($row = 2; $row <= $highestRow;$row++){
	$dataRow = $sheet->rangeToArray('A'.$row.':'.$highestColumn.$row,null, true, true, false);
	if (is_array($dataRow)) {
		foreach ($dataRow as $key => $data) { 
      if (empty($data[0])) {
        // CEK MEMBER
        if (empty($data[18])) {
          if (!empty($data[19]) || !empty($data[17])) {
            $verify_member = "SELECT member.id FROM member WHERE member.client = ".$data[19]." AND member.policy_no = '".$data[17]."';";
            $r_verify_member = mysqli_query($con, $verify_member);
            $count = mysqli_num_rows($r_verify_member);
            if ($count = 1) {
              while ($data1 = mysqli_fetch_row($r_verify_member)) {
                  $patient = $data1[0];
                  $patient_error_remarks = NULL;
              }
            }
            else {
              $patient = 'NULL';
              $patient_error_remarks = "Not Found"; 
            }
          }else{
             $patient = 'NULL';
             $patient_error_remarks = "Please Check Client or Dob";
          }
        } else {
          if (!empty($data[19]) || !empty($data[17])) {
            $verify_member2 = "SELECT member.id FROM member WHERE member.client = ".$data[19]." AND member.policy_no = '".$data[17]."' AND member.member_dob = '".$data[18]."';";
            $r_verify_member2 = mysqli_query($con, $verify_member2);
            $count = mysqli_num_rows($r_verify_member2);
            if ($count = 1) {
              while ($data4 = mysqli_fetch_row($r_verify_member2)) {
                  $patient = $data4[0];
                  $patient_error_remarks = NULL;
              }
            }
            else {
              $patient = 'NULL';
              $patient_error_remarks = "Not Found"; 
            }
          }else{
             $patient = 'NULL';
             $patient_error_remarks = "Please Check Client or Dob";
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
                                  member.exclusion
                                FROM
                                  member
                                WHERE
                                  member.id = ".$patient.";") or die(mysqli_error($con));
          while ($data3 = mysqli_fetch_row($member_data)) {
            $patient_name = $data3[1];
            $member_id = $data3[2];
            $member_card = $data3[3];
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
            } else {
              $policy_lapse_date = "'$data3[13]'";
            }        
            if (is_null($data3[14])) {
              $policy_revival_date = 'NULL';
            } else {
              $policy_revival_date = "'$data3[14]'";
            }          
            if (is_null($data3[15])) {
              $policy_termination_date = 'NULL';
            } else {
              $policy_termination_date = "'$data3[15]'";
            }
            if (is_null($data3[16])) {
              $policy_suspend_date = 'NULL';
            } else {
              $policy_suspend_date = "'$data3[16]'";
            }
            if (is_null($data3[17])) {
              $policy_unsuspend_date = 'NULL';
            } else {
              $policy_unsuspend_date = "'$data3[17]'";
            }
            $program = $data3[18];
            $plan = $data3[19];
            $plan_attach_date = "'$data3[20]'";
            $plan_expiry_date = "'$data3[21]'";
            $special_condition = "'$data3[22]'";
            $exclusion = "'$data3[23]'";          
          }
        }

        // CEK PROVIDER
        if (!empty($data[15])) {
          $verify_rs = "SELECT provider.id FROM provider WHERE provider.full_name LIKE '%".$data[15]."%';";
          $r_verify_rs = mysqli_query($con, $verify_rs) or die(mysqli_error($con));
          $count = mysqli_num_rows($r_verify_rs);
          if ($count > 0) {
            while ($data2 = mysqli_fetch_row($r_verify_rs)) {
              $id_provider = $data2[0];
              $provider_error_remarks = "";
            }
          } else {
            $id_provider = 310;
            $provider_error_remarks = "Not Found";
          }
        } else {
          $id_provider = 'NULL';
          $provider_error_remarks = "Please Check Provider Name";
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
        $policy_no = $data[17];
        $dob = $data[18];
        $client = $data[19];

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
                            `policy_no`,
                            `dob`,
                            `client`,
                            `patient_error_remarks`,
                            `patient`,
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
                              '".$provider."',
                              '".$other_provider."',
                              '".$policy_no."',
                              '".$dob."',
                              ".$client.",
                              '".$patient_error_remarks."',
                              ".$patient.",
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
                              ".$special_condition.",
                              ".$exclusion.",
                              NOW(),
                              NULL
                            );";
      //print_r($insert);
      $result = mysqli_query($con, $insert) or die(mysqli_error($con));
      } else {
        // IF EXISTS
        $cek_id = mysqli_query($con, "SELECT r_create_case.id FROM r_create_case WHERE r_create_case.id = ".$data[0].";") or die(mysqli_error($con));
        $count = mysqli_num_rows($cek_id);
        if ($count = 1) {
          // CEK MEMBER
          if (empty($data[18])) {
            if (!empty($data[19]) || !empty($data[17])) {
              $verify_member = "SELECT member.id FROM member WHERE member.client = ".$data[19]." AND member.policy_no = '".$data[17]."';";
              $r_verify_member = mysqli_query($con, $verify_member) or die(mysqli_error($con));
              $count = mysqli_num_rows($r_verify_member);
              if ($count = 1) {
                while ($data1 = mysqli_fetch_row($r_verify_member)) {
                    $patient = $data1[0];
                    $patient_error_remarks = NULL;
                }
              }
              else {
                $patient = 'NULL';
                $patient_error_remarks = "Not Found"; 
              }
            }else{
               $patient = 'NULL';
               $patient_error_remarks = "Please Check Client Or Dob";
            }
          } else {
            if (!empty($data[19]) || !empty($data[17])) {
              $verify_member2 = "SELECT member.id FROM member WHERE member.client = ".$data[19]." AND member.policy_no = '".$data[17]."' AND member.member_dob = '".$data[18]."';";
              $r_verify_member2 = mysqli_query($con, $verify_member2);
              $count = mysqli_num_rows($r_verify_member2);
              if ($count = 1) {
                while ($data4 = mysqli_fetch_row($r_verify_member2)) {
                    $patient = $data4[0];
                    $patient_error_remarks = NULL;
                }
              }
              else {
                $patient = 'NULL';
                $patient_error_remarks = "Not Found"; 
              }
            }else{
               $patient = 'NULL';
               $patient_error_remarks = "Please Check Client Or Dob";
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
                                    member.exclusion
                                  FROM
                                    member
                                  WHERE
                                    member.id = ".$patient.";") or die(mysqli_error($con));
            while ($data3 = mysqli_fetch_row($member_data)) {
              $patient_name = $data3[1];
              $member_id = $data3[2];
              $member_card = $data3[3];
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
              } else {
                $policy_lapse_date = "'$data3[13]'";
              }        
              if (is_null($data3[14])) {
                $policy_revival_date = 'NULL';
              } else {
                $policy_revival_date = "'$data3[14]'";
              }          
              if (is_null($data3[15])) {
                $policy_termination_date = 'NULL';
              } else {
                $policy_termination_date = "'$data3[15]'";
              }
              if (is_null($data3[16])) {
                $policy_suspend_date = 'NULL';
              } else {
                $policy_suspend_date = "'$data3[16]'";
              }
              if (is_null($data3[17])) {
                $policy_unsuspend_date = 'NULL';
              } else {
                $policy_unsuspend_date = "'$data3[17]'";
              }
              $program = $data3[18];
              $plan = $data3[19];
              $plan_attach_date = "'$data3[20]'";
              $plan_expiry_date = "'$data3[21]'";
              $special_condition = "'$data3[22]'";
              $exclusion = "'$data3[23]'";          
            }
          }

          // CEK PROVIDER
          if (!empty($data[15])) {
            $verify_rs = "SELECT provider.id FROM provider WHERE provider.full_name LIKE '%".$data[15]."%';";
            $r_verify_rs = mysqli_query($con, $verify_rs) or die(mysqli_error($con));
            $count = mysqli_num_rows($r_verify_rs);
            if ($count > 0) {
              while ($data2 = mysqli_fetch_row($r_verify_rs)) {
                $id_provider = $data2[0];
                $provider_error_remarks = "";
              }
            } else {
              $id_provider = 310;
              $provider_error_remarks = "Not Found";
            }
          } else {
            $id_provider = 'NULL';
            $provider_error_remarks = "Please Check Provider Name";
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
          $policy_no = $data[17];
          $dob = $data[18];
          $client = $data[19];

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
                              `policy_no`,
                              `dob`,
                              `client`,
                              `patient_error_remarks`,
                              `patient`,
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
                                '".$provider."',
                                '".$other_provider."',
                                '".$policy_no."',
                                '".$dob."',
                                ".$client.",
                                '".$patient_error_remarks."',
                                ".$patient.",
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
                                ".$special_condition.",
                                ".$exclusion.",
                                NOW(),
                                NULL
                              );";
        //print_r($insert);
        $result = mysqli_query($con, $insert) or die(mysqli_error($con));
        } else {
          // IF NOT EXISTS
          if (empty($data[18])) {
            if (!empty($data[19]) || !empty($data[17])) {
              $verify_member = "SELECT member.id FROM member WHERE member.client = ".$data[19]." AND member.policy_no = '".$data[17]."';";
              $r_verify_member = mysqli_query($con, $verify_member) or die(mysqli_error($con));
              $count = mysqli_num_rows($r_verify_member);
              if ($count = 1) {
                while ($data1 = mysqli_fetch_row($r_verify_member)) {
                    $patient = $data1[0];
                    $patient_error_remarks = NULL;
                }
              }
              else {
                $patient = 'NULL';
                $patient_error_remarks = "Not Found"; 
              }
            }else{
               $patient = 'NULL';
               $patient_error_remarks = "Please Check Client Or Dob";
            }
          } else {
            if (!empty($data[19]) || !empty($data[17])) {
              $verify_member2 = "SELECT member.id FROM member WHERE member.client = ".$data[19]." AND member.policy_no = '".$data[17]."' AND member.member_dob = '".$data[18]."';";
              $r_verify_member2 = mysqli_query($con, $verify_member2) or die(mysqli_error($con));
              $count = mysqli_num_rows($r_verify_member2);
              if ($count = 1) {
                while ($data4 = mysqli_fetch_row($r_verify_member2)) {
                    $patient = $data4[0];
                    $patient_error_remarks = NULL;
                }
              }
              else {
                $patient = 'NULL';
                $patient_error_remarks = "Not Found"; 
              }
            }else{
               $patient = 'NULL';
               $patient_error_remarks = "Please Check Client, Policy No or DOB";
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
                                    member.exclusion
                                  FROM
                                    member
                                  WHERE
                                    member.id = ".$patient.";") or die(mysqli_error($con));
            while ($data3 = mysqli_fetch_row($member_data)) {
              $patient_name = $data3[1];
              $member_id = $data3[2];
              $member_card = $data3[3];
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
              } else {
                $policy_lapse_date = "'$data3[13]'";
              }        
              if (is_null($data3[14])) {
                $policy_revival_date = 'NULL';
              } else {
                $policy_revival_date = "'$data3[14]'";
              }          
              if (is_null($data3[15])) {
                $policy_termination_date = 'NULL';
              } else {
                $policy_termination_date = "'$data3[15]'";
              }
              if (is_null($data3[16])) {
                $policy_suspend_date = 'NULL';
              } else {
                $policy_suspend_date = "'$data3[16]'";
              }
              if (is_null($data3[17])) {
                $policy_unsuspend_date = 'NULL';
              } else {
                $policy_unsuspend_date = "'$data3[17]'";
              }
              $program = $data3[18];
              $plan = $data3[19];
              $plan_attach_date = "'$data3[20]'";
              $plan_expiry_date = "'$data3[21]'";
              $special_condition = "'$data3[22]'";
              $exclusion = "'$data3[23]'";          
            }
          }

          // CEK PROVIDER
          if (!empty($data[15])) {
            $verify_rs = "SELECT provider.id FROM provider WHERE provider.full_name LIKE '%".$data[15]."%';";
            $r_verify_rs = mysqli_query($con, $verify_rs) or die(mysqli_error($con));
            $count = mysqli_num_rows($r_verify_rs);
            if ($count > 0) {
              while ($data2 = mysqli_fetch_row($r_verify_rs)) {
                $id_provider = $data2[0];
                $provider_error_remarks = "";
              }
            } else {
              $id_provider = 310;
              $provider_error_remarks = "Not Found";
            }
          } else {
            $id_provider = 'NULL';
            $provider_error_remarks = "Please Check Provider Name";
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
          $policy_no = $data[17];
          $dob = $data[18];
          $client = $data[19];

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
                     `provider` = '".$provider."',
                     `other_provider` = '".$other_provider."',
                     `policy_no` = '".$policy_no."',
                     `dob` = '".$dob."',
                     `client` = ".$client.",
                     `patient_error_remarks` = '".$patient_error_remarks."',
                     `patient` = ".$patient.",
                     `gender` = ".$gender.",
                     `relation` = ".$gender.",
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
                     `special_condition` = ".$special_condition.",
                     `exclusion` = ".$exclusion.",
                     `edit_date` = NOW()
                    WHERE
                      (`id` = ".$data[0].");
                    ";
          $update_r = mysqli_query($con, $update) or die(mysqli_error($con));
        }        
      }
    } 
  }
}