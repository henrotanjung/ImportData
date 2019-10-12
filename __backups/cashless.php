<?php

$con = mysqli_connect("192.168.0.24","luki.handoko","MysticRiver20140310","his-tpa");

// Check connection
if (mysqli_connect_errno())
  {
  echo "Failed to connect to MySQL: " . mysqli_connect_error();
  }

$path = glob('Y:/ADMIN/EXTERNAL USE/UPLOAD_FILE/*');
if (count($path) === 0) {
	echo "<script type='text/javascript'>alert('Could not find files in directory!')</script>"; 
}
else {
	foreach($path as $files){
	    //Use the is_file function to make sure that it is not a directory.
	    if(is_file($files)){
	        $filename = basename($files); 
	        $split = explode('_', $filename);
	        $case_id = $split[0];
	        //$file = $split[1];
	        $query = "UPDATE `his-tpa`.`case` INNER JOIN (SELECT
						`his-tpa`.`case`.id,
						`his-tpa`.`case`.`status`,
						`his-tpa`.`case`.upload_original_bill,
						`his-tpa`.`case`.original_bill_uploaded_by,
						`his-tpa`.`case`.original_bill_upload_date
						FROM
						`his-tpa`.`case`
						WHERE
						`his-tpa`.`case`.`status` in (11,12)
						AND `his-tpa`.`case`.id = '".$case_id."') t ON `his-tpa`.`case`.id = t.id
						SET
						`his-tpa`.`case`.userlevel = 7, 
						 `his-tpa`.`case`.upload_original_bill = IF (`his-tpa`.`case`.upload_original_bill IS NOT NULL, CONCAT(`his-tpa`.`case`.upload_original_bill,',','".$filename."'),'".$filename."'),
						 `his-tpa`.`case`.bill_issue_date = NOW(),
						 `his-tpa`.`case`.bill_due_date = IF(`his-tpa`.`case`.category = 0, DATE_ADD(NOW(),INTERVAL 7 DAY), DATE_ADD(NOW(),INTERVAL 14 DAY)),
						 `his-tpa`.`case`.edited_by = 'Administrator',
						 `his-tpa`.`case`.edit_date = NOW(),
						 `his-tpa`.`case`.`status` = 12";
	        $result = mysqli_query($con,$query);
			if ($result) {
				if (mysqli_affected_rows($con) == 1) {
					$folder = "//10.10.10.27/c$/xampp/htdocs/insecure/intsys/his-tpa/AAint/upload/".$case_id."";
					$srcfile = "Y:/ADMIN/EXTERNAL USE/UPLOAD_FILE/".$filename."";
			    	$destfile = "//10.10.10.27/c$/xampp/htdocs/insecure/intsys/his-tpa/AAint/upload/".$case_id."/".$filename."";
			    	if (!file_exists($folder)) {
			    		mkdir($folder, 0777, true);
			    	}
			   		if (copy($srcfile, $destfile)) {
			   			unlink($srcfile);
			   		}
				}
			}
	    }
	}
echo "<script type='text/javascript'>alert('Upload successfully!')</script>";  
}
?>