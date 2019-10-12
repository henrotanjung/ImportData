<?php 
/** 
| Header include for microframework
|========================================================== */
define('ROOTDIR', str_replace('\\', '/', __DIR__ . '/'));

include(ROOTDIR . 'framework/autoloader.php');

use Pusaka\Database\Manager as Database;
use Pusaka\Database\DatabaseException;
use Pusaka\Microframework\Loader;
use Pusaka\Microframework\Log;
use Pusaka\Library\Import;
use Pusaka\Library\Export;
use Pusaka\Library\AAICopy;
//===========================================================

Loader::lib('import');
Loader::lib('export');

$folder = './';

$file_xlsx = 'file_import.xlsx';

echo "<pre>";

$Import = Import::open(path($folder) . $file_xlsx);

$Import->setMap([
	'A' => 'nama',
	// 'B' => 'alamat',
	'D'	=> 'telepon',
]);
$query = Database::on('dblocal')->builder();
$datas = array();

$Import->each(function($row, $index) use (&$datas) {
	// unset($row['apa']);
	$datas[] = $row;

	// $datas[] = [
	// 	'nama' => $row["nama"],
	// 	'alamat' => $row["alamat"],
	// 	'telepon' => $row["telepon"]
	// ];	
	
}, true);
var_dump($datas);

$res_sql = $query->into('data_pegawai')->insert($datas);
// echo $res_sql;


// $Import->setMapAuto([
// 	'client',
// 	'branch',
//     'company',	
//     'department',	
//     'position',	
//     'id',	
//     'member_name',	
// 	'member_id',	
// 	'member_card',	
// 	'member_card_edc',	member_card_edc
// 	'member_dob',	
// 	'member_gender',	
// 	'member_marital',	
// 	'member_address',	
// 	'member_country',	
// 	'member_province',	
// 	'member_city',	
// 	'member_postal',	
// 	'member_phone',	
// 	'member_mobile',	
// 	'member_email',	
// 	'member_password',	
// 	'member_relation',	
//     'member_principle',
//     'policy_no',	
//     'policy_holder',			
//     'policy_issue_date',

//     'policy_effective_date',			
//     'policy_effective_date_card',			
//     'policy_expiry_date',			
//     'policy_lapse_date',			
//     'policy_revival_date',			
//     'policy_termination_date',			
//     'policy_suspend_date',			policy_suspend_date
//     'policy_unsuspend_date',			policy_unsuspend_date
//     'policy_status',			
//     'program',			
//     'plan',			
//     'plan_external',			
//     'plan_attach_date',			
//     'plan_expiry_date',			
//     'rider',			
//     'rider_attach_date',			
//     'rider_expiry_date',			
//     'special_condition',			
//     'exclusion',			
//     'member_remarks',			member_remarks
//     'remarks_by',			remarks_by
//     'remarks_date',			remarks_date
//     'bank',			
//     'location',			
//     'account_no',			
//     'on_behalf_of',			
//     'medisclick_activation',	medisclick_activation		
//     'medisclick_activation_date',		medisclick_activation_date	
//     'created_by',			
//     'create_date',			
//     'edited_by',			
//     'edit_date',			
//     'member_username',			
//     'member_name_encrypt',			
//     'member_dob_encrypt',			
//     'member_phone_encrypt',			
//     'flag_login_medisclick',			
//     'member_device',			
//     'member_os',			
//     'member_app',			
//     'member_serial',			
//     'member_regid',			
//     'medisclick_email',			
//     'medisclick_password',			
//     'medisclick_phone',			
//     'member_card_edc_01'		
// ]); 
