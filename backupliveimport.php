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

// no backslash at the end.

$folder = '//192.168.0.20/Z Folder/IT/INTERNAL ONLY/003 Internal Minutes of Meeting/Programmer Team/006 Hendro/';
$folder = './';
// $file_xlsx = 'addition_20191007.xlsx';
$file_xlsx = 'test_member_his_tpa.xlsx';

echo "<pre>";

$Import = Import::open(path($folder) . $file_xlsx);
// $Import->setMap([
// 	'A' => 'client',
// 	'B' => 'branch',
//     'C'	=> 'company'	]);
    

    

$Import->setMap([
	'A' => 'client',
	'B' => 'branch',
	'C'	=> 'company',	
	// 'D'	=> 'department',	
	// 'E'	=> 'position',	
	'D'	=> 'id',	
	'E'	=> 'member_name',	
	'F'	=> 'member_id',	
	'G'	=> 'member_card',	
	// 'J'	=> 'member_card_edc',	
	'H'	=> 'member_dob',	
	'I'	=> 'member_gender',	
	'J'	=> 'member_marital',	
	'K'	=> 'member_address',	
	'L'	=> 'member_country',	
	'M'	=> 'member_province',	
	'N'	=> 'member_city',	
	'O'	=> 'member_postal',	
	'P'	=> 'member_phone',	
	'Q'	=> 'member_mobile',	
	'R'	=> 'member_email',	
	'S'	=> 'member_password',	
	'T'	=> 'member_relation',	
    'U'	=> 'member_principle',
    'V'	=> 'policy_no',	
    'W'	=> 'policy_holder',			
    'X'	=> 'policy_issue_date',
    'Y'	=> 'policy_declare_date',	
    'Z'	=> 'policy_effective_date',			
    'AA'	=> 'policy_effective_date_card',			
    'AB'	=> 'policy_expiry_date',			
    'AC'	=> 'policy_lapse_date',			
    'AD'	=> 'policy_revival_date',			
    'AE'	=> 'policy_termination_date',			
    // 'AF'	=> 'policy_suspend_date',			
    // 'AG'	=> 'policy_unsuspend_date',			
    'AF'	=> 'policy_status',			
    'AG'	=> 'program',			
    'AH'	=> 'plan',			
    'AI'	=> 'plan_external',			
    'AJ'	=> 'plan_attach_date',			
    'AK'	=> 'plan_expiry_date',			
    'AL'	=> 'rider',			
    'AM'	=> 'rider_attach_date',			
    'AN'	=> 'rider_expiry_date',			
    'AO'	=> 'special_condition',			
    'AP'	=> 'exclusion',			
    // 'AI'	=> 'member_remarks',			member_remarks
    // 'AI'	=> 'remarks_by',			remarks_by
    // 'AI'	=> 'remarks_date',			remarks_date
    'AQ'	=> 'bank',			
    'AR'	=> 'location',			
    'AS'	=> 'account_no',			
    'AT'	=> 'on_behalf_of',			
    // 'AI'	=> 'medisclick_activation',	medisclick_activation		
    // 'AI'	=> 'medisclick_activation_date',		medisclick_activation_date	
    'AU'	=> 'created_by',			
    'AV'	=> 'create_date',			
    'AW'	=> 'edited_by',			
    'AX'	=> 'edit_date',			
    // 'member_username',			
    // 'member_name_encrypt',			
    // 'member_dob_encrypt',			
    // 'member_phone_encrypt',			
    // 'flag_login_medisclick',			
    // 'member_device',			
    // 'member_os',			
    // 'member_app',			
    // 'member_serial',			
    // 'member_regid',			
    // 'medisclick_email',			
    // 'medisclick_password',			
    // 'medisclick_phone',			
    // 'member_card_edc_01'	
		
]); 
    
// $Import->setMapAuto([
// 	'client',
// 	'branch',
//     'company',	
//     // 'department',	
//     // 'position',	
//     'id',	
//     'member_name',	
// 	'member_id',	
// 	'member_card',	
// 	// 'member_card_edc',	//member_card_edc
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
//     // 'policy_suspend_date',			//policy_suspend_date
//     // 'policy_unsuspend_date',			//policy_unsuspend_date
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
//     // 'member_remarks',			//member_remarks
//     // 'remarks_by',			//remarks_by
//     // 'remarks_date',			//remarks_date
//     'bank',			
//     'location',			
//     'account_no',			
//     'on_behalf_of',			
//     // 'medisclick_activation',	//medisclick_activation		
//     // 'medisclick_activation_date',		//medisclick_activation_date	
//     'created_by',			
//     'create_date',			
//     'edited_by',			
//     // 'edit_date',	
//     // 'member_username',			
//     // 'member_name_encrypt',			
//     // 'member_dob_encrypt',			
//     // 'member_phone_encrypt',			
//     // 'flag_login_medisclick',			
//     // 'member_device',			
//     // 'member_os',			
//     // 'member_app',			
//     // 'member_serial',			
//     // 'member_regid',			
//     // 'medisclick_email',			
//     // 'medisclick_password',			
//     // 'medisclick_phone',			
//     // 'member_card_edc_01' 	
// ]); 

$query = Database::on('dblocal_dev_aia')->builder();
$datas = array();

$Import->each(function($row, $index) use (&$datas) {
    if ($row['client'] === ''){
        return;
    }
    // unset($row['department']);
    // unset($row['position']);
	$datas[] = $row;		
}, true);
var_dump($datas);

$res_sql = $query->into('member')->insert($datas);
// echo $res_sql;
