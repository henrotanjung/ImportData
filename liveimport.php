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
    'AX'	=> 'edit_date'	    		
]); 
   
$query = Database::on('dblocal_dev_aia')->builder();
$datas = array();

$Import->each(function($row, $index) use (&$datas) {
    if ($row['client'] === ''){
        return;
    }    
	$datas[] = $row;		
}, true);
var_dump($datas);

$res_sql = $query->into('member')->insert($datas);
// echo $res_sql;
