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

$workdir 	= $GLOBALS['config']['path']['finance'];
$oldupload 	= $config['path']['oldupload'];
$newupload 	= $config['path']['newupload'];

// Config :
//===========================================================
$dir_master 	= path($workdir . '\obv_remarks\data_master');
$dir_from 		= path($workdir . '\obv_remarks\file');
$dir_result 	= path($workdir . '\obv_remarks\result');
$dir_upload 	= path($newupload);
//===========================================================
$file_xlsx 		= 'upload_obv_remarks_'.date('Y-m-d').'.xlsx';
//===========================================================

echo '<pre>';

if(!file_exists(path($dir_master) . $file_xlsx)) {
	
	echo('Cannot open file [file not found] : '.$file_xlsx . "\r\n\r\n");

	echo('Check Folder : <a href="file:///'.$dir_master.'">'.$dir_master.'</a>.' . "\r\n\r\n");

	foreach(scandir($dir_master, 1) as $file) {

		if($file === '.' || $file === '..') {
			continue;	
		}

		echo $file;
		echo "\r\n";

	}

	die();

}

if(!isset($_POST['start'])) {
	echo "execute file : [$file_xlsx] ? ";
	echo '
		<form action="" method="post">
			<button name="start" type="submit">Submit</button>
		</form>
	';
	die();
}

Loader::lib('import');
Loader::lib('export');

function search_index($search_id, $column, $array) {

	foreach ($array as $key => $val) {
		if ($val->{$column} === $search_id) {
			return $key;
		}
	}

	return NULL;

}

function compare_value($search_value, $column, $array) {

	if(trim($array->{$column}) === trim($search_value)) {

		return true;

	}

	return false;

}

$Import = Import::open(path($dir_master) . $file_xlsx);

$Import->setMap([
	'A' => 'id',
	'B' => 'patient',
	'C'	=> 'client',
	'D'	=> 'original_bill_verified_remarks',
	'E' => 'policy_no'
]);

$Export = Export::open(path($dir_result) . 'Result_' . $file_xlsx);

$Export->setMap([
	'A' => 'id',
	'B' => 'patient',
	'C'	=> 'client',
	'D'	=> 'original_bill_verified_remarks',
	'E' => 'policy_no',
	'F'	=> 'result',
	'G'	=> 'reason'
]);

$Export->setImport($Import);

$case 			= [];

$logs 			= [];

$logs_detail 	= [];

$update_case 	= [];

// import from excel and insert into variable $case
//------------------------------------------------------------
$Import->each(function($row, $index) use (&$case, &$compare) {

	if($row['id'] === '') {
		return;
	}

	$row['id']		= trim($row['id']);
	
	$row['$index'] 	= $index;
	$case[] 		= $row;


}, TRUE);

// get from database by case id
//-----------------------------------------------------------
$query = Database::on('his-tpa-dev')->builder();

$rows   = 
	$query
		->select(
			'case.id',
			'case.client',
			'case.policy_no',
			'member.member_name'
		)
		->from('`case`')
		->joinLeft('member', 'member.id', '`case`.patient')
		->whereIn('`case`.id', array_column($case, 'id'))
		->get();

unset($query);

// check diffrence and not found id
//---------------------------------------------------------
foreach ($case as $i => $valcase) {
	
	$id 	= search_index($valcase['id'], 'id', $rows);
	
	$index 	= $valcase['$index'];

	$row 	= $valcase;

	$reason = '';
	$status = '';

	if(is_null($id)) {

		// error occured ( case id not found )
		//------------------------------------------------
		$Export->on('result', $index)->setText('error');
		$Export->on('reason', $index)->setText($reason = 'Case ID Not Found.');
		$status 	= 'error';

	}// end if
	else if(!compare_value($valcase['client'], 'client', $rows[$id])) {

		// error occured ( case id not found )
		//------------------------------------------------
		$Export->on('result', $index)->setText('error');
		$Export->on('reason', $index)->setText($reason = 'Client ['.$rows[$id]->member_name.'] not match with Case ID, on database ('.$rows[$id]->client.').');
		$status 	= 'error';

	}
	else if(!compare_value($valcase['policy_no'], 'policy_no', $rows[$id])) {

		// error occured ( case id not found )
		//------------------------------------------------
		$Export->on('result', $index)->setText('error');
		$Export->on('reason', $index)->setText($reason = 'Policy NO ['.$rows[$id]->member_name.'] not match with Case ID, on database ('.$rows[$id]->policy_no.').');
		$status 	= 'error';

	}
	else {
		
		// success
		//------------------------------------------------
		$Export->on('result', $index)->setText('success');
		$status 	= 'success';

		$update_case[] = [
			'id' => $valcase['id'],
			'original_bill_verified_remarks' => $valcase['original_bill_verified_remarks']
		];

	}

	// create logs
	//--------------------------------------------------------
	$logs[]		= [
		'code' 			=> $code = uniquuid(),
		'log_number'	=> $row['id'],
		'create_date'	=> date('Y-m-d H:i:s'),
		'created_by'	=> 'system',
		'status'		=> $status,
		'reason'		=> $reason
	];

	$logs_detail[] = [
		'code'			=> $code,
		'line'			=> 1,
		'attribute'		=> 'module',
		'value'			=> 'finance'
	];

	$logs_detail[] = [
		'code'			=> $code,
		'line'			=> 2,
		'attribute'		=> 'patient',
		'value'			=> $row['patient']
	];

	$logs_detail[] = [
		'code'			=> $code,
		'line'			=> 3,
		'attribute'		=> 'client',
		'value'			=> $row['client']
	];

	$logs_detail[] = [
		'code'			=> $code,
		'line'			=> 4,
		'attribute'		=> 'policy_no',
		'value'			=> $row['policy_no']
	];

	$logs_detail[] = [
		'code'			=> $code,
		'line'			=> 5,
		'attribute'		=> 'original_bill_verified_remarks',
		'value'			=> $row['original_bill_verified_remarks']
	];

}

print_r($update_case);

$success = TRUE;

// Save result file
//-----------------------------------------------------------------------
	if(!$Export->save()) {

		Log::create('finance', basename(__FILE__), $Export->log(), []);

		$success = FALSE;

	}
//-----------------------------------------------------------------------

// Update case_log
//-----------------------------------------------------------------------
	$query = Database::on('his-tpa-dev')->builder();

	try {

		$query
			->into('`case`')
			->update($update_case);

	}catch(DatabaseException $e) {

		Log::create('finance', basename(__FILE__), $e->getMessage(), []);

		$success = FALSE;

	}

	unset($query);
//-----------------------------------------------------------------------


// Insert into log
//-----------------------------------------------------------------------
	$query = Database::on('local')->builder();

	try {
		
		$query->transaction();

		$query
			->into('logs')
			->insert($logs);

		$query
			->into('logs_detail')
			->insert($logs_detail);

		$query->commit();

	}catch(DatabaseException $e) {

		$query->rollback();

		Log::create('finance', basename(__FILE__), $e->getMessage(), []);

		$success = FALSE;

	}

	unset($query);
//----------------------------------------------------------------------

if(!$success) {
	throw new Exception("Error occured check log.", 1);
}