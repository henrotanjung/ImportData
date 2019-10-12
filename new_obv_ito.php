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
$dir_master 	= path($workdir . '\obv\data_master');
$dir_from 		= path($workdir . '\obv\file');
$dir_result 	= path($workdir . '\obv\result');
$dir_upload 	= path($newupload);
//===========================================================
$file_xlsx 		= 'ITO2_upload_obv_'.date('Y-m-d').'.xlsx';
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

$Import = Import::open(path($dir_master) . $file_xlsx);

$Import->setMap([
	'A' => 'gop_number',
	'B' => 'discount_taken',
	'C'	=> 'gop_verification_remark',
	'D'	=> 'edited_by'
]);

$Export = Export::open(path($dir_result) . 'Result_' . $file_xlsx);

$Export->setMap([
	'A' => 'gop_number',
	'B' => 'discount_taken',
	'C'	=> 'gop_verification_remark',
	'D'	=> 'edited_by',
	'F'	=> 'result',
	'G'	=> 'reason'
]);

$Export->setImport($Import);

$case_log 		= [];

$logs 			= [];

$logs_detail 	= [];

$Import->each(function($row, $index) use (&$case_log, &$logs, &$logs_detail, &$Export) {

	if($row['gop_number'] === '') {
		return;
	}

	$row['gop_number'] 			= trim($row['gop_number']);

	$row['gop_verification'] 	= 1;
	$row['edit_date']  			= date('Y-m-d');

	$status 	= 'success';
	$reason  	= '';
	
	//---------------------------------------------------
	if(!in_array($row['discount_taken'], [1,2])) {
		$Export->on('result', $index)->setText('error');
		$Export->on('reason', $index)->setText($reason = 'Discount Taken must be number format. value is beetween 1 or 2');
		$status 	= 'error';
	}else {
		$Export->on('result', $index)->setText('success');
		$status 	= 'success';
		$case_log[] = $row;
	}
	//---------------------------------------------------

	// create logs
	//--------------------------------------------------------
	$logs[]		= [
		'code' 			=> $code = uniquuid(),
		'log_number'	=> $row['gop_number'],
		'create_date'	=> date('Y-m-d H:i:s'),
		'created_by'	=> $row['edited_by'],
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
		'attribute'		=> 'discount_taken',
		'value'			=> $row['discount_taken']
	];

}, TRUE);

print_r($case_log);

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
	$query = Database::on('ito_beta')->builder();

	try {

		$query
			->into('case_log')
			->update($case_log);

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