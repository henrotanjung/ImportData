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
  
$Import->setMap(
    $column
); 
   
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
