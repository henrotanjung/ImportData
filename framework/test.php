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
	'B' => 'alamat',
	'C'	=> 'telepon'	
]);

$Import->each(function($row, $index) {
  var_dump($row["nama"]);
  
}, true);
var_dump($Import);
$query = "insert into data_pegawai values('henro','jaksel','898e')";


/*
$query = Database::on('dblocal')->builder();
$datas = $query->select('*')->from('data_pegawai');
$datas = $query->get();

$datas = json_decode(json_encode($datas), true);

// var_dump($datas);
foreach ($datas as $key => $data){
	unset($datas[$key]["id"]);
	
	// var_dump($res);
	// echo $data->nama;
}
var_dump($datas);


$res = $query->into('data_pegawai')->insert($datas);


// [
// 	"nama" => $data->nama,
// 	"alamat" => $data->alamat,
// 	"telepon" =>$data->telepon
// ]
