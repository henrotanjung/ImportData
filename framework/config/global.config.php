<?php 
/*
| APP : Config
|=============================== */
$config['app']['url'] 		= URL;
$config['app']['security'] 	= 'roU2c195jzTQbr72H25LeW0BvpsEtFwi';
$config['app']['csrf'] 		= FALSE;
$config['app']['storage'] 	= ROOTDIR . 'storage';
$config['app']['log'] 		= LOG;
$config['app']['evpath'] 	= ROOTDIR . '';

/*
| API : Config
|=============================== */
$config['api']['url']		= URL . 'restapi';
$config['api']['key']		= '<<api_key>>';


/*
| Upload
|=============================== */
$config['upload']['link']		= ':rootpath:/storage/{filename}'; 
$config['upload']['save']		= ROOTDIR . 'storage/';
$config['upload']['max'] 		= '100 KB';
$config['upload']['ext'] 		= 'jpg|jpeg|png|gif';
$config['upload']['encrypt'] 	= FALSE;
$config['upload']['remove']		= '#';


/*
| Load utils class
|------------------------------- */
$config['app']['utils']		= [
	'Array', 'Byte', 'File', 'Directory', 'IO', 'Iterator'
];