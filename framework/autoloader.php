<?php
define('FRAMEWORK_DIR', str_replace('\\', '/', __DIR__ . '/'));

// load config
include(FRAMEWORK_DIR . 'config/microframework.config.php');
include(FRAMEWORK_DIR . 'config/database.config.php');
include(FRAMEWORK_DIR . 'config/path.config.php');

// load utils
include(FRAMEWORK_DIR . 'system/core/Functions.php');

// load utils
include(FRAMEWORK_DIR . 'system/utils/IO.php');
include(FRAMEWORK_DIR . 'system/utils/File.php');
include(FRAMEWORK_DIR . 'system/utils/Directory.php');

// load database
include(FRAMEWORK_DIR . 'system/database/driver/mysql/driver.php');
include(FRAMEWORK_DIR . 'system/database/driver/mysql/builder.php');
include(FRAMEWORK_DIR . 'system/database/driver/mysql/result.php');

include(FRAMEWORK_DIR . 'system/database/Exception.php');
include(FRAMEWORK_DIR . 'system/database/Database.php');

// load core
include(FRAMEWORK_DIR . 'system/core/Loader.php');
include(FRAMEWORK_DIR . 'system/core/Log.php');

// loader library
require_once(ROOTDIR . 'libraries/PHPExcel.php');
require_once(ROOTDIR . 'libraries/PHPExcel/Writer/Excel2007.php');
require_once(ROOTDIR . 'libraries/PHPExcel/IOFactory.php');


