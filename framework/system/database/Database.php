<?php 
namespace Pusaka\Database;

use Pusaka\Database\DatabaseException;

require_once( strtr(__DIR__, '\\', '/') . '/Exception.php');

class Manager {

	public static function on($config) {

		$name 	= $config;

		$config = $GLOBALS['config']['database'][$config] ?: NULL;

		if($config == NULL) {
			throw new DatabaseException("Configuration database[$name] not found", DatabaseException::CONFIG_ERROR);
			return NULL;
		}

		if($config['driver'] == 'mysql') {
			return new \Pusaka\Database\Mysql\Driver($config);
		}

		return NULL;

	}

}