<?php 
namespace Pusaka\Database;

use Exception;

class DatabaseException extends Exception {

	const NONE 				= 0;
	const CONFIG_ERROR 		= 1;
	const CONNECTION_ERROR 	= 2;
	const INVALID_PARAMETER = 3;
	const QUERY_ERROR 		= 4;

	function __construct($message = null, $code = 0, $previous = null) {

		$this->message 	= $message;

		$this->$code 	= $code;

	}

	function getCodeDescription() {

		switch ($this->code) {
			case DatabaseException::NONE:
				return 'None';
				break;
			case DatabaseException::CONFIG_ERROR:
				return 'Configuration Error';
				break;
			case DatabaseException::CONNECTION_ERROR:
				return 'Connection Error';
				break;
			case DatabaseException::INVALID_PARAMETER:
				return 'Invalid Parameter';
				break;
			case DatabaseException::QUERY_ERROR:
				return 'Query Error';
				break;
		}

	}

	public function __toString() {
        return __CLASS__ . ": [{$this->getCodeDescription()}]: {$this->message}\n";
    }

}