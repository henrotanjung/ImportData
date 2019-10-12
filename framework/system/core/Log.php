<?php 
namespace Pusaka\Microframework;

class Log {

	static function create($module, $file, $message, $backtrace) {

		$log_dir 	= path($GLOBALS['config']['app']['log']);

		$log_file 	= $log_dir . 'errorlog_' . date('ymdhis') . uniqid() . '.txt';

		$log 		= 
			json_encode(
				[
					"script"	=> $file,
					"module"	=> $module,
					"message" 	=> $message
				],
				JSON_PRETTY_PRINT
			) 
			. "\r\n"
			. "\r\n" 
			. "MESSAGE : " . $message
			. "\r\n"
			. "\r\n"
			.
			json_encode(
				$backtrace,
				JSON_PRETTY_PRINT
			);
			

		file_put_contents($log_file, $log);

	}

}