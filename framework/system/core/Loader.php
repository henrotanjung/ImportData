<?php 
namespace Pusaka\Microframework;

class Loader {

	static function lib($name) {

		$file = FRAMEWORK_DIR . 'libraries/' . $name . '.php';

		if(file_exists($file)) {
			require_once($file);
		}else {
			echo "file not found : " . ($file);
		}

	}

}