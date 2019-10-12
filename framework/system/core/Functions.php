<?php 

if(!function_exists('env')) {

	function env($key, $val) {
		define($key, $val);
		return $val;
	}

}

if(!function_exists('foridx')) {

	function foridx($object, $do) {

		for($i=0, $c=count($object); $i<$c; $i++) {
			$do($object[$i]);
		}

	}

}

if(!function_exists('url')) {

	function url($additional = '') {
		return BASEURL . $additional;
	}

}

if(!function_exists('str_replace_first')) {

	function str_replace_first($from, $to, $content) {
		$from = '/'.preg_quote($from, '/').'/';

		return preg_replace($from, $to, $content, 1);
	}

}

if(!function_exists('is_assoc')) {
	function is_assoc($arr)
	{
		return array_keys($arr) !== range(0, count($arr) - 1);
	}
}

if(!function_exists('config')) {
	function config($keys)
	{
		return $GLOBALS['config'][$keys] ?: NULL;
	}
}

if(!function_exists('path')) {
	function path($path) {

		$path = strtr($path, ['\\' => '/']);

		return rtrim($path, '/') . '/';

	}
}

if(!function_exists('uniquuid')) {
	function uniquuid() {
		return date('Ymdhis').uniqid();
	}
}