<?php 
namespace Pusaka\Utils;

class ArrayUtils {

	static function merge(&$var, ...$merge) {

		foreach ($merge as $value) {
			$var = array_merge($var, $value);
		}

	}

}