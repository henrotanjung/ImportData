<?php 
namespace Pusaka\Utils;

use closure;

class CryptoUtils {

	public static function token($length, $idxset) {

		$charset[0] = '0123456789';
		$charset[1] = 'abcdefghijklmnopqrstuvwxyz';
		$charset[2] = $charset[0].$charset[1];
		$charset[3] = $charset[2].'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

		$pool 		= isset($charset[$idxset]) ? $charset[$idxset] : NULL;
		$plen 	 	= strlen($pool) - 1;

		if($pool == NULL) {
			return NULL;
		}

		$token 		= '';

		for($i=0; $i<$length; $i++) {
			$char 	= substr($pool, rand(0, $plen), 1);
			$token .= $char; 
		}

		return $token;

	}

	

}