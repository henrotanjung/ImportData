<?php 
namespace App\Config;

class Route {
	
	static function register() {

		return [
			':default' 			=> 'main/home/',
			':notfound'			=> 'notfound',
			'tutorial' 			=> 'main/tutorial/',
			'tutorial/(:any)'	=> 'main/tutorial/$1',
			'docs' 				=> 'main/docs/',
			'docs/(:any)'		=> 'main/docs/$1',
			'debug'				=> 'main/debug/',
			'debug/(:any)'		=> 'main/debug/$1',
			'blog/(:num)'		=> 'main/blog/$1',
			'blog/(:any)'		=> 'main/blog/$1',
			'signin'			=> 'admin/login',
			'signin/(:any)'		=> 'admin/login/$1',
			'integrated-application-protocol' => 'main/blog/integrated-application-protocol',
			'working-with-cli' 				  => 'main/blog/working-with-cli',
		];

	}

}