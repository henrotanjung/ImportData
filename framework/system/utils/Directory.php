<?php 
namespace Pusaka\Utils;

use closure;

class DirectoryUtils {

	private $src;

	public function __construct($src) {

		$this->src = $src;

	}

	public function __set($property, $value) {

		if(in_array($property, 'error')){
			$this->{$property} = $value;
		}
	
	}

	public function __get($property) {

		return $this->{$property};

	}

	private function __recursiveDelete($path) {

		$del = $files = glob($path . '/*');
		
		foreach ($files as $file) {
			is_dir($file) ? $this->__recursiveDelete($file) : unlink($file);
		}

		rmdir($path);

		return !is_dir($path);

	}

	public function make() {

		if(is_dir($this->src)) {
			return true;
		}

		return mkdir($this->src, true);

	}

	public function delete() {

		return $this->__recursiveDelete($this->src);

	}

}