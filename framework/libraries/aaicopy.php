<?php
namespace Pusaka\Library;

use Pusaka\Microframework\Log;

class AAICopy {

	private $module;
	private $script;

	private $from;
	private $to;
	private $log;

	function __construct($from, $to) {

		$this->from = $from;
		$this->to 	= $to;

	}

	function __set($attr, $value) {

		$this->{$attr} = $value;

	}

	function success() {


		if(!file_exists($this->from)) {
			
			$this->log = "File source not found (" . $this->from . ")";
			
			Log::create($this->module, $this->script, $this->log, debug_backtrace());
			
			return false;

		}

		if(!is_dir(dirname($this->to))) {

			$this->log = "Destination directory not found. (" . dirname($this->to) . ")";

			Log::create($this->module, $this->script, $this->log, debug_backtrace());
			
			return false;

		}

		if(!copy($this->from, $this->to)) {

			$this->log = "Copy failed unknown reason.";

			Log::create($this->module, $this->script, $this->log, debug_backtrace());
			
			return false;

		}

		return true;

	}

	function log() {

		return $this->log;

	}

}