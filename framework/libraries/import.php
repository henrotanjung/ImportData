<?php 
namespace Pusaka\Library;

use Exception;
use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Cell;
use stdClass;

class Import {
	
	const TYPE_STRING2  = 'str';
	const TYPE_STRING   = 's';
	const TYPE_FORMULA  = 'f';
	const TYPE_NUMERIC  = 'n';
	const TYPE_BOOL     = 'b';
	const TYPE_NULL     = 'null';
	const TYPE_INLINE   = 'inlineStr'; // Rich text
	const TYPE_ERROR    = 'e';

	private $type;

	private $excel;
	private $sheet;

	private $map;

	static function open($file) {

		$file = strtr($file, ['\\' => '/']);

		if(!file_exists($file)) {
			throw new Exception("$file not found.");
		}

		return new Import($file);

	}

	function __construct($file) {

		try {

			$type 			= PHPExcel_IOFactory::identify($file);
			$reader 		= PHPExcel_IOFactory::createReader($type);
			$object 		= $reader->load($file);

			$this->type 	= $type;

		}catch(Exception $e) {
			throw new Exception("Error loading file on PHPExcel.", 1);
		}

		$this->excel = $object;
		$this->sheet = $object->getSheet(0);

	}

	function __get($attr) {

		return $this->{$attr};

	}

	function setMap($map) {	
		// var_dump($map);		
		$this->map = $map;
	}

	function setMapAuto($map) {
		$setmap = [];
		foreach($map as $key => $val){
			$colString = PHPExcel_Cell::stringFromColumnIndex($key);
			var_dump($colString);
			$setmap[$colString] = $val;
		}
		// var_dump($setmap);
		$this->setMap($setmap);
		
	}

	function each($funct, $array = FALSE) {

		$count = $this->sheet->getHighestRow();

		for ($i=2; $i<=$count; $i++) {
			
			if($array === FALSE) {
				$row = new stdClass;
			}else {
				$row = [];
			}

			foreach ($this->map as $key => $as) {

				$col 	= PHPExcel_Cell::columnIndexFromString($key) - 1;
				$cell 	= $this->sheet->getCellByColumnAndRow($col, $i);

				if($array === FALSE) {
					$row->{$as} = $cell->getFormattedValue();
				}else {
					$row[$as] = $cell->getFormattedValue();
				}

			}

			$funct($row, $i);

			unset($row);

		}

	}



}
