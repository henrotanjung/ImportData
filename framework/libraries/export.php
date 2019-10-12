<?php 
namespace Pusaka\Library;

use Exception;
use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Cell;
use stdClass;

class Export {
	
	const TYPE_STRING2  = 'str';
	const TYPE_STRING   = 's';
	const TYPE_FORMULA  = 'f';
	const TYPE_NUMERIC  = 'n';
	const TYPE_BOOL     = 'b';
	const TYPE_NULL     = 'null';
	const TYPE_INLINE   = 'inlineStr'; // Rich text
	const TYPE_ERROR    = 'e';

	private $file;
	private $import;

	private $excel;
	private $sheet;
	private $active_cell;

	private $map;

	static function open($file) {

		$file = strtr($file, ['\\' => '/']);

		if(!is_dir(dirname($file))) {
			throw new Exception("$fie not found.");
		}

		return new Export($file);

	}

	function __construct($file) {

		$this->file  = $file;
		// $this->excel = $object;
		// $this->sheet = $object->getSheet(0);

	}

	function setImport($import) {
		
		$this->import = $import;

	}

	function setMap($map) {
		
		$this->map = $map;

	}

	function on($attr, $index) {
	
		$column = array_search($attr, $this->map);
		
		$this->active_cell = $column.$index;
		return $this;

	}

	function setText($text) {

		$this->import->sheet->setCellValue($this->active_cell, $text);

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

			$funct($row);

			unset($row);

		}

	}

	function save() {

		try {

			$objWriter = PHPExcel_IOFactory::createWriter($this->import->excel, $this->import->type);
			$objWriter->save($this->file);

			return true;
		
		}catch(Exception $e) {

			return false;

		}


	}



}
