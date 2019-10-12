<?php 
namespace Pusaka\Utils;

use closure;

class FileUtils {

	const FORM 	= 'FORM';
	const FILE 	= 'FILE';
	const URL 	= 'URL';

	const JSON 	= 'JSON';
	const TEXT 	= 'TEXT';
	const BYTE 	= 'BYTE';

	// inaccessable to set
	//------------------------
	private $info;
	private $category;
	private $config;
	private $src;
	private $response;
	private $multiple;
	private $upload;
	private $skip_error;
	private $contents;

	// accessable to set
	//------------------------
	private $error;
	private $location;
	private $link;

	function __construct($src) {

		$this->category 	= FileUtils::FILE;

		$this->src 			= $src;

		$this->config 		= config('upload');

		$this->response 	= NULL;

		$this->error 		= NULL;

		$this->multiple 	= FALSE;

		$this->upload 		= NULL;

		$this->skip_error 	= FALSE;

	}

	function __set($property, $value) {

		if(in_array($property, 'error', 'location', 'link')){
			$this->{$property} = $value;
		}
	
	}

	function __get($property) {

		return $this->{$property};

	}

	function multiple() {

		$this->multiple 	= TRUE; 

	}

	function category( $category = NULL ) {

		if($category === NULL) {
			return $this->category;
		}

		$this->category = $category;

	}

	function mime() {

		if($this->category == FileUtils::FILE) {
			return mime_content_type($this->src);
		}

		return NULL;

	}

	function name() {

		if($this->category == FileUtils::FILE) {
			return basename($this->src);
		}

		return NULL;

	}

	function ext() {

		if($this->category == FileUtils::FILE) {
			return strtolower(pathinfo($this->src, PATHINFO_EXTENSION));
		}

		return NULL;

	}

	function info() {

		return $this->info;
	
	}

	function src() {

		return $this->src;
	
	}

	function copy($to = NULL) {

		if($this->category !== FileUtils::FILE) {
			$this->error 	= 'Category '.$this->category.' not allowed to copy file.';
			$this->response = 'Copy failed. cause : '.$this->error;
			return $this;
		}

		if(!file_exists($this->src)) {
			$this->error 	= 'File not found. ('.$this->src.')' ;
			$this->response = 'Copy failed. cause : '.$this->error;
			return $this;
		}

		if(!is_dir($to)) {
			$this->error 	= 'Destination not found. ('.$to.')' ;
			$this->response = 'Copy failed. cause : '.$this->error;	
			return $this;
		}

		$dest = path($to) . $this->name();
	
		if(!copy($this->src, $dest)){
			$this->error 	= 'Copy return false.';
			$this->response = 'Copy failed. cause : '.$this->error;	
			return $this;
		}

		return $this;
	}

	function move($to = NULL) {

		if($this->category !== FileUtils::FILE) {
			$this->error 	= 'Category '.$this->category.' not allowed to move file.';
			$this->response = 'Move failed. cause : '.$this->error;
			return $this;
		}

		if(!file_exists($this->src)) {
			$this->error 	= 'File not found. ('.$this->src.')' ;
			$this->response = 'Move failed. cause : '.$this->error;
			return $this;
		}

		if(!is_dir($to)) {
			$this->error 	= 'Destination not found. ('.$to.')' ;
			$this->response = 'Move failed. cause : '.$this->error;	
			return $this;
		}

		$dest = path($to) . $this->name();
	
		if(!rename($this->src, $dest)){
			$this->error 	= 'Copy return false.';
			$this->response = 'Move failed. cause : '.$this->error;	
			return $this;
		}

		return $this;

	}

	function delete() {

		if($this->category 	=== FileUtils::FILE) {
		
			if(!file_exists($this->src)) {
				$this->error 	= 'File not found. ('.$this->src.')' ;
				$this->response = 'Delete failed. cause : '.$this->error;
				return $this;
			}

			if(!unlink($this->src)){
				$this->error 	= 'Unlink return false.';
				$this->response = 'Delete failed. cause : '.$this->error;	
				return $this;
			}

			return $this;

		}

		if($this->category 	=== FileUtils::FORM) {
		
			if($this->upload === NULL) {
				$this->error 	= 'Upload file not found.';
				$response 		= 'Delete failed. cause : '.$this->error;
				return $this;
			}

			if($this->multiple) {

				if(!is_array($this->upload)) {
					$this->error 	= 'Multiple upload file invalid format.';
					$this->response = 'Delete failed. cause : '.$this->error;
					return $this;		
				}else {

					foreach ($this->upload as $key => $File) {
						$File->delete();
					}

				}

			}else {
				
				if($this->upload === NULL) {
					$this->error 	= 'Upload is NULL.';
					$response 		= 'Delete failed. cause : '.$this->error;
					return $this;
				}else {
					return $this->upload->delete();
				}

			}

			return $this;			

		}

		$this->error 	= 'Category '.$this->category.' not allowed to delete file.';
		$this->response = 'Delete failed. cause : '.$this->error;
		return $this;

		return $this;

	}

	function write($contents = '', $format = 'text') {

		if(is_array($contents)) {
			$contents = json_encode($contents);
		}

		return file_put_contents($this->src, $contents);
	}

	function contents($mime = self::TEXT) {

		$text = '';

		if(file_exists($this->src)) {
			$text = file_get_contents($this->src);
		};

		if($mime === self::JSON) {
			$text = json_decode($text, true);
		}

		return $text;		

	}

	function replace(array $replace) {

		if($this->category === self::FILE) {
			
			$contents = $this->contents();

			$this->contents = strtr($contents, $replace);

		}

		return $this;

	}

	function save($file) {

		return file_put_contents($file, $this->contents);

	}

	function read($closure) {

		$text = '';

		if(is_file($this->src)) {
			$text = file_get_contents($this->src);
		};
		

		if($closure instanceof closure) {
			$closure($text);
		}

	}

	function config($config = NULL) {

		if(is_array($config)) {

			foreach ($config as $key => $value) {
				if(isset($this->config[$key])){
					$this->config[$key] = $value;
				}
			}

		}

		return $this;

	}

	function upload($idx = NULL) {

		$file 		= NULL;

		if($this->category !== FileUtils::FORM) {
			$this->error 	= "Not a form File.";
			$this->response = "Upload failed. cause : ".$this->error;
			return $this;
		}

		$allowed 	= explode('|', strtr($this->config['ext'], [' '=>'']) );

		if($this->multiple) {

			if(!is_int($idx)) {
				$this->error 	= "No file upload.";
				$this->response = "Upload failed. cause : ".$this->error;
				return $this;
			}

			$file 			= [
				'name' 		=> $this->src['name'][$idx] ?: '',
				'type' 		=> $this->src['type'][$idx] ?: '',
				'tmp_name' 	=> $this->src['tmp_name'][$idx] ?: '',
				'error' 	=> $this->src['error'][$idx] ?: '',
				'size' 		=> $this->src['size'][$idx] ?: ''
			];

			$file['ext'] 	= pathinfo($file['name'], PATHINFO_EXTENSION);

		}else if(isset($this->src['name'])) {

			if(is_array($this->src['name'])) {
				$this->multiple();
				$this->skip_error 	= TRUE; 	 		
				$this->upload 		= [];
				$this->response 	= [];

				// edit 2019-08-26
				//------------------------
				foreach($this->src['name'] as $i => $val) {
					$this->upload($i);
				}

				// depreciated
				//------------------------
				// for($i=0, $c=count(); $i<$c; $i++) {
				// 	$this->upload($i);
				// }
				return $this;
			}

			$file 			= $this->src;
			$file['ext']  	= pathinfo($file['name'], PATHINFO_EXTENSION);
		}

		$FileUtils = new FileUtils($file);

		/*
		| Empty name => no file upload
		|-------------------------------------- */
		if($file['name'] === '') {
			
			$this->error 		= "No file upload.";
			$response 			= "Upload failed. cause : ".$this->error;

			$FileUtils->error 	= $this->error;

			if($this->multiple) {
				$this->response[$idx] 	= $FileUtils;
			}else {
				$this->response 		= $FileUtils;
			}

			return $this;
		}

		/*
		| File is NULL
		|-------------------------------------- */
		if($file === NULL) {

			$this->error 		= "File not found.";
			$response 			= "Upload failed. cause : ".$this->error;

			$FileUtils->error 	= $this->error;

			if($this->multiple) {
				$this->response[$idx] 	= $FileUtils;
			}else {
				$this->response 		= $FileUtils;
			}

			return $this;
		}

		$tmp 	= $file['tmp_name'];
 		$upload = path($this->config['save']);

 		/*
		| Directory for save upload file not found
		|-------------------------------------- */
 		if(!is_dir($upload)) {
 			
 			$this->error 		= "Folder not found. in $upload";
 			$response 			= "Upload failed. cause : ".$this->error;

 			$FileUtils->error 	= $this->error;

			if($this->multiple) {
				$this->response[$idx] 	= $FileUtils;
			}else {
				$this->response 		= $FileUtils;
			}

 			return $this;
 		}

 		/*
		| Extention not allowed
		|-------------------------------------- */
 		if(!in_array($file['ext'], $allowed)) {
 			
 			$this->error 		= "File extension not allowed.";
 			$response 			= "Upload failed. cause : ".$this->error;

 			$FileUtils->error 	= $this->error;

			if($this->multiple) {
				$this->response[$idx] 	= $FileUtils;
			}else {
				$this->response 		= $FileUtils;
			}

 			return $this;
 		}

 		/*
		| Oversize
		|-------------------------------------- */
 		if($file['size'] > ByteUtils::value($this->config['max']) ) {
			
			$this->error 		= 'Max filesize is '.ByteUtils::string($this->config['max']). '. ('.ByteUtils::string($file['size']).') ';
			$response 			= "Upload failed. cause : ".$this->error;

			$FileUtils->error 	= $this->error;

			if($this->multiple) {
				$this->response[$idx] 	= $FileUtils;
			}else {
				$this->response 		= $FileUtils;
			}

			return $this;
		}

		$letsEncrypt = function() use ($file) {
			return strtolower(date('ymdHi').uniqid()) . '.' . $file['ext'];
		};

		$letsClear  = function() use ($file) {
			$name = $file['name'];
			$name = strtr($name, $this->config['remove'], '');
			return $name;
		};

		if($this->config['encrypt']) {
 			$file['name'] = $letsEncrypt();
 		}else {
 			$file['name'] = $letsClear();
 		}

 		$saveas = $upload.$file['name'];

 		if(file_exists($saveas)) {
 			$file['name'] 	= $letsEncrypt();
			$saveas 		= $upload.$file['name'];
 		}

 		/*
		| Error when upload
		|-------------------------------------- */
 		if(!move_uploaded_file($tmp, $saveas)) {
			
			$this->error 		= 'Move upload file return false.';
			$response 			= "Upload failed. cause : ".$this->error;

			$FileUtils->error 	= $this->error;

			if($this->multiple) {
				$this->response[$idx] 	= $FileUtils;
			}else {
				$this->response 		= $FileUtils;
			}

			return $this;
		}

		/*
		| Success state
		|-------------------------------------- */

		$link 		= strtr($this->config['link'], ['{filename}' => ($file['name'])]);
		$location 	= $saveas;

		$FileUtils 				= new FileUtils($location);
		$FileUtils->link 		= $link;
		$FileUtils->location 	= $location;

		if($this->multiple) {
			if(is_array($this->response)) {

				$this->upload[] 		= $FileUtils;
				$this->response[] 		= $FileUtils;
			
			}
		}else {
			$this->upload 	  = $FileUtils;
			$this->response   = $FileUtils;
		}

		return $this;
	}

	function success($closure = NULL) {

		if($closure === NULL AND $this->error !== NULL) {
			return FALSE;
		}

		if($closure === NULL AND $this->error === NULL) {
			return TRUE;
		}

		if($this->error !== NULL AND $this->skip_error === FALSE) {
			return $this;
		}

		if($closure instanceof closure) {
			$closure($this->response);
		}

		return $this;

	}

	function error($closure = NULL) {
		
		if($closure === NULL AND $this->error === NULL) {
			return FALSE;
		}

		if($closure === NULL AND $this->error !== NULL) {
			return TRUE;
		}

		if($this->error === NULL AND $this->skip_error !== FALSE) {
			return $this;
		}

		if($closure instanceof closure) {
			$closure($this->response);
		}

		return $this;

	}

	function deepSearch($do, $ignore = [], $pattern = NULL, $recursive = NULL) {

		$dir = is_null($recursive) ? $this->src : $recursive;

		if(is_dir($dir)) {

			$list = scandir($dir);
		
			foreach ($list as $file) {

				if($file === '.' OR $file === '..') {
					continue;
				}

				$fullpath = path($dir) . $file;

				if(is_dir($fullpath)) {
					if(!in_array(basename($fullpath), $ignore) ) {
						$this->deepSearch($do, $ignore, $pattern, $fullpath);
					}
				}
				else if(is_file($fullpath)) {

					if($pattern === NULL) {
						$do($fullpath);
					}else if(preg_match($pattern, $file)) {
						$do($fullpath);
					}

				}

			}

		}

	}

}