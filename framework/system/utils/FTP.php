<?php 
namespace Pusaka\Utils;

use closure;

class FTPUtils {

	private $failed 	= [];

	private $workdir 	= '/';

	private $conn 		= NULL;
	private $host 		= '';
	private $user 		= '';
	private $pass 		= '';

	function __construct($host, $user, $pass) {
		$this->host = $host;
		$this->user = $user;
		$this->pass = $pass;
	}

	public static function make($path) {

		$path = strtr($path, ['\\' => '/']);

		$path = preg_replace('/(\.\.\/)|(\.\/)|(^\.)/', '/', $path);

		return $path;

	}

	function connect() {
		
		$this->conn = ftp_connect($this->host);
		
		if(!$this->conn) {
			throw new Exception('Cannot connect ftp server.');
		}
		
		$is_login 	= ftp_login($this->conn, $this->user, $this->pass);
		
		if(!$is_login) {
			ftp_close($this->conn);
			throw new Exception('Cannot login to ftp server. Close Connection');
		}else {
			ftp_pasv($this->conn, true);
		}

		return TRUE;
	}

	function close() {
		
		if($this->conn !== NULL) {
			if(!is_integer($this->conn)) {
				ftp_close($this->conn);
			}
		}

	}

	function setWorkdir($workdir) {
		$this->workdir = $workdir;
	}

	function mkdir($ftp_basedir, $ftp_path) {

		@ftp_chdir($this->conn, $ftp_basedir); // /var/www/uploads
			
		$parts = array_filter(explode('/',$ftp_path)); // 2013/06/11/username
		
		foreach($parts as $part){

			if(!@ftp_chdir($this->conn, $part)) {
			
				ftp_mkdir($this->conn, $part);
				//ftp_chmod($ftpcon, 0775, $part);
				ftp_chdir($this->conn, $part);
			
			}

		}

	}

	function found($find) {

		$list = ftp_nlist( $this->conn, '.' );

		if(!$list) {
			return [];
		}

		return in_array($find, $list);

	}

	function ls($do = NULL) {

		$list = ftp_nlist( $this->conn, $this->workdir );

		if($do !== NULL) {

			foreach ($list as $file) {
				$do( $file );
			}

		}

	}

	function upload($from, $to) {

		$path 	= strtr(pathinfo($to)['dirname'], '\\', '/');

		$to = $this->workdir . $to;

		$to = "/" . ltrim($to, "/");

		// create folder
		if($path !== '/') {

			$this->mkdir($this->workdir,$path);
		
		}

		$d = ftp_nb_put($this->conn, $to, $from, FTP_BINARY);

		while ($d == FTP_MOREDATA) {

			$d = ftp_nb_continue($this->conn);
		
		}

		if ($d != FTP_FINISHED) {

			$this->failed[] = $from;

			return false;

		}else {

			return true;
			
		}

	}

	function __destruct() {
		
		$this->close();
		unset($this->conn);

	}

}