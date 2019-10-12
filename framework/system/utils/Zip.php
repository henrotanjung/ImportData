<?php 
namespace Pusaka\Utils;

use Pusaka\Utils\FileUtils;

use closure;
use ZipArchive;

class ZipUtils {

	private $zip_archive;

	private $zip_file;

	/**
     * Array of well known zip status codes
     *
     * @var array
     */
    private static $zip_status_codes = Array(
        ZipArchive::ER_OK          => 'No error',
        ZipArchive::ER_MULTIDISK   => 'Multi-disk zip archives not supported',
        ZipArchive::ER_RENAME      => 'Renaming temporary file failed',
        ZipArchive::ER_CLOSE       => 'Closing zip archive failed',
        ZipArchive::ER_SEEK        => 'Seek error',
        ZipArchive::ER_READ        => 'Read error',
        ZipArchive::ER_WRITE       => 'Write error',
        ZipArchive::ER_CRC         => 'CRC error',
        ZipArchive::ER_ZIPCLOSED   => 'Containing zip archive was closed',
        ZipArchive::ER_NOENT       => 'No such file',
        ZipArchive::ER_EXISTS      => 'File already exists',
        ZipArchive::ER_OPEN        => 'Can\'t open file',
        ZipArchive::ER_TMPOPEN     => 'Failure to create temporary file',
        ZipArchive::ER_ZLIB        => 'Zlib error',
        ZipArchive::ER_MEMORY      => 'Malloc failure',
        ZipArchive::ER_CHANGED     => 'Entry has been changed',
        ZipArchive::ER_COMPNOTSUPP => 'Compression method not supported',
        ZipArchive::ER_EOF         => 'Premature EOF',
        ZipArchive::ER_INVAL       => 'Invalid argument',
        ZipArchive::ER_NOZIP       => 'Not a zip archive',
        ZipArchive::ER_INTERNAL    => 'Internal error',
        ZipArchive::ER_INCONS      => 'Zip archive inconsistent',
        ZipArchive::ER_REMOVE      => 'Can\'t remove file',
        ZipArchive::ER_DELETED     => 'Entry has been deleted'
    );

    /**
     * Class constructor
     *
     * @param   string $zip_file ZIP file name
     *
     */
    public function __construct($zip_file) {
		
        if (empty($zip_file)) {
            throw new \Exception(self::getStatus(ZipArchive::ER_NOENT));
        }
        
        $this->zip_file 	= $zip_file;

        $this->zip_archive 	= new ZipArchive();
    
    }

    public static function create($zip_file) {

    	$ZipUtils = new ZipUtils($zip_file);

    	$ZipUtils->__open(ZipArchive::CREATE|ZipArchive::OVERWRITE);

    	return $ZipUtils;

    }

    public static function open($zip_file) {

    	$ZipUtils = new ZipUtils($zip_file);

    	$ZipUtils->__open(NULL);

    	return $ZipUtils;

    }

    /**
     * Open zip archive (ZipUtils)
     *
     * @return  ZipUtils
     */
    private function __open($options) {

    	$this->zip_archive->open($this->zip_file, $options);

    	return $this;

    }

    public function add($to, $from = NULL) {

    	if(is_string($from)) {

    		$this->zip_archive->addFromString($to, $from);

    	}else 

    	if($to instanceof DirectoryUtils) {

			$this->zip_archive->addEmptyDir($to->src);    		

    	}else 

    	if($from instanceof FileUtils) {

    		$this->zip_archive->addFile($from->src, $to);

    		unset($from);

    	}

    }

    public function unzip($path) {

    	if(is_dir($path)) {
    		$this->zip_archive->extractTo($path);
    	}else {
    		throw new Exception(self::getStatus("directory not exits."));
    	}

    }

    /**
     * Close zip archive (void)
     *
     * @return  void
     */
    public function close() {

    	if($this->zip_archive !== NULL) {
    		$this->zip_archive->close();
    	}

    }

    /**
     * Get status of archive (string)
     *
     * @return  string
     */
    public static function getStatus($type) {

    	return self::$zip_status_codes[$type] ?? '';

    }

    /**
     * Get a list of files in archive (array)
     *
     * @return  array
     */
    public function listFiles() {
        
        $list = [];
        
        for ($i = 0; $i < $this->zip_archive->numFiles; $i++) {
            $name = $this->zip_archive->getNameIndex($i);
            if ($name === false) {
                throw new Exception(self::getStatus($this->zip_archive->status));
            }
            $list[] = $name;
        }

        return $list;
    }

    /**
     * Destructor freeing memory and close the zip archive
     *
     * @return  void
     */
    public function __destruct() {

    	$this->close();

    }


}