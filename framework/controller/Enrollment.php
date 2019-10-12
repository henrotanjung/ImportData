<?php
use Pusaka\Library\Import;
use Pusaka\Microframework\Loader;
use Pusaka\Database\Manager as Database;
Loader::lib('import');

    class Enrollment{

        public function insert($fileName,$column){
            $folder = '//192.168.0.20/Z Folder/IT/INTERNAL ONLY/003 Internal Minutes of Meeting/Programmer Team/005 Hendro/';
            // $folder = ROOTDIR;
            // $file_xlsx = 'addition_20191007.xlsx';
            $file_xlsx = $fileName;

            echo "<pre>";

            $Import = Import::open(path($folder) . $file_xlsx);
            
            $Import->setMap($column);
            
            $query = Database::on('dblocal_dev_aia')->builder();
            $datas = array();

            $Import->each(function($row, $index) use (&$datas) {
                if ($row['client'] === ''){
                    return;
                }    
                $datas[] = $row;		
            }, true);
            // var_dump($datas);

            $res_sql = $query->into('member')->insert($datas);
            // echo $res_sql;

    }
    }