<?php
use Pusaka\Database\Manager as Database;
use Pusaka\Database\DatabaseException;
use Pusaka\Microframework\Loader;
use Pusaka\Microframework\Log;
use Pusaka\Library\Import;
use Pusaka\Library\Export;
use Pusaka\Library\AAICopy;


Loader::lib('import');

    class Enrollment{     
        
        public function insert($fileName,$column,$client,$company){
            $folder = '//192.168.0.20/Z Folder/IT/INTERNAL ONLY/003 Internal Minutes of Meeting/Programmer Team/005 Hendro/';
            // $folder = ROOTDIR;
            // $file_xlsx = 'addition_20191007.xlsx';
            $file_xlsx = $fileName;

            // echo "<pre>";

            $Import = Import::open(path($folder) . $file_xlsx);
            
            $Import->setMap($column);
            
            $query = Database::on('dblocal_dev_aia')->builder();
            $datas = array();

            // include '../config/koneksi.php';
            include 'Model.php';       
            include 'adding_data.php';     
            $model_obj = new Model();    
            $add_data_obj = new AddingDatas(); 
            $clientDatas = $add_data_obj->client($client);
            // $compDatas = $add_data_obj->company($client);            
            $Import->each(function($row, $index) use (&$client,&$datas,&$clientDatas,&$add_data_obj) {                               
                include 'cleaning_data.php';      
                if ($row['client'] === ''){
                    return;
                }                  
                $planDatas = $add_data_obj->plan($row['plan']);
                $row['program'] = $planDatas['program'];
                $row['client'] = $clientDatas['id'];                
                $row['plan'] = $planDatas['id'];
                $compDatas = $add_data_obj->company($client,$planDatas['remarks']);  
                $row['company'] = $compDatas['id'];
                $row['member_principle'] = Null;
                // if ($row['member_principle'] == "" & $row['member_relation'] == '1'){
                //     $row['member_principle'] = $clientDatas['id'];
                // }
                $row['policy_declare_date'] = date('Y-m-d');
                $row['create_date'] = date('Y-m-d');
                $datas[] = $row;		
            }, true);                                                 

            $res_sql = $query->into('member')->insert($datas); //comment for temporary
            return true;
            // echo $res_sql;

    }
    }