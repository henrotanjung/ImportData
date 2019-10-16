<?php
session_start();
use Pusaka\Database\Manager as Database;
use Pusaka\Database\DatabaseException;
use Pusaka\Microframework\Loader;
use Pusaka\Microframework\Log;
use Pusaka\Library\Import;
use Pusaka\Library\Export;
use Pusaka\Library\AAICopy;


Loader::lib('import');

class Enrollment
{

    public function insert($fileName, $column, $client, $company)
    {
        // $folder = '//192.168.0.20/Z Folder/IT/INTERNAL ONLY/003 Internal Minutes of Meeting/Programmer Team/005 Hendro/';
        $folder = '//10.10.10.28/WebServer/intsys/aaienrollment/';
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
        $Import->each(function ($row, $index) use (&$client, &$datas, &$clientDatas, &$add_data_obj,&$company) {

            include 'cleaning_data.php';
            if ($row['member_id'] === '') {
                return;
            }
            $planDatas = $add_data_obj->plan($row['plan']);
            // if ()
            $row['program'] = $planDatas['program'];
            $row['client'] = $clientDatas['id'];
            $row['plan'] = $planDatas['id'];
            $compDatas = $add_data_obj->company($client, $planDatas['remarks']);
            $row['company'] = $company;
            // echo $row['policy_effective_date'];
            foreach (array_keys($row) as $key) {
                if ($row["$key"] == "") {
                    $row["$key"] = Null;
                }
                                
            }
            
            if ($row['policy_issue_date'] == ''){
                $row['policy_issue_date'] = $row['policy_effective_date'];
            }
            
            $row['policy_declare_date'] = date('Y-m-d');
            
            $row['create_date'] = date('Y-m-d');
            $datas[] = $row;
        }, true);

        $res_sql = $query->into('member')->insert($datas); //comment for temporary
        include ('../config/koneksi.php');
        $que = "SELECT id FROM member WHERE client = 289 and company = 327";
        $dque = $koneksi->query($que);
        $dque = mysqli_fetch_array($dque);        

        $dtget = $model_obj->select('id,client,company')->from('member')->where("client = $client and company = $company and member_relation=1 ");
        $dtget1 = $model_obj->getDataMember($dtget);
        while($row = $dtget1->fetch_assoc()){
            $id = $row['id'];
            $query = "UPDATE member set member_principle= $id where id = $id";            
            $koneksi->query($query);
        }          
        
    }

    public function updateLapse($fileName,$column){
        $folder = '//192.168.0.20/Z Folder/IT/INTERNAL ONLY/003 Internal Minutes of Meeting/Programmer Team/005 Hendro/';
        // $folder = '//10.10.10.28/WebServer/intsys/aaienrollment/';
        $file_xlsx = $fileName;
        $Import = Import::open(path($folder) . $file_xlsx);
        $Import->setMap($column);
        $query = Database::on('dblocal_dev_aia')->builder();
        $datas = array();
        include 'Model.php';
        include 'adding_data.php';
        $model_obj = new Model();
        $add_data_obj = new AddingDatas();
        // $compDatas = $add_data_obj->company($client);          
        include ('../config/koneksi.php');  
        $Import->each(function ($row, $index) use (&$datas,$koneksi) {
            if ($row['member_id'] === '') {
                return;
            }           
            $policy_lapse_date = $row['policy_lapse_date'];
            $policy_lapse_date = date('Y/m/d', strtotime($policy_lapse_date));
            
            $datas[] = $row;
        }, true);

        foreach ($datas as $data){            
            $policy_lapse_date = $data['policy_lapse_date'];
            $member_id = $data['member_id'];
            $client = $data['client'];
            $policy_lapse_date = date('Y-m-d', strtotime($policy_lapse_date));            
            // $query = "UPDATE member set policy_lapse_date = '$policy_lapse_date',policy_status = 0 where member_id = $member_id and client= $client";   
            // $koneksi->query($query);
            $update = $model_obj->update('member')->set("policy_lapse_date = '$policy_lapse_date', policy_status = 0")->where("member_id = $member_id and id=53691044");
            $queryUpdate = $model_obj->executeUpdate($update);
            
            $res = $koneksi->query($queryUpdate);
            $_SESSION['lapse_succed!'] = $res;
            return $res;

        }
        
    }
}