<?php

class AddingDatas{
    public function client($client){        
        $model_obj = new Model();     
        $data_source = $model_obj->select('*')->from('client')->where("id = $client"); 
        $clientDatas = $model_obj->getDatas($data_source);
        return $clientDatas;
    }

    public function company($client,$compName){        
        $model_obj = new Model();
        $data_source = $model_obj->select('*')->from('company')->where("client = $client")->limit(1);
        $company_datas = $model_obj->getDatasLimit($data_source);
        return $company_datas;
    }

    public function plan($plan){    
        $model_obj = new Model();
        $data_source = $model_obj->select('*')->from('plan')->where("name = '$plan'");
        $planDatas = $model_obj->getDatas($data_source);
        return $planDatas;
    }
    
}
