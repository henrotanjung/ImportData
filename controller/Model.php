<?php
class Model {
    private $select;
    private $from;
    private $where;
    private $getDatas;
    private $limit;

    public function insert($tableName,$values=[]){
        return "insert into $tableName Values $values";
    }

    public function select($values){
        $this->select = "select $values ";
        return $this;
    }    

    public function from($tableName){
        $this->from = "from $tableName ";
        return $this;
    }
    public function where($parameter){
        $this->where = "where $parameter ";
        return $this;
    } 
    
    public function limit($limit){
        $this->limit = "LIMIT $limit ";
        return $this;
    }
    public function getDatas($sourceDatas){
        include '../config/koneksi.php';
        // global $koneksi;
        // var_dump($sourceDatas->where);
        $select = $sourceDatas->select;
        $from = $sourceDatas->from;
        $where = $sourceDatas->where;
        $query = $select.$from.$where;                
        $res = $this->getDatas = $koneksi->query($query);        
        return $res->fetch_assoc();
    }

    public function getDatasLimit($sourceDatas){
        include '../config/koneksi.php';
        // global $koneksi;
        // var_dump($sourceDatas->where);
        $select = $sourceDatas->select;
        $from = $sourceDatas->from;
        $where = $sourceDatas->where;
        $limit = $sourceDatas->limit;
        $query = $select.$from.$where.$limit;                
        $res = $this->getDatas = $koneksi->query($query);        
        return $res->fetch_assoc();
    }

    public function getDatas_Obselect($tableName,$client){            
        $query = $this->select($tableName,$client);
        $get_datas = $koneksi->query($query);        
        return $get_datas = $get_datas->fetch_assoc();
    }

}