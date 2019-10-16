<?php
class Model {
    private $select;
    private $from;
    private $where;
    private $getDatas;
    private $getDatasFectArray;
    private $limit;
    private $values;
    private $valueSet;
    private $set;
    private $updateQ;

    public function value($values){
        $this->value="VALUES $values ";
        return $this;
    }

    public function valueSet($values){
        $this->valueSet="=$values ";
        return $this;
    }
    
    public function set($columnValues){
        $this->set="SET $columnValues ";
        return $this;
    }
    public function insert($tableName){
        $this->insert="insert into $tableName ";
        return $this;
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
    public function update($tableName){
        $set = $this->set;
        $where = $this->where;
        $this->update = "UPDATE $tableName ";
        return $this;
    }

    public function executeUpdate($source){        
        $update = $this->update;
        $set = $this->set;
        $valueSet = $this->valueSet;
        $where = $this->where;
        $query = $update.$set.$where;      
        return $query;
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

    public function getDataMember($sourceDatas){
        include '../config/koneksi.php';
        // global $koneksi;
        // var_dump($sourceDatas->where);
        $select = $sourceDatas->select;
        $from = $sourceDatas->from;
        $where = $sourceDatas->where;
        $query = $select.$from.$where;        
        $res = $koneksi->query($query);     
        
        return $res;
    }


    public function getDatasFectArray($sourceDatas){
        include '../config/koneksi.php';
        $select = $sourceDatas->select;
        $from = $sourceDatas->from;
        $where = $sourceDatas->where;
        $query = $select.$from.$where;                
        $res = $this->getDatasFectArray = $koneksi->query($query);         
        
        return mysqli_fetch_array($res);
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

}