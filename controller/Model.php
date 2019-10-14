<?php
class Model {
    private $select;
    private $from;
    private $where;
    private $getDatas;
    private $limit;
    private $values;
    private $columns;

    public function value($values){
        $this->value="VALUES $values";
        return $this;
    }
    public function column($column){
        $this->columns="$column";
        return $this;
    }
    public function insert($tableName){
        return "insert into $tableName";
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
        $que = "SELECT id,client,company FROM member WHERE id in (5369890,5369871)";
        $d = $koneksi->query($que); 
        $d = $d->fetch_assoc();
        foreach ($d as $value) {
            echo "$value <br>";
        }
        var_dump($d);   
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

}