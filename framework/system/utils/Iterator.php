<?php 
namespace Pusaka\Utils;

use Iterator;

class IteratorUtils implements Iterator {

	private $position = 0;
    private $array;

    private $first;
    private $last;
    private $parent;

    public function __set($attribute, $value) {

    	switch ($attribute) {

    		case 'parent':
    			$this->parent = $value;
    			break;

    	}

    }

    public function __get($attribute) {

    	switch ($attribute) {

    		case 'first':

    			$keys = array_keys( $this->array );

    			return ($this->current() === $this->array[$keys[0]]);

    			break;

    		case 'last':

    			$keys = array_keys( $this->array );

    			return ($this->current() === $this->array[$keys[count($keys)-1]]);

    			break;

    		case 'parent':

    			return $this->parent;

    			break;
    		
    		default:
    			return NULL;
    			break;
    	
    	}

    }

    public function __construct(array $array) {
        $this->array = $array;
    }

    public function rewind() {
    	$this->position = 0;
        return reset($this->array);
    }

    public function current() {
        return current($this->array);
    }

    public function key() {
        return key($this->array);
    }

    public function next() {
    	++$this->position;
        next($this->array);
    }

    public function valid() {

    	return key($this->array) !== null;

    }

    public function __destruct() {

        unset($this->parent);
        unset($this->array);

    }

}  