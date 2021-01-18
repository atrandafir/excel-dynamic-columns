<?php
    
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
    
class ExCol {
  
  static private $_map=array();
  
  static public function get($col, $row=null) {
    if (!in_array($col, self::$_map)) {
      self::$_map[]=$col;
    }
    $index = array_search($col, self::$_map);
    $columnLetter = Coordinate::stringFromColumnIndex($index + 1);
    return $columnLetter.($row?$row:null);
  }
  
  static public function getLast() {
    return Coordinate::stringFromColumnIndex(count(self::$_map));
  }

  static public function reset() {
    self::$_map=array();
  }
  
}
