<?php

function prt($val, $die = false){
  echo '<pre>';
  print_r($val);
  echo '<br>';
  if ($die) die('dead');
}

function prtd($val){
  prt($val, true);
}