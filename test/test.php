<?php

require __DIR__ . '/../vendor/autoload.php';

use Glitchbl\Excel;

$excel = new Excel(__DIR__ . '/test.xlsx');
$excel->add(['test' => 123, 'Kappa' => 'Keppo']);
$excel->add(['Choco' => 'Pistache', 'test' => 456]);
$excel->save();