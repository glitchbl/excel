<?php

require __DIR__ . '/../vendor/autoload.php';

use Glitchbl\Excel;

$dir = __DIR__ . '/tests';
if (!is_dir($dir))
    mkdir($dir);

$file = "{$dir}/save.xlsx";

$excel = new Excel;
$excel->add(['test' => 123, 'Kappa' => 'Keppo']);
$excel->add(['Choco' => 'Pistache', 'test' => 456]);
$excel->save($file);

$file = "{$dir}/bytes.xlsx";

file_put_contents($file, $excel->getBytes());

$excel = Excel::create(['1ère colonne', '2ème colonne', '3ème colonne'], [
    ['1x1', '1x2', '1x3'],
    ['2x1', '2x2', '2x3'],
    ['3x1', '3x2', '3x3'],
]);

$file = "{$dir}/create.xlsx";

$excel->save($file);