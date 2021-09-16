<?php

require __DIR__ . '/../vendor/autoload.php';

use Glitchbl\Excel;

$dir = __DIR__ . '/tests';
if (!is_dir($dir))
    mkdir($dir);

$file = "{$dir}/test.xlsx";

$excel = new Excel;
$excel->addRow(['test' => 123, 'Kappa' => 'Keppo']);
$excel->addRow(['Choco' => 'Pistache', 'test' => 456]);
$excel->save($file);

$file = "{$dir}/bytes.csv";

file_put_contents($file, $excel->getBytes('csv'));

$excel = Excel::create(['1ère colonne', '2ème colonne', '3ème colonne'], [
    ['1x1', '1x2', '1x3'],
    ['2x1', '2x2', '2x3'],
    ['3x1', '3x2', '3x3'],
]);

$file = "{$dir}/create.xlsx";

$excel->save($file);

var_dump($excel->toArray());

$excel = Excel::create([
    [
        '1ère colonne' => '1x1',
        '2ème colonne' => '1x2',
        '3ème colonne' => '1x3',
    ],
    [
        '1ère colonne' => '2x1',
        '2ème colonne' => '2x2',
        '3ème colonne' => '2x3',
    ],
    [
        '1ère colonne' => '3x1',
        '2ème colonne' => '3x2',
        '3ème colonne' => '3x3',
    ],
]);

$file = "{$dir}/create2.csv";

$excel->save($file, 'csv');

var_dump($excel->toAssocArray());

$excel = new Excel("{$dir}/test.xlsx");
var_dump($excel->toArray());

$excel = new Excel;
$excel->addRows([
    [1, 2, 3],
], false);
$excel->addRow([4, 5, 6], false);
$excel->addRows([
    [7, 8, 9],
], false);
var_dump($excel->toArray());
