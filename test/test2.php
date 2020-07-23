<?php

require __DIR__ . '/../vendor/autoload.php';

use Glitchbl\Excel\Excel;

$excel = Excel::create(['1ère colonne', '2ème colonne', '3ème colonne'], [
    ['1x1', '1x2', '1x3'],
    ['2x1', '2x2', '2x3'],
    ['3x1', '3x2', '3x3'],
]);

$a = $excel->toArray()[0][0];

echo $a->sanitize(fn($v) => trim(str_replace('colonne', '', $v)));
