<?php

namespace Glitchbl\Excel;

use Closure;

class Cell {
    protected $value;

    protected $formattedValue;

    public function __construct(\PhpOffice\PhpSpreadsheet\Cell\Cell $cell)
    {
        $this->value = $cell->getValue();
        $this->formattedValue = $cell->getFormattedValue();
    }

    public function __toString()
    {
        return $this->formattedValue;
    }

    public function sanitize(Closure $closure)
    {
        return $closure->call(null, $this->value);
    }
}
