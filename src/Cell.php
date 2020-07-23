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

    public function getValue(?Closure $closure = null)
    {
        if (!is_null($closure))
            return $closure->call($this, $this->value);
        return $this->value;
    }

    public function getFormattedValue(?Closure $closure = null)
    {
        if (!is_null($closure))
            return $closure->call($this, $this->formattedValue);
        return $this->formattedValue;
    }
}
