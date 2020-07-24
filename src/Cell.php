<?php

namespace Glitchbl\Excel;

use PhpOffice\PhpSpreadsheet\Cell\Cell as PhpSpreadsheetCell;
use ReflectionFunction;
use Closure;

class Cell {
    protected $value;

    protected $formattedValue;

    public function __construct(PhpSpreadsheetCell $cell)
    {
        $this->value = $cell->getValue();
        $this->formattedValue = $cell->getFormattedValue();
    }

    public function __toString()
    {
        return $this->formattedValue;
    }

    protected function getValueWithClosure($value, Closure $closure)
    {
        $reflector = new ReflectionFunction($closure);
        return $closure->call($reflector->getClosureThis(), $value);
    }

    public function getValue(?Closure $closure = null)
    {
        if (!is_null($closure))
            return $this->getValueWithClosure($this->value, $closure);
        return $this->value;
    }

    public function getFormattedValue(?Closure $closure = null)
    {
        if (!is_null($closure))
            return $this->getValueWithClosure($this->formattedValue, $closure);
        return $this->formattedValue;
    }
}
