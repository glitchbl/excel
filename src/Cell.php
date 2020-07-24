<?php

namespace Glitchbl\Excel;

use PhpOffice\PhpSpreadsheet\Cell\Cell as PhpSpreadsheetCell;
use ReflectionFunction;
use Closure;

class Cell {
    /**
     * @var mixed
     */
    protected $value;

    /**
     * @var mixed
     */
    protected $formattedValue;

    /**
     * @param \PhpOffice\PhpSpreadsheet\Cell\Cell $cell
     * @return void
     */
    public function __construct(PhpSpreadsheetCell $cell)
    {
        $this->value = $cell->getValue();
        $this->formattedValue = $cell->getFormattedValue();
    }

    /**
     * @return string
     */
    public function __toString()
    {
        return $this->formattedValue;
    }

    /**
     * @param mixed $value
     * @param \Closure $closure
     * @return mixed
     */
    protected function getValueWithClosure($value, Closure $closure)
    {
        $reflector = new ReflectionFunction($closure);
        return $closure->call($reflector->getClosureThis(), $value);
    }

    /**
     * @param \Closure|null $closure
     * @return mixed
     */
    public function getValue(?Closure $closure = null)
    {
        if (!is_null($closure))
            return $this->getValueWithClosure($this->value, $closure);
        return $this->value;
    }

    /**
     * @param \Closure|null $closure
     * @return mixed
     */
    public function getFormattedValue(?Closure $closure = null)
    {
        if (!is_null($closure))
            return $this->getValueWithClosure($this->formattedValue, $closure);
        return $this->formattedValue;
    }
}
