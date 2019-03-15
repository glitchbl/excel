<?php

namespace Glitchbl;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class Excel
{
    private $spreadsheet;
    private $sheet;

    private $file;

    public function __construct($file)
    {
        $this->file = $file;
        if ($this->isNew()) {
            $this->spreadsheet = new Spreadsheet();
        } else {
            $this->spreadsheet = IOFactory::load($this->file);
        }

        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    public function isNew()
    {
        return !is_file($this->file);
    }

    public function save($file = null)
    {
        $writer = new Xlsx($this->spreadsheet);
        if (!$file)
            $writer->save($this->file);
        else
            $writer->save($file);   
    }

    public function add(array $row)
    {
        $columns = [];

        if ($this->sheet->getHighestRow() === 1) {
            $columns = array_keys($row);
            $this->writeColumns($columns);
        } else {
            $first_row = $this->sheet->getRowIterator()->current();
            foreach ($first_row->getCellIterator() as $cell) {
                $columns[] = $cell->getValue();
            }
        }

        $new_columns = array_diff(array_keys($row), $columns);

        if (count($new_columns)) {
            $columns = array_merge($columns, $new_columns);
            $this->writeColumns($columns);
        }

        $row_index = $this->sheet->getHighestRow() + 1;

        foreach ($row as $column => $value) {
            if (is_array($value)) {
                $value = implode("\n", $value);
            }
            $column_index = array_search($column, $columns);
            if ($column_index !== false) {
                $this->sheet->setCellValueByColumnAndRow($column_index + 1, $row_index, $value);
                $this->sheet->getStyleByColumnAndRow($column_index + 1, $row_index)->getAlignment()->setWrapText(true);
            }
        }
    }

    public function getBytes()
    {
        ob_start();
        $writer = new Xlsx($this->spreadsheet);
        $writer->save('php://output');
        return ob_get_clean();
    }

    public function toArray($assoc = true)
    {
        $array = [];
        $columns = [];

        foreach ($this->sheet->getRowIterator() as $row_index => $row) {
            $cells = $row->getCellIterator();
            $cells->setIterateOnlyExistingCells(false);

            $tmp = [];
            $i = 0;
            foreach ($cells as $cell_index => $cell) {
                if ($row_index == 1) {
                    $columns[$cell_index] = $cell->getValue();
                } else {
                    if ($assoc) {
                        $tmp[$columns[$cell_index]] = $cell->getFormattedValue();
                    } else {
                        $tmp[$i++] = [
                            $columns[$cell_index],
                            $cell->getFormattedValue()
                        ];
                    }
                }
            }
            if ($row_index != 1)
                $array[] = $tmp;
        }

        return $array;
    }

    private function writeColumns($columns) {
        foreach ($columns as $col_index => $column) {
            $this->sheet->setCellValueByColumnAndRow($col_index + 1, 1, $column);
            $this->sheet->getStyleByColumnAndRow($col_index + 1, 1)->getFont()->setBold(true);
            $this->sheet->getColumnDimensionByColumn($col_index + 1)->setWidth(25);
        }
    }
}
