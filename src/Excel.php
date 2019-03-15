<?php

namespace Glitchbl;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class Excel
{
    private $spreadsheet;
    private $sheet;

    private $fichier;

    public function __construct($fichier)
    {
        $this->fichier = $fichier;
        if ($this->isNew()) {
            $this->spreadsheet = new Spreadsheet();
        } else {
            $this->spreadsheet = IOFactory::load($this->fichier);
        }

        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    public function isNew()
    {
        return !is_file($this->fichier);
    }

    public function save($fichier = null)
    {
        $writer = new Xlsx($this->spreadsheet);
        if (!$fichier)
            $writer->save($this->fichier);
        else
            $writer->save($fichier);   
    }

    public function add(array $row)
    {
        $write_columns = function($sheet, $columns) {
            foreach ($columns as $col_index => $column) {
                $sheet->setCellValueByColumnAndRow($col_index + 1, 1, $column);
                $sheet->getStyleByColumnAndRow($col_index + 1, 1)->getFont()->setBold(true);
                $sheet->getColumnDimensionByColumn($col_index + 1)->setWidth(25);
            }
        };

        if ($this->sheet->getHighestRow() === 1) {
            $columns = array_keys($row);
            $write_columns($this->sheet, $columns);
        } else {
            $first_row = $this->sheet->getRowIterator()->current();
            $first_row_cells = $first_row->getCellIterator();
            $columns = [];
            foreach ($first_row_cells as $cell) {
                $columns[] = $cell->getValue();
            }
        }

        $new_columns = array_diff(array_keys($row), $columns);

        if (count($new_columns)) {
            $columns = array_merge($columns, $new_columns);

            $write_columns($this->sheet, $columns);
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
        $writer = new Xlsx($this->spreadsheet);
        $writer->save('php://output');
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
}
