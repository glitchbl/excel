<?php

namespace Glitchbl;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Exception;

class Excel
{
    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    private $spreadsheet;

    /**
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    private $sheet;

    /**
     * @var string
     */
    private $file;

    /**
     * @param string|null $file
     * @return void
     */
    public function __construct($file = null)
    {
        $this->file = $file;
        $this->make();
    }

    /**
     * @param array $columns
     * @param array $rows
     * @return \Glitchbl\Excel
     */
    static public function create($columns, $rows)
    {
        $excel = new self;
        $excel->writeColumns($columns);
        $excel->addRows($rows);
        return $excel;
    }

    /**
     * @param array $columns
     * @return void
     */
    private function writeColumns($columns) {
        foreach ($columns as $col_index => $column) {
            $this->sheet->setCellValueByColumnAndRow($col_index + 1, 1, $column);
            $this->sheet->getStyleByColumnAndRow($col_index + 1, 1)->getFont()->setBold(true);
            $this->sheet->getColumnDimensionByColumn($col_index + 1)->setWidth(25);
        }
    }

    /**
     * @return void
     */
    protected function make() {
        if ($this->file) {
            if (!is_file($this->file)) {
                throw new Exception("Le fichier {$this->file} n'existe pas");
            }
            $this->spreadsheet = IOFactory::load($this->file);
        } else {
            $this->spreadsheet = new Spreadsheet;
        }
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    /**
     * @param string|null $file
     * @return void
     */
    public function save($file = null)
    {
        $writer = new Xlsx($this->spreadsheet);
        if ($file) {
            $writer->save($file);   
        } elseif ($this->file) {
            $writer->save($this->file);
        } else {
            throw new Exception('Veuillez spécifier un nom de fichier');
        }
    }

    /**
     * @param array $row
     * @return void
     */
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

    /**
     * @param array $rows
     * @return void
     */
    public function addRows(array $rows)
    {
        $row_index = $this->sheet->getHighestRow() + 1;

        foreach ($rows as $row) {
            foreach ($row as $column_index => $value) {
                if (is_array($value)) {
                    $value = implode("\n", $value);
                }
                $this->sheet->setCellValueByColumnAndRow($column_index + 1, $row_index, $value);
                $this->sheet->getStyleByColumnAndRow($column_index + 1, $row_index)->getAlignment()->setWrapText(true);
            }
            $row_index++;
        }
    }

    /**
     * @return string
     */
    public function getBytes()
    {
        ob_start();
        $writer = new Xlsx($this->spreadsheet);
        $writer->save('php://output');
        return ob_get_clean();
    }

    /**
     * @return array
     */
    public function toArray()
    {
        $array = [];

        foreach ($this->sheet->getRowIterator() as $row) {
            $cells = $row->getCellIterator();
            $cells->setIterateOnlyExistingCells(false);

            $tmp = [];
            foreach ($cells as $cell) {
                $tmp[] = $cell->getFormattedValue();
            }
            $array[] = $tmp;
        }

        return $array;
    }

    /**
     * @param boolean $assoc
     * @return array
     */
    public function toAssocArray($assoc = true)
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
