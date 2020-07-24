<?php

namespace Glitchbl\Excel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Exception;

class Excel
{
    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    protected $spreadsheet;

    /**
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    protected $sheet;

    /**
     * @var string
     */
    protected $file;

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
     * @param array|null $rows
     * @return \Glitchbl\Excel
     */
    static public function create($columns, $rows = null)
    {
        $excel = new static;
        if (!is_null($rows)) {
            $excel->writeColumns($columns);
            $excel->addRows($rows, false);
        } else {
            $excel->addRows($columns);
        }
        return $excel;
    }

    /**
     * @param array $columns
     * @return void
     */
    protected function writeColumns($columns) {
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
        if (!is_null($this->file)) {
            if (!is_file($this->file)) {
                throw new Exception("File {$this->file} does not exist");
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
    public function save($file = null, $type = 'xlsx', $delimiter = ',')
    {
        if ($type == 'xlsx')
            $writer = new Xlsx($this->spreadsheet);
        elseif ($type == 'csv') {
            $writer = new Csv($this->spreadsheet);
            $writer->setDelimiter($delimiter);
        } else {
            throw new Exception("Type: [xlsx,csv]");
        }

        if (!is_null($file)) {
            $writer->save($file);
            $this->file = $file;
        } elseif (!is_null($this->file)) {
            $writer->save($this->file);
        } else {
            throw new Exception('Please specify a filename');
        }
    }

    /**
     * @param array $rows
     * @return void
     */
    public function addRow(array $row, $assoc = true)
    {
        $highest_row = $this->sheet->getHighestRow();
        if ($assoc == false) {
            $row_index = $highest_row + 1;

            if ($row_index == 2) {
                $tmp = $this->toArray();
                if (count($tmp) == 1 && count($tmp[0]) == 1 && $tmp[0][0] == '')
                    $row_index = 1;
            }

            foreach ($row as $column_index => $value) {
                if (is_array($value)) {
                    $value = implode("\n", $value);
                }
                $this->sheet->setCellValueByColumnAndRow($column_index + 1, $row_index, $value);
                $this->sheet->getStyleByColumnAndRow($column_index + 1, $row_index)->getAlignment()->setWrapText(true);
            }
        } else {
            $columns = [];

            if ($highest_row === 1) {
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

            $row_index = $highest_row + 1;

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
    }

    /**
     * @param array $rows
     * @return void
     */
    public function addRows(array $rows, $assoc = true)
    {
        foreach ($rows as $row) {
            $this->addRow($row, $assoc);
        }
    }

    /**
     * @return string
     */
    public function getBytes($type = 'xlsx', $delimiter = ',')
    {
        ob_start();
        if ($type == 'xlsx')
            $writer = new Xlsx($this->spreadsheet);
        elseif ($type == 'csv') {
            $writer = new Csv($this->spreadsheet);
            $writer->setDelimiter($delimiter);
        } else {
            throw new Exception("Type: [xlsx,csv]");
        }
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
                $tmp[] = new Cell($cell);
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
                        $tmp[$columns[$cell_index]] = new Cell($cell);
                    } else {
                        $tmp[$i++] = [
                            $columns[$cell_index],
                            new Cell($cell)
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
