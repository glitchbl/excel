<?php

namespace Glitchbl;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Glitchbl\Excel\Sheet;
use Exception;

/**
 * @method \Glitchbl\Excel\Sheet asRaw()
 * @method \Glitchbl\Excel\Sheet asFormated()
 * @method void writeColumns($columns)
 * @method array toArray()
 * @method array toAssocArray($assoc = true)
 * @method void addRow(array $row, $assoc = true)
 * @method void addRows(array $rows, $assoc = true)
 */
class Excel
{
    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    protected $spreadsheet;

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

    public function __call($name, $arguments)
    {
        return (new Sheet($this->spreadsheet->getActiveSheet()))->{$name}(...$arguments);
    }
}
