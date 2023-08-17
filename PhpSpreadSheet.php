<?php


namespace App\Utility;

ini_set('memory_limit', '-1');
ini_set('max_execution_time', 3000);


//require_once(ROOT . '/vendor/box/spout/src/Spout/Autoloader/autoload.php');
require ROOT.'/vendor/autoload.php';

use Cake\View\ViewBuilder;
use Cake\Filesystem\Folder;
use Cake\Filesystem\File;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;




class PhpSpreadSheet{

/*

            // pass custom style instead of style name from calling method
            $customStyle = [
                'font' => [
                    'bold' => true,
                    'size' => 10,
                ],
                'alignment' => [
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                ]
            ];


            // for multisheet --start
            $excel = new PhpSpreadSheet('Pragnesh', true);
            $sheets = ['one', 'two', 'three'];
            foreach ($sheets as $key=>$sheet){
                $excel->createNewSheet($sheet, true, $key);
                $excel->setHeading($headings,"",'TABLE_HEAD');
                $excel->setArrayData($arrayData);
            }
            // for multisheet --false



            $excel = new PhpSpreadSheet('Pragnesh');

            // pass individual header column style
            $headings = [
                 ['col_name'=>'first fsad dafsad dafsd f adfas',
                  'width'=>true,
                  'style'=>[
                        'fill' => [
                                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                                'color' => ['rgb' => '9c9c9c'],
                            ]
                  ]
                 ],,
                 ['col_name'=>'second',
                  'width'=>20
                    ],
                 ['col_name'=>'third',
                  'width'=>30
                    ],
                 ['col_name'=>'forth',
                  'width'=>false
                    ]
            ];
            $arrayData = [
                ['Q1',   12,   15,   21],
                ['Q2',   56,   73,   86],
                ['Q3',   52,   61,   69],
                ['Q4',   30,   32,    null],
            ];
            $totaling = [10,30,50,60];

            $excel->setProperty("Pragnesh", "generate excel");


            $excel->setTitle("Hello dudne nice to meet you", 4, "B2", "TABLE_TITLE");
            $excel->setHeading($headings,"",'TABLE_HEAD');
            $excel->setArrayData($arrayData);
//            $excel->generateExcelFromHtml();
            $excel->setRowGap(4);
            $excel->setTitle("Hello again", 4, $excel->lastCellAddress(), "TABLE_TITLE");
            $excel->setHeading($headings,"",'TABLE_HEAD');
            $excel->setArrayData($arrayData);
            $excel->setHeading($totaling,"",'TABLE_HEAD_LIGHT');

            $excel->write();

            die('fsa');*/


    public $spreadsheet;
    private $sheet;
    private $writer ;
    private $rowNum = 1;
    private $startCell = "A1";
    private $removeMainSheet = false;
    private $STYLES = [

       'TABLE_TITLE'=>[
                'font' => [
                    'bold' => true,
                    //'color' => ['rgb' => '088DCF'],
                    'size' => 12,
                ],
               'alignment' => [
                            'horizontal' => Alignment::HORIZONTAL_LEFT,
                ],
                /*'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'color' => ['argb' => 'FFCCFFCC'],
                ]*/
        ],
        'TABLE_TITLE_CENTER'=>[
            'font' => [
                'bold' => true,
                //'color' => ['rgb' => '088DCF'],
                'size' => 12,
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
            ],
            /*'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'color' => ['argb' => 'FFCCFFCC'],
            ]*/
        ],
       'TABLE_HEAD' => [
                'font' => [
                            //'bold' => true,
                            'color' => ['rgb' => 'FFFFFF'],
                            'size' => 10
                        ],

               'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'color' => ['rgb' => '000000'],
                ],
                /*'borders' => [
                            'bottom' => [
                                'borderStyle' => Border::BORDER_THIN,
                            ],*/
       ],
       'TABLE_HEAD_LIGHT' => [
                'font' => [
                            // 'bold' => true,
                            'color' => ['rgb' => 'FFFFFF'],
                            'size' => 10
                        ],

               'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'color' => ['rgb' => '9c9c9c'],
                ],
                /*'borders' => [
                            'bottom' => [
                                'borderStyle' => Border::BORDER_THIN,
                            ],*/
       ],
        'ONLY_BOLD' => [
                'font' => [
                            'bold' => true
                        ],

               'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'color' => ['rgb' => '9c9c9c'],
                ]
       ],
    ];

    /*
     * $removeMainSheet = false : set true only if you surely multisheet
     * */
    public function __construct($sheetName = "vrmanaged", $removeMainSheet = false)
    {
        $this->removeMainSheet = $removeMainSheet;
        $this->spreadsheet = new Spreadsheet();
        $this->spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
        $this->spreadsheet->getDefaultStyle()->getFont()->setSize(10);

        $validLocale = \PhpOffice\PhpSpreadsheet\Settings::setLocale('en_us');

        // https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#setting-the-default-column-width
        // $this->spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);
        // $this->spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(12);

        $this->setSheetTitle($sheetName);
        // $this->sheet = $this->spreadsheet->getActiveSheet();
        // $this->sheet = $this->spreadsheet->getActiveSheet()->getProtection()->setSheet(true);
    }

    public function setSheetTitle($title = "test"){
       $this->sheet = $this->spreadsheet->getActiveSheet()->setTitle($title);
       return $this;
    }

    public function setActiveSheet($name){
        $this->sheet = $this->spreadsheet->setActiveSheetIndexByName($name);
        return $this;
    }

    public function createNewSheet($name,$setActive = false , $sheetIndex = 0 , $startCell = 'A1'){

        if($this->removeMainSheet){
            $this->spreadsheet->removeSheetByIndex(0);
            $this->removeMainSheet = false;
        }

        $myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($this->spreadsheet, $name);
        $this->spreadsheet->addSheet($myWorkSheet, $sheetIndex);


        if($setActive){
            $this->startCell = ($startCell == '') ? 'A1' : $startCell;
            $this->setActiveSheet($name);
            /*echo "<pre>";
            print_r($this->spreadsheet->getActiveSheet());
            die;*/

        }
         return $this;
    }

    /*
    Set properties like : title, subject, descreption
    */
    public function setProperty($title = "", $subject = "", $descreption = "Demo"){
        $this->spreadsheet->getProperties()->setCreator('test')
        ->setLastModifiedBy('test')
        ->setTitle($title)
        ->setSubject($subject)
        ->setDescription($descreption);
        //->setKeywords('office 2007 openxml php')
        // ->setCategory('Test result file');
    }

    public function writeFrom($cellAddress = "A1"){
        $this->startCell = $cellAddress;
    }

    /*
    Coordinate::coordinateFromString('A2')
    echo Coordinate::stringFromColumnIndex(1);
        echo "<br>";
        echo Coordinate::columnIndexFromString('E');
        die;*/
    private function getRowNum($cell){

     /*preg_match('/(\d+)/', $cell, $nums);
     return $nums[0];*/

     return  Coordinate::coordinateFromString($cell)[1];

    }

    private function getAlpha($cell){
        return  Coordinate::coordinateFromString($cell)[0];
//        return preg_split('/(\d+)/', $cell)[0];
    }
    private function getColumnNumber($number){
        return Coordinate::columnIndexFromString($number);
    }
    private function getColumn($last){
        return Coordinate::stringFromColumnIndex($last);
    }


    /*
    Get cell last cell positoin of after setting the data
    */
    public function lastCellAddress(){
        return $this->startCell;
    }

    /*
    create row gaps
    pass number of rows in argument
    */
    public function setRowGap($numberRows = 0){
        $this->startCell = $this->getAlpha($this->startCell).($this->getRowNum($this->startCell)+$numberRows);
    }

    /*
    Set Excel title here
    */
    public function setTitle($titleName, $mergeTill, $startCell ,$styleName = ""){

        if($startCell != ""){
            $this->startCell = $startCell;
        }


        $this->sheet->setCellValue($this->startCell, $titleName);

        /*$this->spreadsheet->getActiveSheet()->getStyle($this->startCell)
            ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
        $this->spreadsheet->getActiveSheet()->getStyle($this->startCell)
            ->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
        $this->spreadsheet->getActiveSheet()->getStyle($this->startCell)
            ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
        $this->spreadsheet->getActiveSheet()->getStyle($this->startCell)
            ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
        $this->spreadsheet->getActiveSheet()->getStyle($this->startCell)
            ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
        $this->spreadsheet->getActiveSheet()->getStyle($this->startCell)
            ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
        $this->spreadsheet->getActiveSheet()->getStyle($this->startCell)
            ->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $this->spreadsheet->getActiveSheet()->getStyle($this->startCell)
            ->getFill()->getStartColor()->setARGB('FFFF0000');*/

        if($styleName != ""){
            $this->spreadsheet->getActiveSheet()->getStyle($this->startCell)->applyFromArray($this->STYLES[$styleName]);;
        }


        $tillCol = ($this->getRowNum($this->startCell) + $mergeTill - 1);
        $col = $this->getColumn($mergeTill);
        $row = $this->getRowNum($this->startCell);
        $endCell = $col.$row;


        $this->mergeCell($this->startCell,$endCell);

        $this->rowNum = $row+2;
        $this->startCell = $this->getAlpha($this->startCell).$this->rowNum;


    }


    /*

    Pass headerdata in array OR if you looking for cusomised width of column.
    true : set automatically based on content
    false : default width
    integer : 15 // gives you exact number of width

    $colData = [
        ['col_name'=>'first',
          'width'=>true | number
        ],
        ['col_name'=>null,
          'width'=>true | number
        ]
    ];*/

    public function setHeading($arrayData, $startCell = "", $styleName = ""){
        if(!empty($arrayData) && is_array($arrayData[0]) && isset($arrayData[0]['col_name'])){
            $this->setHeadingCustom($arrayData, $startCell, $styleName );
        }else{
            $this->setHeadingNormally($arrayData, $startCell , $styleName );
        }
    }
    public function setHeadingCustom($arrayData, $startCell = "", $styleName = ""){

        if($startCell != ""){
            $this->startCell = $startCell;
        }


        $columNumber = $this->getColumnNumber($this->getAlpha($this->startCell));
        $tillCol =  $columNumber + (count($arrayData) - 1);

        $col = $this->getColumn($tillCol);
        $row = $this->getRowNum($this->startCell);


        if($styleName != ""){
            $endCell = $col.$row;
            $this->applyStyle($styleName , $this->startCell, $endCell);
        }


        $column = $this->getAlpha($this->startCell);
        $row = $this->getRowNum($this->startCell);

        foreach ($arrayData as $coldata){

            if(isset($coldata['style']) && !empty($coldata['style'])){
                $this->spreadsheet->getActiveSheet()->getStyle($column.$row)->applyFromArray($coldata['style']);
            }

            if(isset($coldata['width']) && is_numeric($coldata['width'])){
                $this->spreadsheet->getActiveSheet()->getColumnDimension($column)->setWidth($coldata['width']);
            }elseif(isset($coldata['width']) && is_bool($coldata['width']) && $coldata['width']){
                $this->spreadsheet->getActiveSheet()->getColumnDimension($column)->setAutoSize(true);
            }else{
                $this->spreadsheet->getActiveSheet()->getColumnDimension($column)->setAutoSize(false);
            }
            $this->spreadsheet->getActiveSheet()->setCellValue($column . $row, $coldata['col_name']);
            $column = $this->getColumn($this->getColumnNumber($column) + 1);
        }

        $columNumber = $this->getColumnNumber($this->getAlpha($this->startCell));
        $tillCol =  $columNumber + (count($arrayData) - 1);

        $col = $this->getColumn($tillCol);
        $row = $this->getRowNum($this->startCell);


        if($styleName != ""){
            $endCell = $col.$row;
            $this->applyStyle($styleName , $this->startCell, $endCell);
        }

        $this->rowNum = $row+1;
        $this->startCell = $this->getAlpha($this->startCell).$this->rowNum;


    }


    public function setHeadingNormally($arrayData, $startCell = "", $styleName = ""){


        if($startCell != ""){
            $this->startCell = $startCell;
        }


        $columNumber = $this->getColumnNumber($this->getAlpha($this->startCell));
        $tillCol =  $columNumber + (count($arrayData) - 1);

        $col = $this->getColumn($tillCol);
        $row = $this->getRowNum($this->startCell);


        if($styleName != ""){
            $endCell = $col.$row;

            $this->applyStyle($styleName , $this->startCell, $endCell);
        }

        $this->spreadsheet->getActiveSheet()
            ->fromArray(
                $arrayData,
                NULL,
                ($startCell != "") ? $startCell : $this->startCell,
                true
            );
        $this->rowNum = $row+1;
        $this->startCell = $this->getAlpha($this->startCell).$this->rowNum;

    }

    public function setInColumn($rowArray = [], $startCell = ""){
        // $rowArray = ['Value1', 'Value2', 'Value3', 'Value4'];
        $columnArray = array_chunk($rowArray, 1);
        $this->spreadsheet->getActiveSheet()
            ->fromArray(
                $columnArray,
                NULL,
                ($startCell != "") ? $startCell : $this->startCell
            );
    }

    /*
    ---pass data in array
    ---null if want to blank cell
    */
    public function setArrayData($arrayData = [], $startCell = "", $styleName = ""){

        if(!empty($arrayData)){
            if($startCell != ""){
                $this->startCell = $startCell;
            }
            $tillCol = ($this->getRowNum($this->startCell) + count($arrayData[0]) - 1);
            $col = $this->getColumn($tillCol);
            $row = $this->getRowNum($this->startCell);

            $this->spreadsheet->getActiveSheet()
                ->fromArray(
                    $arrayData,
                    null,
                    ($startCell != "") ? $startCell : $this->startCell,
                    true
                );

            $this->rowNum = $row + count($arrayData);
            $this->startCell = $this->getAlpha($this->startCell).$this->rowNum;

        }
    }

    public function setAutoFilter($column = "C"){

        // https://phpspreadsheet.readthedocs.io/en/latest/topics/autofilters/
        //        $this->spreadsheet->getActiveSheet()->setAutoFilter('A1:E20');
        $this->spreadsheet->getActiveSheet()->setAutoFilter(
            $this->spreadsheet->getActiveSheet()
                ->calculateWorksheetDimension()
        );

        $autoFilter = $this->spreadsheet->getActiveSheet()->getAutoFilter();
        $autoFilter->getColumn($column);
    }

    public function clearingWorkbook(){
        $this->spreadsheet->disconnectWorksheets();
        unset($this->spreadsheet);
    }



    public function write($fileName = "sample", $filedir = "booking_reports", $pdfOrExcel = "excel"){


        //$path = WWW_ROOT.$filedir.'/'.$fileName;
        if($pdfOrExcel == "pdf"){
//            $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, 'Pdf');
//            $writer->save($path);

        }else{
            try {

            // $writer = new Xlsx($this->spreadsheet);
            // $writer->save($path);


            // Redirect output to a client’s web browser (Xls)
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="'.$fileName.'.xls"');
            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');

            $writer = IOFactory::createWriter($this->spreadsheet, 'Xls');
            $writer->save('php://output');
            exit;

//            $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, "Xlsx");
//            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//            header('Content-Disposition: attachment; filename="'.$fileName.'"');
//            $writer->save("php://output");

            } catch(\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
                $this->clearingWorkbook();
                die('Error file: '.$e->getMessage());
            }
        }

        $this->clearingWorkbook();
        return true;
    }

    public function applyStyle($styleName, $rangesStart, $rangeEnd){
        //
        // $spreadsheet->getActiveSheet()->getStyle('A3')-> applyFromArray($styleArray);
        // $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()->setFillType(Fill::FILL_SOLID);
        // $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()->getStartColor()->setARGB('FF808080');

        // set custom style
        if(is_array($styleName) && !is_string($styleName)){
            $this->spreadsheet->getActiveSheet()->getStyle($rangesStart.':'.$rangeEnd)->applyFromArray($styleName);
        }else if(isset($this->STYLES[$styleName])){
            $this->spreadsheet->getActiveSheet()->getStyle($rangesStart.':'.$rangeEnd)->applyFromArray($this->STYLES[$styleName]);
        }else{

        }
    }



    public function generateExcelFromHtml($htmlString = ""){
        // https://phpspreadsheet.readthedocs.io/en/latest/topics/reading-and-writing-to-file/
        $htmlString = '<table>
                <thead>
                <tr>
                    <th colspan="2">PRaga</th>
                </tr>
                </thead>
                <tbody>
                  <tr>
                    <td  colspan="2"><b>Nice heading</b></td>
                  </tr>
                  <tr>
                      <td>one</td>
                      <td>Two</td>
                  </tr>
                  </tbody>
              </table>';

            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Html();
            $this->spreadsheet = $reader->loadFromString($htmlString);

    }

    public function mergeCell($start, $end){
        $this->spreadsheet->getActiveSheet()->mergeCells($start.':'.$end);
    }
    public function unMergeCell($start, $end){
        $this->spreadsheet->getActiveSheet()->unmergeCells($start.':'.$end);
    }





    /*

    default width  = 64 pixels (e.g. 8.43 chars)

    $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(12);
    $spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);

    Setting a row's height
    $spreadsheet->getActiveSheet()->getRowDimension('10')->setRowHeight(100);

    The following code inserts 2 new rows, right before row 7:
    $spreadsheet->getActiveSheet()->insertNewRowBefore(7, 2);


    */

}
