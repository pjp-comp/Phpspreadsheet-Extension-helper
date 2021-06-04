<?php

require_once('PhpSpreadSheet.php');

            $excel = new PhpSpreadSheet('Pragnesh');

            // to make width flexible OR pass simple array to set header
            $headings = [
                ['col_name'=>'first fsad dafsad dafsd f adfas',
                  'width'=>true
                    ],
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

            die('fsa');

?>
