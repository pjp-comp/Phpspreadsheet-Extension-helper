<?php

        require_once('PhpSpreadSheet.php');

            // pass custom style instead of style name from calling method - it must be array
            $customStyle = [
                'font' => [
                    'bold' => true,
                    'size' => 10,
                ],
                'alignment' => [
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER, // use full path
                ]
            ];

            // to make width flexible OR pass simple array to set header
            $headings = [
                ['col_name'=>'First column name to adjust width',
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


            // for multisheet --start
            $excel = new PhpSpreadSheet('Pragnesh', true);
            $sheets = ['one', 'two', 'three'];
            foreach ($sheets as $key=>$sheet){
                $excel->createNewSheet($sheet, true, $key);
                $excel->setTitle("Hello sheet name : ".$sheet, count($headings), $excel->lastCellAddress(), "TABLE_TITLE");
                $excel->setRowGap(1);
                $excel->setHeading($headings,"",'TABLE_HEAD');
                $excel->setArrayData($arrayData);

                // check custom style appled
                $excel->setHeading($totaling,"",$customStyle);
            }
            $excel->write();
            die('DONE');