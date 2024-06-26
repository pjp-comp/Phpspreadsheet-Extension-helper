# phpspreadsheet_simpler

Here https://github.com/PHPOffice/PhpSpreadsheet is being used and made simpler for genaral purpose use.

prerequisite : phpspreadsheet
checkout the installation steps : https://github.com/PHPOffice/PhpSpreadsheet

Just pass data in array and get excel downloaded.
Here you can extend functionalities.

What you can do using this extension code ?
1. You can make a excel sheet in a very simpler way. Just create an object and pass data as given in example.
2. Simple way to format table data with different styles given as heading. Also you can provide width as per your requirement.
3. Store data by creating multisheets. 

            // Create simple sheet and store data
            $headings = ['Heading1','Heading2','Heading3','Heading4'];
            $arrayData = [
                ['Q1',   12,   15,   21],
                ['Q2',   56,   73,   86],
                ['Q3',   52,   61,   69],
                ['Q4',   30,   32,    null],
            ];
      
            $excel->setProperty("Pragnesh", "generate excel");


            $excel->setTitle("Hello dudne nice to meet you", 4, "B2", "TABLE_TITLE");
            $excel->setHeading($headings,"",'TABLE_HEAD');
            $excel->setArrayData($arrayData);
            $excel->generateExcelFromHtml();
            $excel->write();


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




            // create multisheet excel : here it will create three sheets and store data in each.
   
            $excel = new PhpSpreadSheet('Pragnesh', true);
            $sheets = ['one', 'two', 'three'];
            foreach ($sheets as $key=>$sheet){
                $excel->createNewSheet($sheet, true, $key);
                $excel->setHeading($headings,"",'TABLE_HEAD');
                $excel->setArrayData($arrayData);
            }
   
            // for single sheet : and provide styles ( stylish heading with flexible column width )
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
            // $excel->generateExcelFromHtml();
            $excel->setRowGap(4);
            $excel->setTitle("Hello again", 4, $excel->lastCellAddress(), "TABLE_TITLE");
            $excel->setHeading($headings,"",'TABLE_HEAD');
            $excel->setArrayData($arrayData);
            $excel->setHeading($totaling,"",'TABLE_HEAD_LIGHT');

            $excel->write();
          

