<?php
/*
*Laravel框架项目,将此文件放入App\Common\Excel
*如果是其他框架，请自行修改配置（命名空间、文件路径等），主要导出逻辑不影响
*/
namespace App\Common\Excel;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;

class ExportExcel
{
    /**文件存放位置，可手动修改*/
    const DIR = 'app/excel/detail/';

    // 设置excel表默认样式
    protected static $styleArray = [
        //格式
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        ],
        //字体
        'font' => [
            'name' => '宋体',
            'size' => 11,
        ],
        //边框
        'borders' => [
            'allBorders' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN //细边框
            ],
        ],
    ];

    //sample title--示例title
    protected static $exportTitle = [
        [
            'sheet' => '白血病',//sheetName
            'title' => '手动勾选数据X条', //titleName
            'children' => [
                [
                    'key' => "clinical",
                    'name' => '临床信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "clinical_wbc",
                            'name' => '血常规-WBC',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_hb",
                            'name' => '血常规-HB',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_rbc",
                            'name' => '血常规-RBC',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_plt",
                            'name' => '血常规-PLT',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_aptt",
                            'name' => '出凝血指标-活化部分凝血活酶时间（APTT）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_pt",
                            'name' => '出凝血指标-凝血酶原时间（PT）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_tt",
                            'name' => '出凝血指标-凝血酶时间（TT）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fib",
                            'name' => '出凝血指标-纤维蛋白原（FIB）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fdp",
                            'name' => '出凝血指标-纤维蛋白（原）降解产物（FDP）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "biochemistry",
                            'name' => '生化指标',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "alt",
                                    'name' => 'alt',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ast",
                                    'name' => 'ast',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "akp",
                                    'name' => 'akp',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ggt",
                                    'name' => 'ggt',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ldh",
                                    'name' => 'ldh',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "tbil",
                                    'name' => 'tbil',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "dbil",
                                    'name' => 'dbil',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "alb",
                                    'name' => 'alb',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "bun",
                                    'name' => 'bun',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "crea",
                                    'name' => 'crea',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ua",
                                    'name' => 'ua',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "k",
                                    'name' => 'k',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "na",
                                    'name' => 'na',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "cl",
                                    'name' => 'cl',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ca",
                                    'name' => 'ca',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                            ],
                        ],
                    ],
                ],
                [
                    'key' => "micm",
                    'name' => 'MICM分型',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "bone_marrow_cell",
                            'name' => '外周血原始细胞比例',
                            'is_start' => false,
                            'children' => false,
                        ],
                        [
                            'key' => "immune",
                            'name' => '免疫分型',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "list",
                                    'name' => 'B系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'CD19标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD20标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD22标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD79a标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'Cyμ标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD10标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD34标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                                [
                                    'key' => "list",
                                    'name' => 'T系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'cCD3标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD4标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD8标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD5标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD7标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD1a标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD2标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD56标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                            ]
                        ],
                    ]
                ],
                [
                    'key' => "transfer",
                    'name' => '转归信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "transfer_const",
                            'name' => '全程治疗费用',
                            'is_start' => false,
                            'children' => false,
                        ],
                    ]
                ],
            ],
        ],
        [
            'sheet' => '白血病',//sheetName
            'title' => '手动勾选数据X条', //titleName
            'children' => [
                [
                    'key' => "clinical",
                    'name' => '临床信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "clinical_wbc",
                            'name' => '血常规-WBC',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_hb",
                            'name' => '血常规-HB',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_rbc",
                            'name' => '血常规-RBC',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_plt",
                            'name' => '血常规-PLT',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_aptt",
                            'name' => '出凝血指标-活化部分凝血活酶时间（APTT）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_pt",
                            'name' => '出凝血指标-凝血酶原时间（PT）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_tt",
                            'name' => '出凝血指标-凝血酶时间（TT）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fib",
                            'name' => '出凝血指标-纤维蛋白原（FIB）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fdp",
                            'name' => '出凝血指标-纤维蛋白（原）降解产物（FDP）',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "biochemistry",
                            'name' => '生化指标',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "alt",
                                    'name' => 'alt',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ast",
                                    'name' => 'ast',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "akp",
                                    'name' => 'akp',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ggt",
                                    'name' => 'ggt',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ldh",
                                    'name' => 'ldh',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "tbil",
                                    'name' => 'tbil',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "dbil",
                                    'name' => 'dbil',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "alb",
                                    'name' => 'alb',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "bun",
                                    'name' => 'bun',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "crea",
                                    'name' => 'crea',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ua",
                                    'name' => 'ua',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "k",
                                    'name' => 'k',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "na",
                                    'name' => 'na',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "cl",
                                    'name' => 'cl',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ca",
                                    'name' => 'ca',
                                    'is_start' => false,
                                    'children' => false,
                                ],
                            ],
                        ],
                    ],
                ],
                [
                    'key' => "micm",
                    'name' => 'MICM分型',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "bone_marrow_cell",
                            'name' => '外周血原始细胞比例',
                            'is_start' => false,
                            'children' => false,
                        ],
                        [
                            'key' => "immune",
                            'name' => '免疫分型',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "list",
                                    'name' => 'B系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'CD19标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD20标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD22标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD79a标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'Cyμ标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD10标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD34标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                                [
                                    'key' => "list",
                                    'name' => 'T系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'cCD3标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD4标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD8标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD5标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD7标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD1a标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD2标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD56标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                            ]
                        ],
                    ]
                ],
                [
                    'key' => "transfer",
                    'name' => '转归信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "transfer_const",
                            'name' => '全程治疗费用',
                            'is_start' => false,
                            'children' => false,
                        ],
                    ]
                ],
            ],
        ],
    ];
    //sample data--示例数据
    protected static $exportData = [
        [
            [
                [
                    'key' => "clinical",
                    'name' => '临床信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "clinical_wbc",
                            'name' => '血常规-WBC',
                            'is_start' => false,
                            'value' => ['1'],
                            'children' => false
                        ],
                        [
                            'key' => "clinical_hb",
                            'name' => '血常规-HB',
                            'value' => '2',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_rbc",
                            'name' => '血常规-RBC',
                            'value' => '3',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_plt",
                            'name' => '血常规-PLT',
                            'value' => '4',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_aptt",
                            'name' => '出凝血指标-活化部分凝血活酶时间（APTT）',
                            'value' => '5',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_pt",
                            'name' => '出凝血指标-凝血酶原时间（PT）',
                            'value' => '6',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_tt",
                            'name' => '出凝血指标-凝血酶时间（TT）',
                            'value' => '7',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fib",
                            'name' => '出凝血指标-纤维蛋白原（FIB）',
                            'value' => '8',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fdp",
                            'name' => '出凝血指标-纤维蛋白（原）降解产物（FDP）',
                            'value' => '9',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "biochemistry",
                            'name' => '生化指标',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "alt",
                                    'name' => 'alt',
                                    'value' => [1,2,3],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ast",
                                    'name' => 'ast',
                                    'value' => [4,5,6],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "akp",
                                    'name' => 'akp',
                                    'value' => [7,8,9],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ggt",
                                    'name' => 'ggt',
                                    'value' => 10,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ldh",
                                    'name' => 'ldh',
                                    'value' => [10,9,8,7,6],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "tbil",
                                    'name' => 'tbil',
                                    'value' => [5,4,3,2,1],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "dbil",
                                    'name' => 'dbil',
                                    'value' => [2],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "alb",
                                    'name' => 'alb',
                                    'value' => [4,5],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "bun",
                                    'name' => 'bun',
                                    'value' => 3,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "crea",
                                    'name' => 'crea',
                                    'value' => 6,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ua",
                                    'name' => 'ua',
                                    'value' => 9,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "k",
                                    'name' => 'k',
                                    'value' => 16,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "na",
                                    'name' => 'na',
                                    'value' => 18,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "cl",
                                    'name' => 'cl',
                                    'value' => 20,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ca",
                                    'name' => 'ca',
                                    'value' => 22,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                            ],
                        ],
                    ],
                ],
                [
                    'key' => "micm",
                    'name' => 'MICM分型',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "bone_marrow_cell",
                            'name' => '外周血原始细胞比例',
                            'value' => 11.2,
                            'is_start' => false,
                            'children' => false,
                        ],
                        [
                            'key' => "immune",
                            'name' => '免疫分型',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "list",
                                    'name' => 'B系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'CD19标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD20标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD22标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD79a标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'Cyμ标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD10标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD34标志数值比例',
                                            'value' => 22.3,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                                [
                                    'key' => "list",
                                    'name' => 'T系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'cCD3标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD4标志数值比例',
                                            'value' => 22.5,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD8标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD5标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD7标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD1a标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD2标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD56标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'value' => 22.6,
                                            'is_unit' => '%',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                            ]
                        ],
                    ]
                ],
                [
                    'key' => "transfer",
                    'name' => '转归信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "transfer_const",
                            'name' => '全程治疗费用',
                            'value' => 9999,
                            'is_unit' => '元',
                            'is_start' => false,
                            'children' => false,
                        ],
                    ]
                ],
            ],
            [
                [
                    'key' => "clinical",
                    'name' => '临床信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "clinical_wbc",
                            'name' => '血常规-WBC',
                            'is_start' => false,
                            'value' => ['1'],
                            'children' => false
                        ],
                        [
                            'key' => "clinical_hb",
                            'name' => '血常规-HB',
                            'value' => '2',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_rbc",
                            'name' => '血常规-RBC',
                            'value' => '3',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_plt",
                            'name' => '血常规-PLT',
                            'value' => '4',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_aptt",
                            'name' => '出凝血指标-活化部分凝血活酶时间（APTT）',
                            'value' => '5',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_pt",
                            'name' => '出凝血指标-凝血酶原时间（PT）',
                            'value' => '6',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_tt",
                            'name' => '出凝血指标-凝血酶时间（TT）',
                            'value' => '7',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fib",
                            'name' => '出凝血指标-纤维蛋白原（FIB）',
                            'value' => '8',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fdp",
                            'name' => '出凝血指标-纤维蛋白（原）降解产物（FDP）',
                            'value' => '9',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "biochemistry",
                            'name' => '生化指标',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "alt",
                                    'name' => 'alt',
                                    'value' => [1,2,3],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ast",
                                    'name' => 'ast',
                                    'value' => [4,5,6],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "akp",
                                    'name' => 'akp',
                                    'value' => [7,8,9],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ggt",
                                    'name' => 'ggt',
                                    'value' => 10,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ldh",
                                    'name' => 'ldh',
                                    'value' => [10,9,8,7,6],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "tbil",
                                    'name' => 'tbil',
                                    'value' => [5,4,3,2,1],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "dbil",
                                    'name' => 'dbil',
                                    'value' => [2],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "alb",
                                    'name' => 'alb',
                                    'value' => [4,5],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "bun",
                                    'name' => 'bun',
                                    'value' => 3,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "crea",
                                    'name' => 'crea',
                                    'value' => 6,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ua",
                                    'name' => 'ua',
                                    'value' => 9,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "k",
                                    'name' => 'k',
                                    'value' => 16,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "na",
                                    'name' => 'na',
                                    'value' => 18,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "cl",
                                    'name' => 'cl',
                                    'value' => 20,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ca",
                                    'name' => 'ca',
                                    'value' => 22,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                            ],
                        ],
                    ],
                ],
                [
                    'key' => "micm",
                    'name' => 'MICM分型',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "bone_marrow_cell",
                            'name' => '外周血原始细胞比例',
                            'value' => 11.2,
                            'is_start' => false,
                            'children' => false,
                        ],
                        [
                            'key' => "immune",
                            'name' => '免疫分型',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "list",
                                    'name' => 'B系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'CD19标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD20标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD22标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD79a标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'Cyμ标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD10标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD34标志数值比例',
                                            'value' => 22.3,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                                [
                                    'key' => "list",
                                    'name' => 'T系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'cCD3标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD4标志数值比例',
                                            'value' => 22.5,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD8标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD5标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD7标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD1a标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD2标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD56标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'value' => 22.6,
                                            'is_unit' => '%',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                            ]
                        ],
                    ]
                ],
                [
                    'key' => "transfer",
                    'name' => '转归信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "transfer_const",
                            'name' => '全程治疗费用',
                            'value' => 9999,
                            'is_unit' => '元',
                            'is_start' => false,
                            'children' => false,
                        ],
                    ]
                ],
            ],
        ],
        [
            [
                [
                    'key' => "clinical",
                    'name' => '临床信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "clinical_wbc",
                            'name' => '血常规-WBC',
                            'is_start' => false,
                            'value' => ['1'],
                            'children' => false
                        ],
                        [
                            'key' => "clinical_hb",
                            'name' => '血常规-HB',
                            'value' => '2',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_rbc",
                            'name' => '血常规-RBC',
                            'value' => '3',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_plt",
                            'name' => '血常规-PLT',
                            'value' => '4',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_aptt",
                            'name' => '出凝血指标-活化部分凝血活酶时间（APTT）',
                            'value' => '5',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_pt",
                            'name' => '出凝血指标-凝血酶原时间（PT）',
                            'value' => '6',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_tt",
                            'name' => '出凝血指标-凝血酶时间（TT）',
                            'value' => '7',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fib",
                            'name' => '出凝血指标-纤维蛋白原（FIB）',
                            'value' => '8',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fdp",
                            'name' => '出凝血指标-纤维蛋白（原）降解产物（FDP）',
                            'value' => '9',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "biochemistry",
                            'name' => '生化指标',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "alt",
                                    'name' => 'alt',
                                    'value' => [1,2,3],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ast",
                                    'name' => 'ast',
                                    'value' => [4,5,6],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "akp",
                                    'name' => 'akp',
                                    'value' => [7,8,9],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ggt",
                                    'name' => 'ggt',
                                    'value' => 10,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ldh",
                                    'name' => 'ldh',
                                    'value' => [10,9,8,7,6],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "tbil",
                                    'name' => 'tbil',
                                    'value' => [5,4,3,2,1],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "dbil",
                                    'name' => 'dbil',
                                    'value' => [2],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "alb",
                                    'name' => 'alb',
                                    'value' => [4,5],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "bun",
                                    'name' => 'bun',
                                    'value' => 3,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "crea",
                                    'name' => 'crea',
                                    'value' => 6,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ua",
                                    'name' => 'ua',
                                    'value' => 9,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "k",
                                    'name' => 'k',
                                    'value' => 16,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "na",
                                    'name' => 'na',
                                    'value' => 18,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "cl",
                                    'name' => 'cl',
                                    'value' => 20,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ca",
                                    'name' => 'ca',
                                    'value' => 22,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                            ],
                        ],
                    ],
                ],
                [
                    'key' => "micm",
                    'name' => 'MICM分型',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "bone_marrow_cell",
                            'name' => '外周血原始细胞比例',
                            'value' => 11.2,
                            'is_start' => false,
                            'children' => false,
                        ],
                        [
                            'key' => "immune",
                            'name' => '免疫分型',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "list",
                                    'name' => 'B系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'CD19标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD20标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD22标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD79a标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'Cyμ标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD10标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD34标志数值比例',
                                            'value' => 22.3,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                                [
                                    'key' => "list",
                                    'name' => 'T系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'cCD3标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD4标志数值比例',
                                            'value' => 22.5,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD8标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD5标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD7标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD1a标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD2标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD56标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'value' => 22.6,
                                            'is_unit' => '%',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                            ]
                        ],
                    ]
                ],
                [
                    'key' => "transfer",
                    'name' => '转归信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "transfer_const",
                            'name' => '全程治疗费用',
                            'value' => 9999,
                            'is_unit' => '元',
                            'is_start' => false,
                            'children' => false,
                        ],
                    ]
                ],
            ],
            [
                [
                    'key' => "clinical",
                    'name' => '临床信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "clinical_wbc",
                            'name' => '血常规-WBC',
                            'is_start' => false,
                            'value' => ['1'],
                            'children' => false
                        ],
                        [
                            'key' => "clinical_hb",
                            'name' => '血常规-HB',
                            'value' => '2',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_rbc",
                            'name' => '血常规-RBC',
                            'value' => '3',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_plt",
                            'name' => '血常规-PLT',
                            'value' => '4',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_aptt",
                            'name' => '出凝血指标-活化部分凝血活酶时间（APTT）',
                            'value' => '5',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_pt",
                            'name' => '出凝血指标-凝血酶原时间（PT）',
                            'value' => '6',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_tt",
                            'name' => '出凝血指标-凝血酶时间（TT）',
                            'value' => '7',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fib",
                            'name' => '出凝血指标-纤维蛋白原（FIB）',
                            'value' => '8',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "clinical_fdp",
                            'name' => '出凝血指标-纤维蛋白（原）降解产物（FDP）',
                            'value' => '9',
                            'is_start' => false,
                            'children' => false
                        ],
                        [
                            'key' => "biochemistry",
                            'name' => '生化指标',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "alt",
                                    'name' => 'alt',
                                    'value' => [1,2,3],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ast",
                                    'name' => 'ast',
                                    'value' => [4,5,6],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "akp",
                                    'name' => 'akp',
                                    'value' => [7,8,9],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ggt",
                                    'name' => 'ggt',
                                    'value' => 10,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ldh",
                                    'name' => 'ldh',
                                    'value' => [10,9,8,7,6],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "tbil",
                                    'name' => 'tbil',
                                    'value' => [5,4,3,2,1],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "dbil",
                                    'name' => 'dbil',
                                    'value' => [2],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "alb",
                                    'name' => 'alb',
                                    'value' => [4,5],
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "bun",
                                    'name' => 'bun',
                                    'value' => 3,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "crea",
                                    'name' => 'crea',
                                    'value' => 6,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ua",
                                    'name' => 'ua',
                                    'value' => 9,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "k",
                                    'name' => 'k',
                                    'value' => 16,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "na",
                                    'name' => 'na',
                                    'value' => 18,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "cl",
                                    'name' => 'cl',
                                    'value' => 20,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                                [
                                    'key' => "ca",
                                    'name' => 'ca',
                                    'value' => 22,
                                    'is_start' => false,
                                    'children' => false,
                                ],
                            ],
                        ],
                    ],
                ],
                [
                    'key' => "micm",
                    'name' => 'MICM分型',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "bone_marrow_cell",
                            'name' => '外周血原始细胞比例',
                            'value' => 11.2,
                            'is_start' => false,
                            'children' => false,
                        ],
                        [
                            'key' => "immune",
                            'name' => '免疫分型',
                            'is_start' => false,
                            'children' => [
                                [
                                    'key' => "list",
                                    'name' => 'B系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'CD19标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD20标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD22标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD79a标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'Cyμ标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD10标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD34标志数值比例',
                                            'value' => 22.3,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                                [
                                    'key' => "list",
                                    'name' => 'T系',
                                    'is_start' => false,
                                    'children' => [
                                        [
                                            'key' => "value",
                                            'name' => 'cCD3标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD4标志数值比例',
                                            'value' => 22.5,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD8标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD5标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD7标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD1a标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD2标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'CD56标志数值比例',
                                            'value' => 22.1,
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                        [
                                            'key' => "value",
                                            'name' => 'TdT标志数值比例',
                                            'value' => 22.6,
                                            'is_unit' => '%',
                                            'is_start' => false,
                                            'children' => false,
                                        ],
                                    ],
                                ],
                            ]
                        ],
                    ]
                ],
                [
                    'key' => "transfer",
                    'name' => '转归信息',
                    'is_start' => false,
                    'children' => [
                        [
                            'key' => "transfer_const",
                            'name' => '全程治疗费用',
                            'value' => 9999,
                            'is_unit' => '元',
                            'is_start' => false,
                            'children' => false,
                        ],
                    ]
                ],
            ],
        ]
    ];
    //标题的列数据坐标
    protected static $titleRowCoordinate = [];
    //标题的行数据坐标
    protected static $titleColumnCoordinate = [];
    
    protected static $dataKey = [];

    //每个sheet标题所到最大行
    protected static $exportColumn;
    //每个sheet标题所到最大列
    protected static $exportRow;

    const EXT = 'Xlsx';

    /**
     * Desc:excel基础路径
     *
     * @return string
     */
    private static function baseDir(): string
    {
        /**文件存放位置，可手动修改*/
        $baseDir = Config('filesystems.disks.public.root') . self::DIR;
        if (!is_dir($baseDir)) {
            mkdir($baseDir, 0777, true);
        }
        return $baseDir;
    }

    /**
     * Desc：处理导出EXCEL,数据结构要和上面的两个示例一样
     * 理论上可以有无限个维度的标题和数据，目前只测了5级（标题和数据）及以下维度的数据。
     * 在excel中 column指的是：单元格的数字列坐标；row指的是：单元格的数字行坐标
     * 在代码中  column指的是：单元格的数字行坐标；row指的是：单元格的数字列坐标
     *
     * @param array $exportTitle
     * @param array $exportData
     * @param array $fileInfo
     * @param string $ext
     *
     * @return array
     * User: 慕容明月
     */
    public static function explorExcel($exportTitle = [], $exportData = [], $fileInfo = [], $ext = self::EXT): array
    {
        //设置无限时间
        set_time_limit(0);
        $baseFileName = $fileInfo['file_name'] ?? '示例文件';
        $uniqueName = md5(uniqid($baseFileName));
        /**文件存放位置，可手动修改*/
        $signBaseFileName = self::baseDir() . $uniqueName . '.' . $ext;
        $filePath = self::DIR . $uniqueName . '.' . $ext;
        $spreadSheet = new Spreadsheet();
        !empty($exportTitle) && self::$exportTitle = $exportTitle;
        !empty($exportData) && self::$exportData = $exportData;
        foreach (self::$exportTitle as $key => $value) {
            self::opTitleSheet($spreadSheet, $key, $value);
        }
        foreach (self::$exportData as $key => $value) {
            self::opDataSheet($spreadSheet, $key, $value);
        }
        $spreadSheet->getProperties()
            ->setCreator("Mr.Xu.Rui")   //作者
            ->setLastModifiedBy("Mr.Xu.Rui")    //最后修改者
            ->setTitle($baseFileName)   //标题
            ->setDescription("导出统计相关信息")  //描述
            ->setKeywords("慕容明月");    //关键字
        $IWriter = IOFactory::createWriter($spreadSheet, $ext);
        $IWriter->save($signBaseFileName);
        return [
            'uniqueName' => $uniqueName,
            'ext' => strtolower($ext),
            'pathFile' => $filePath,
            /**文件存放位置，可手动修改*/
            'urlFile' => self::objectUrl($filePath),
            'fileName' => $baseFileName,
            'localPath' => $signBaseFileName
        ];
    }

    /**
     * Desc：标题标头的整理
     *
     * @param $spreadSheet
     * @param $key
     * @param $value
     *
     * @return array|int[]
     * User: 慕容明月
     */
    protected static function opTitleSheet($spreadSheet, $key, $value)
    {
        $spreadSheet->createSheet($key);//创建sheet
        $sheet = $spreadSheet->setActiveSheetIndex($key);
        $spreadSheet->getDefaultStyle()->applyFromArray(self::$styleArray);
        !empty($value['sheet']) && $sheet->setTitle($value['sheet']);
        $children = $value['children'] ?? [];
        $title = $value['title'] ?? '';
        $rowBaseNum = empty($title) ? 0 : 1;
        //在excel中--column指的是：单元格的数字列坐标；row指的是：单元格的数字行坐标
        $counts = count($children);
        $baseRow = 1;
        $baseColumn = 1 + $rowBaseNum;
        //放入标题数据
        list($maxNextRow, $totalColumn) = self::opTitleChildren($key, $children, $sheet, $baseRow, $baseColumn);
        self::$dataKey[$key] = self::$titleColumnCoordinate[$key];
        self::$exportColumn[$key] = $totalColumn;
        self::$exportRow[$key] = $totalRow = $maxNextRow - 1;
        //合并带标题的特殊第一行
        if ($rowBaseNum) {
            $sheet->mergeCellsByColumnAndRow($rowBaseNum, $rowBaseNum, $totalRow, $rowBaseNum);
            $objRichText = new RichText();
            $objPayableTitle = $objRichText->createTextRun($title);
            $objPayableTitle->getFont()->applyFromArray(self::$styleArray['font']);
            $sheet->setCellValueByColumnAndRow($rowBaseNum, $rowBaseNum, $objRichText);
        }
        //合并其他行和列
        self::opMergeColumnRow($sheet, $baseRow, $totalRow, $baseColumn, $totalColumn, $key);
        return [$totalRow, $totalColumn];
    }

    /**
     * Desc：合并其他行和列
     *
     * @param $sheet
     * @param $baseRow
     * @param $totalRow
     * @param $baseColumn
     * @param $totalColumn
     * @param $sheetKey
     *
     * @return void
     * User: 慕容明月
     */
    protected static function opMergeColumnRow($sheet, $baseRow, $totalRow, $baseColumn, $totalColumn, $sheetKey)
    {
        //合并其他列
        foreach(self::$titleRowCoordinate[$sheetKey] as $c => $value) {
            //初始化
            $sr = $baseRow;
            foreach ($value as $r => $item) {
                $differenceRow = $r - $sr;
                if ($differenceRow > 1) {
                    //这列的右边列上面没有数据才可以往右边合并
                    $srb = empty(self::$titleColumnCoordinate[$sheetKey][$sr + 1]);
                    if ($srb) {
                        $sheet->mergeCellsByColumnAndRow($sr, $c, $r - 1, $c);
                    } else {
                        $srb = true;
                        foreach (self::$titleColumnCoordinate[$sheetKey][$sr + 1] as $rc => $rcItem) {
                            if ($c > $rc) {
                                $srb = false;
                                break;
                            }
                        }
                        if ($srb) {
                            $sheet->mergeCellsByColumnAndRow($sr, $c, $r - 1, $c);
                        }
                    }
                }
                $sr = $r;
            }
            //最后一个列
            if ($sr < $totalRow) {
                //这列的右边列上面没有数据才可以往右边合并
                $srb = empty(self::$titleColumnCoordinate[$sheetKey][$sr + 1]);
                if ($srb) {
                    $sheet->mergeCellsByColumnAndRow($sr, $c, $totalRow - 1, $c);
                } else {
                    $srb = true;
                    foreach (self::$titleColumnCoordinate[$sheetKey][$sr + 1] as $rc => $rcItem) {
                        if ($c > $rc) {
                            $srb = false;
                            break;
                        }
                    }
                    if ($srb) {
                        $sheet->mergeCellsByColumnAndRow($sr, $c, $totalRow - 1, $c);
                    }
                }
            }
        }
        //合并其他行
        foreach (self::$titleColumnCoordinate[$sheetKey] as $r => $value) {
            //初始化
            $sc = $baseColumn;
            foreach ($value as $c => $item) {
                if ($sc != $baseColumn) {
                    $differenceColumn = $r - $sc;
                    if ($differenceColumn > 1) {
                        $sheet->mergeCellsByColumnAndRow($r, $sc, $r, $c - 1);
                    }
                }
                $sc = $c;
            }
            //合并到最后一行
            if ($sc < $totalColumn) {
                $sheet->mergeCellsByColumnAndRow($r, $sc, $r, $totalColumn);
            }
        }
    }

    /**
     * Desc：将标题放入excel数据框内
     *
     * @param $sheetKey
     * @param $children
     * @param $sheet
     * @param int $row
     * @param int $column
     * @param int $selfColumn
     * @param int $totalColumn
     *
     * @return array|int[]
     * User: 慕容明月
     */
    protected static function opTitleChildren($sheetKey, $children, $sheet, $row = 1, $column = 0, $selfColumn = -1, $totalColumn = 0)
    {
        if (!empty($children)) {
            $selfColumn += 1;
            foreach ($children as $key => $value) {
                $totalColumn = (($column + $selfColumn) > $totalColumn) ? ($column + $selfColumn) : $totalColumn;
                self::start($sheet, $row, $column + $selfColumn, $value, $sheetKey);
                $row += 1;
                if (!empty($value['children'])) {
                    $row -= 1;
                    list($row, $totalColumn) = self::opTitleChildren($sheetKey, $value['children'], $sheet, $row, $column, $selfColumn, $totalColumn);
                }
            }
        }
        return [
            $row,
            $totalColumn,
        ];
    }

    /**
     * Desc：处理带星标题以及放入标题值
     *
     * @param $sheet
     * @param $column
     * @param $row
     * @param $value
     * @param $sheetKey
     *
     * @return void
     * User: 慕容明月
     */
    private static function start($sheet, $row, $column, $value, $sheetKey)
    {
        $objRichText = new RichText();
        if (isset($value['is_start']) && $value['is_start'] === true) {
            $objPayableOne = $objRichText->createTextRun('*');
            $objPayableOne->getFont()->applyFromArray(self::$styleArray['font'])->setColor(new Color(Color::COLOR_RED));
        }
        $objPayableTwo = $objRichText->createTextRun($value['name']);
        $objPayableTwo->getFont()->applyFromArray(self::$styleArray['font']);
        $sheet->setCellValueByColumnAndRow($row, $column, $objRichText);
        self::$titleRowCoordinate[$sheetKey][$column][$row] = [
            'name' => $value['name'],
            'key' => $value['key'],
            'row' => $row,
            'column' => $column,
        ];
        self::$titleColumnCoordinate[$sheetKey][$row][$column] = [
            'name' => $value['name'],
            'key' => $value['key'],
            'row' => $row,
            'column' => $column,
        ];
    }

    /**
     * Desc：将数据放入excel数据框内
     *
     * @param $spreadSheet
     * @param $key
     * @param $value
     *
     * @return void
     * User: 慕容明月
     
     */
    protected static function opDataSheet($spreadSheet, $key, $value)
    {
        $sheet = $spreadSheet->setActiveSheetIndex($key);
        $spreadSheet->getDefaultStyle()->applyFromArray(self::$styleArray);
        $baseRow = 1;
        $baseColumn = $selfBaseColumn = self::$exportColumn[$key] ?? 3;
        $titles = self::$dataKey[$key];
        $needMerge = [];
        $maxColumn = [];
        foreach ($value as $i => $jv) {
            $maxColumn[$i] = 1;
            foreach ($titles as $k => $vv) {
                $v = array_pop($vv);
                $j = self::opDataChildren($jv, $v['key'] ?? '', $v['code'] ?? '');
                if (is_array($j) && count($j) > 1) {
                    //存在多个值
                    foreach ($j as $n => $m) {
                        $m === '' && $m = '--';
                        $sheet->setCellValueExplicitByColumnAndRow($k, $baseColumn + $n + 1, $m, DataType::TYPE_STRING);
                    }
                    $maxColumn[$i] = $maxColumn[$i] < count($j) ? count($j) : $maxColumn[$i];
                } elseif (is_array($j) && count($j) == 1) {
                    //存在单个值
                    foreach ($j as $n => $m) {
                        $m === '' && $m = '--';
                        $sheet->setCellValueExplicitByColumnAndRow($k, $baseColumn + $n + 1, $m, DataType::TYPE_STRING);
                    }
                    $needMerge[$i][] = $k;
                    $maxColumn[$i] = $maxColumn[$i] < count($j) ? count($j) : $maxColumn[$i];
                } else {
                    //存在单个值
                    $j === '' && $j = '--';
                    $sheet->setCellValueExplicitByColumnAndRow($k, $baseColumn + 1, $j, DataType::TYPE_STRING);
                    $needMerge[$i][] = $k;
                }
            }
            $baseColumn += $maxColumn[$i];
        }
        //合并第一列数据
        self::mergeFirstRow($sheet, $selfBaseColumn, $baseColumn);
        //合并其他列数据
        foreach ($needMerge as $i => $j) {
            //合并多行为一列
            if ($maxColumn[$i] > 1) {
                foreach ($j as $n) {
                    $sheet->mergeCellsByColumnAndRow($n, $selfBaseColumn + 1, $n, $maxColumn[$i] + $selfBaseColumn);
                }
                $selfBaseColumn += $maxColumn[$i];
            }
        }
    }

    /**
     * Desc：插入数据值
     *
     * @param $data
     * @param string $key
     * @param string $code
     *
     * @return mixed|string
     * User: 慕容明月
     */
    private static function opDataChildren($data, $key = '', $code = '')
    {
        if (is_array($data)) {
            foreach ($data as $k => $v) {
                if (isset($v['key']) && empty($v['children'])) {
                    if (!empty($code)) {
                        if ($code == $v['code'] && $key == $v['key']) {
                            return $v['value'];
                        }
                    } elseif ($key == $v['key']) {
                        return $v['value'];
                    }
                } elseif ($key == $k) {
                    return $v;
                } else {
                    $value = self::opDataChildren((isset($v['children']) && !empty($v['children'])) ? $v['children'] : $v, $key, $code);
                    if ($value) {
                        return $value;
                    }
                }
            }
        }
        return '';
    }

    /**
     * Desc：合并第一列数据
     *
     * @param $sheet
     * @param $selfBaseColumn
     * @param $baseColumn
     *
     * @return void
     * User: 慕容明月
     */
    private static function mergeFirstRow($sheet, $selfBaseColumn, $baseColumn)
    {
        //合并第一列相同数据
        $inputArray = [];
        for ($si = $selfBaseColumn + 1; $si <= $baseColumn; $si += 1) {
            $firstCellValue = $sheet->getCellByColumnAndRow(1, $si)->getValue();
            if ($firstCellValue instanceof RichText) {
                $firstCellValue = $firstCellValue->getPlainText();
            }
            $inputArray[$si] = $firstCellValue;
        }

        $result = [];
        $lastValue = null;
        $rangeStart = null;

        foreach ($inputArray as $key => $value) {
            if ($value === $lastValue) {
                // 当前值与上一个值相同，更新结束键
                $rangeEnd = $key;
            } else {
                // 当前值与上一个值不同，检查是否有范围
                if ($rangeStart !== null) {
                    $result[] = ["start" => $rangeStart, "end" => $key - 1];
                }

                // 设置新的起始键
                $rangeStart = $key;
            }

            // 更新上一个值
            $lastValue = $value;
        }

        // 检查最后一个范围是否存在
        if ($rangeStart !== null) {
            $result[] = ["start" => $rangeStart, "end" => $key];
        }

        //合并单元格
        foreach ($result as $kk => $vv) {
            if ($vv['start'] != $vv['end']) {
                $sheet->mergeCellsByColumnAndRow(1, $vv['start'], 1, $vv['end']);
            }
        }
    }

    /**
     * Desc：完整文件url
     *
     * @param $path
     *
     * @return mixed
     * User: 慕容明月
     */
    public static function objectUrl($path)
    {
        return Storage::disk('public')->url($path);
    }
}

//how to use ?
/**
 * 
 * ExportExcel::explorExcel($exportTitle = [], $exportData = [], $fileInfo = [], $ext = self::EXT)
 * 
 */
