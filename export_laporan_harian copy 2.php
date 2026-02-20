<?php
declare(strict_types=1);

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

require __DIR__ . '/vendor/autoload.php';

/* ================= HEADERS ================= */
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="rekap_harian_ais.xlsx"');
header('Cache-Control: max-age=0');

/* ================= LOAD TEMPLATE ================= */
$templatePath = realpath(__DIR__ . '/templates/laporan_template1.xlsx');

if ($templatePath === false) {
    http_response_code(500);
    exit('Template Excel tidak ditemukan');
}

$spreadsheet = IOFactory::load($templatePath);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('DATA KAPAL');

/* ================= READ INPUT JSON ================= */
$raw = file_get_contents('php://input');
$data = json_decode($raw, true);

if (!is_array($data)) {
    http_response_code(400);
    exit('Data JSON tidak valid');
}

/* ================= DATE EXPORT ================= */
$dateExport = date('d-m-Y');

/* ================= START ROW ================= */
$startRow = 4;
$currentRow = $startRow;

/* ================= COUNT FLAGS ================= */
$flagCount = [];
foreach ($data as $row) {
    // Skip baris yang tidak memiliki MMSI atau nama kapal (indikator data kosong/header)
    if (empty($row['mmsi']) && empty($row['name'])) {
        continue;
    }
    
    $flag = trim($row['flag'] ?? '');
    
    // Skip jika flag benar-benar kosong
    if ($flag === '') {
        continue;
    }
    
    if (!isset($flagCount[$flag])) {
        $flagCount[$flag] = 0;
    }
    $flagCount[$flag]++;
}

// Sort by flag name
ksort($flagCount);

/* ================= WRITE DATA TO SHEET 1 ================= */
$no = 1;

foreach ($data as $row) {
    // Skip baris kosong/header
    if (empty($row['mmsi']) && empty($row['name'])) {
        continue;
    }
    
    $sheet->setCellValue("A{$currentRow}", $no++);
    $sheet->setCellValue("B{$currentRow}", $dateExport);
    $sheet->setCellValue("C{$currentRow}", $row['name'] ?? '');
    $sheet->setCellValue("D{$currentRow}", $row['callsign'] ?? '');
    $sheet->setCellValue("E{$currentRow}", $row['mmsi'] ?? '');
    $sheet->setCellValue("F{$currentRow}", $row['course'] ?? '');
    $sheet->setCellValue("G{$currentRow}", $row['speed'] ?? '');
    $sheet->setCellValue("H{$currentRow}", $row['imo'] ?? '');
    $sheet->setCellValue("I{$currentRow}", $row['flag'] ?? '');
    $sheet->setCellValue("J{$currentRow}", $row['destination'] ?? '');

    // Latitude
    $sheet->setCellValue("K{$currentRow}", $row['latitude1'] ?? '');
    $sheet->setCellValue("L{$currentRow}", $row['latitude2'] ?? '');
    $sheet->setCellValue("M{$currentRow}", $row['latitude3'] ?? '');
    $sheet->setCellValue("N{$currentRow}", $row['latitude4'] ?? '');

    // Longitude
    $sheet->setCellValue("O{$currentRow}", $row['longitude1'] ?? '');
    $sheet->setCellValue("P{$currentRow}", $row['longitude2'] ?? '');
    $sheet->setCellValue("Q{$currentRow}", $row['longitude3'] ?? '');

    // Status & Type
    $sheet->setCellValue("R{$currentRow}", $row['status'] ?? '');
    $sheet->setCellValue("S{$currentRow}", $row['vessel_type'] ?? '');
    
    /* ================= APPLY BACKGROUND COLORS, CENTER ALIGNMENT & BORDER ================= */
    
    // Helper function untuk apply background color dengan border dan center alignment
    $applyStyle = function($cell, $bgColor) use ($sheet) {
        $sheet->getStyle($cell)->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => $bgColor],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'font' => [
                'color' => ['rgb' => '000000'], // Black text
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ]);
    };
    
    // Apply colors based on column groups
    $applyStyle("A{$currentRow}", 'FFFFFF');    // No - White
    $applyStyle("B{$currentRow}", 'EEECE0');    // Date - Cream
    $applyStyle("C{$currentRow}", 'FFE8D7');    // Ship Name - Light Orange
    $applyStyle("D{$currentRow}", 'FFE8D7');    // Call Sign - Light Orange
    $applyStyle("E{$currentRow}", 'C0DAF3');    // MMSI - Light Blue
    $applyStyle("F{$currentRow}", 'C0DAF3');    // Course - Light Blue
    $applyStyle("G{$currentRow}", 'FFE8D7');    // Speed - Light Orange
    $applyStyle("H{$currentRow}", 'C0DAF3');    // IMO - Light Blue
    $applyStyle("I{$currentRow}", 'C0DAF3');    // Flag - Light Blue
    $applyStyle("J{$currentRow}", 'FFE8D7');    // Destination - Light Orange
    
    // Latitude columns - Purple
    $applyStyle("K{$currentRow}", 'E5DFED');
    $applyStyle("L{$currentRow}", 'E5DFED');
    $applyStyle("M{$currentRow}", 'E5DFED');
    $applyStyle("N{$currentRow}", 'E5DFED');
    
    // Longitude columns - Pink
    $applyStyle("O{$currentRow}", 'F7DBDA');
    $applyStyle("P{$currentRow}", 'F7DBDA');
    $applyStyle("Q{$currentRow}", 'F7DBDA');
    
    $applyStyle("R{$currentRow}", 'FFE8D7');    // Status - Light Orange
    $applyStyle("S{$currentRow}", 'C0DAF3');    // Type - Light Blue

    $currentRow++;
}

/* ================= CREATE SHEET 2: BENDERA KAPAL (DENGAN FORMULA) ================= */
$flagSheet = $spreadsheet->createSheet();
$flagSheet->setTitle('BENDERA KAPAL');

// Headers
$flagSheet->setCellValue('A4', 'Row Labels');
$flagSheet->setCellValue('B4', 'Count of Flag');

$flagSheet->getStyle('A4:B4')->applyFromArray([
    'font' => ['bold' => true],
    'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'startColor' => ['rgb' => 'D9E1F2']
    ]
]);

// DAFTAR LENGKAP bendera kapal yang paling umum di dunia maritim
$allPossibleFlags = [
    'Antigua and Barbuda',
    'Australia',
    'Bahamas',
    'Bangladesh',
    'Barbados',
    'Belgium',
    'Belize',
    'Brazil',
    'Brunei',
    'Cambodia',
    'Canada',
    'Cayman Islands',
    'China',
    'Cook Islands',
    'Cyprus',
    'Denmark',
    'Egypt',
    'France',
    'Germany',
    'Greece',
    'Hong Kong',
    'India',
    'Indonesia',
    'Iran',
    'Ireland',
    'Isle of Man',
    'Israel',
    'Italy',
    'Japan',
    'Jordan',
    'Kiribati',
    'Kuwait',
    'Liberia',
    'Madagascar',
    'Malaysia',
    'Maldives',
    'Malta',
    'Marshall Islands',
    'Mexico',
    'Moldova',
    'Mongolia',
    'Morocco',
    'Myanmar',
    'Netherlands',
    'New Zealand',
    'Nigeria',
    'North Korea',
    'Norway',
    'Pakistan',
    'Palau',
    'Panama',
    'Philippines',
    'Poland',
    'Portugal',
    'Qatar',
    'Romania',
    'Russia',
    'Saint Kitts and Nevis',
    'Saint Vincent and the Grenadines',
    'Samoa',
    'Saudi Arabia',
    'Sierra Leone',
    'Singapore',
    'South Africa',
    'South Korea',
    'Spain',
    'Sri Lanka',
    'Sweden',
    'Switzerland',
    'Taiwan',
    'Tanzania',
    'Thailand',
    'Togo',
    'Tonga',
    'Trinidad and Tobago',
    'Tunisia',
    'Turkey',
    'Tuvalu',
    'United Arab Emirates',
    'United Kingdom',
    'United States',
    'Vanuatu',
    'Venezuela',
    'Vietnam',
];

sort($allPossibleFlags);

// Write dengan FORMULA (bukan nilai statis)
$flagRow = 5;
foreach ($allPossibleFlags as $flag) {
    $flagSheet->setCellValue("A{$flagRow}", $flag);
    
    // GUNAKAN FORMULA COUNTIF - akan otomatis update ketika data berubah!
    $flagSheet->setCellValue("B{$flagRow}", 
        "=COUNTIF('DATA KAPAL'!I:I,A{$flagRow})"
    );
    
    $flagRow++;
}

// Grand Total dengan SUM formula
$grandTotalRow = $flagRow;
$flagSheet->setCellValue("A{$flagRow}", 'Grand Total');
$flagSheet->setCellValue("B{$flagRow}", 
    "=SUM(B5:B" . ($flagRow - 1) . ")"
);

$flagSheet->getStyle("A{$flagRow}:B{$flagRow}")->applyFromArray([
    'font' => ['bold' => true],
    'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'startColor' => ['rgb' => 'D9E1F2']
    ]
]);

// Auto-size columns
$flagSheet->getColumnDimension('A')->setAutoSize(true);
$flagSheet->getColumnDimension('B')->setAutoSize(true);

/* ================= CREATE CHART FOR BENDERA KAPAL ================= */
$dataSeriesLabels = [
    new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, "'BENDERA KAPAL'!\$B\$4", null, 1)
];

$xAxisTickValues = [
    new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_STRING,
        "'BENDERA KAPAL'!\$A\$5:\$A\$" . ($grandTotalRow - 1),
        null,
        count($allPossibleFlags)
    )
];

$dataSeriesValues = [
    new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_NUMBER,
        "'BENDERA KAPAL'!\$B\$5:\$B\$" . ($grandTotalRow - 1),
        null,
        count($allPossibleFlags)
    )
];

$series = new DataSeries(
    DataSeries::TYPE_BARCHART,
    DataSeries::GROUPING_CLUSTERED,
    range(0, count($dataSeriesValues) - 1),
    $dataSeriesLabels,
    $xAxisTickValues,
    $dataSeriesValues
);

$series->setPlotDirection(DataSeries::DIRECTION_COL);

$plotArea = new PlotArea(null, [$series]);
$legend = new Legend(Legend::POSITION_RIGHT, null, false);
$title = new Title('Total');

$chart = new Chart(
    'flagChart',
    $title,
    $legend,
    $plotArea
);

$chart->setTopLeftPosition('D6');
$chart->setBottomRightPosition('K25');

$flagSheet->addChart($chart);

/* ================= CREATE SHEET 3: VARIASI KAPAL ================= */
// Count vessel types by flag
$variationData = [];
foreach ($data as $row) {
    // Skip baris kosong/header
    if (empty($row['mmsi']) && empty($row['name'])) {
        continue;
    }
    
    $flag = trim($row['flag'] ?? '');
    $type = $row['vessel_type'] ?? 'Other';
    
    // Skip jika flag kosong
    if ($flag === '') {
        continue;
    }
    
    if (!isset($variationData[$type])) {
        $variationData[$type] = [];
    }
    
    if (!isset($variationData[$type][$flag])) {
        $variationData[$type][$flag] = 0;
    }
    
    $variationData[$type][$flag]++;
}

// Get all unique flags from current data for columns
$allFlags = array_keys($flagCount);
sort($allFlags);

$varSheet = $spreadsheet->createSheet();
$varSheet->setTitle('variasi kapal');

// Headers
$varSheet->setCellValue('A1', 'Count of Flag');
$varSheet->setCellValue('A2', 'Row Labels');
$varSheet->setCellValue('B2', 'Column Labels');

// Write flag names as column headers
$col = 2; // Starting from column C (index 2)
foreach ($allFlags as $flag) {
    $varSheet->setCellValueByColumnAndRow($col, 2, $flag);
    $col++;
}
$varSheet->setCellValueByColumnAndRow($col, 2, 'Grand Total');

// Write vessel types and counts
$varRow = 3;
$typeOrder = ['Cargo', 'Other', 'Passenger', 'Tanker', 'Tug', 'Fishing'];

foreach ($typeOrder as $type) {
    if (!isset($variationData[$type])) continue;
    
    $varSheet->setCellValue("A{$varRow}", $type);
    
    $rowTotal = 0;
    $col = 2;
    foreach ($allFlags as $flag) {
        $count = $variationData[$type][$flag] ?? 0;
        if ($count > 0) {
            $varSheet->setCellValueByColumnAndRow($col, $varRow, $count);
            $rowTotal += $count;
        }
        $col++;
    }
    
    // Row total
    $varSheet->setCellValueByColumnAndRow($col, $varRow, $rowTotal);
    $varRow++;
}

// Grand Total row
$varSheet->setCellValue("A{$varRow}", 'Grand Total');
$varSheet->getStyle("A{$varRow}")->getFont()->setBold(true);

$col = 2;
foreach ($allFlags as $flag) {
    $varSheet->setCellValueByColumnAndRow($col, $varRow, $flagCount[$flag]);
    $col++;
}
$grandTotal = array_sum($flagCount);
$varSheet->setCellValueByColumnAndRow($col, $varRow, $grandTotal);

// Auto-size columns
for ($i = 0; $i <= $col; $i++) {
    $varSheet->getColumnDimensionByColumn($i)->setAutoSize(true);
}

/* ================= CREATE STACKED BAR CHART FOR VARIASI KAPAL ================= */
// Create data series for each flag
$varDataSeriesValues = [];
$varDataSeriesLabels = [];
$flagCol = 2;

foreach ($allFlags as $flag) {
    $colLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($flagCol);
    
    // Create label with direct string value
    $label = new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_STRING,
        null,
        null,
        1
    );
    $label->setDataValues([$flag]); // Set flag name directly
    
    $varDataSeriesLabels[] = $label;
    
    $varDataSeriesValues[] = new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_NUMBER,
        "'variasi kapal'!\${$colLetter}\$3:\${$colLetter}\$" . ($varRow - 1),
        null,
        count($typeOrder)
    );
    
    $flagCol++;
}

$varXAxisTickValues = [
    new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_STRING,
        "'variasi kapal'!\$A\$3:\$A\$" . ($varRow - 1),
        null,
        count($typeOrder)
    )
];

$varSeries = new DataSeries(
    DataSeries::TYPE_BARCHART,
    DataSeries::GROUPING_STACKED,
    range(0, count($varDataSeriesValues) - 1),
    $varDataSeriesLabels,
    $varXAxisTickValues,
    $varDataSeriesValues
);

$varSeries->setPlotDirection(DataSeries::DIRECTION_BAR);

$varPlotArea = new PlotArea(null, [$varSeries]);
$varLegend = new Legend(Legend::POSITION_RIGHT, null, false);
$varTitle = new Title('Variasi Kapal');

$varChart = new Chart(
    'varChart',
    $varTitle,
    $varLegend,
    $varPlotArea
);

$varChart->setTopLeftPosition('D5');
$varChart->setBottomRightPosition('T25');

$varSheet->addChart($varChart);

/* ================= CREATE SHEET 4: TYPE KAPAL (DENGAN FORMULA) ================= */
$typeCount = [];
foreach ($data as $row) {
    // Skip baris kosong/header
    if (empty($row['mmsi']) && empty($row['name'])) {
        continue;
    }
    
    $type = $row['vessel_type'] ?? 'Other';
    if (!isset($typeCount[$type])) {
        $typeCount[$type] = 0;
    }
    $typeCount[$type]++;
}

$typeSheet = $spreadsheet->createSheet();
$typeSheet->setTitle('TYPE KAPAL');

// Headers
$typeSheet->setCellValue('A4', 'Row Labels');
$typeSheet->setCellValue('B4', 'Count of type kapal');

$typeSheet->getStyle('A4:B4')->applyFromArray([
    'font' => ['bold' => true],
    'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'startColor' => ['rgb' => 'D9E1F2']
    ]
]);

// Write type counts dengan FORMULA
$typeRow = 5;
foreach ($typeOrder as $type) {
    $typeSheet->setCellValue("A{$typeRow}", $type);
    
    // GUNAKAN FORMULA COUNTIF - akan otomatis update!
    $typeSheet->setCellValue("B{$typeRow}", 
        "=COUNTIF('DATA KAPAL'!S:S,A{$typeRow})"
    );
    
    $typeRow++;
}

// Grand Total dengan SUM formula
$typeSheet->setCellValue("A{$typeRow}", 'Grand Total');
$typeSheet->setCellValue("B{$typeRow}", 
    "=SUM(B5:B" . ($typeRow - 1) . ")"
);

$typeSheet->getStyle("A{$typeRow}:B{$typeRow}")->applyFromArray([
    'font' => ['bold' => true],
    'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'startColor' => ['rgb' => 'D9E1F2']
    ]
]);

// Auto-size columns
$typeSheet->getColumnDimension('A')->setAutoSize(true);
$typeSheet->getColumnDimension('B')->setAutoSize(true);

/* ================= CREATE CHART FOR TYPE KAPAL ================= */
// Create multiple data series (one per vessel type) for different colors
$typeDataSeriesLabels = [];
$typeXAxisTickValues = [];
$typeDataSeriesValues = [];

$typeRowStart = 5;
$typeRowNum = 0;
foreach ($typeOrder as $type) {
    $currentTypeRow = $typeRowStart + $typeRowNum;
    
    $typeDataSeriesLabels[] = new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_STRING,
        "'TYPE KAPAL'!\$A\${$currentTypeRow}",
        null,
        1
    );
    
    $typeXAxisTickValues[] = new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_STRING,
        "'TYPE KAPAL'!\$A\${$currentTypeRow}",
        null,
        1
    );
    
    $typeDataSeriesValues[] = new DataSeriesValues(
        DataSeriesValues::DATASERIES_TYPE_NUMBER,
        "'TYPE KAPAL'!\$B\${$currentTypeRow}",
        null,
        1
    );
    
    $typeRowNum++;
}

$typeSeries = new DataSeries(
    DataSeries::TYPE_BARCHART,
    DataSeries::GROUPING_STANDARD,
    range(0, count($typeDataSeriesValues) - 1),
    $typeDataSeriesLabels,
    $typeXAxisTickValues,
    $typeDataSeriesValues
);

$typeSeries->setPlotDirection(DataSeries::DIRECTION_COL);

$typePlotArea = new PlotArea(null, [$typeSeries]);
$typeLegend = new Legend(Legend::POSITION_RIGHT, null, false);
$typeTitle = new Title('Type Kapal');

$typeChart = new Chart(
    'typeChart',
    $typeTitle,
    $typeLegend,
    $typePlotArea
);

$typeChart->setTopLeftPosition('D3');
$typeChart->setBottomRightPosition('K18');

$typeSheet->addChart($typeChart);

/* ================= OUTPUT ================= */
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->setIncludeCharts(true);
$writer->save('php://output');
exit;