<?php
declare(strict_types=1);

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

require __DIR__ . '/vendor/autoload.php';

/* ================= HEADERS ================= */
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="laporan_harian_ais.xlsx"');
header('Cache-Control: max-age=0');

/* ================= LOAD TEMPLATE ================= */
$templatePath = realpath(__DIR__ . '/templates/laporan_template.xlsx');
if ($templatePath === false) {
    http_response_code(500);
    exit('Template Excel tidak ditemukan');
}

$spreadsheet = IOFactory::load($templatePath);
$sheet = $spreadsheet->getActiveSheet();

/* ================= READ INPUT JSON ================= */
$data = json_decode(file_get_contents('php://input'), true);
if (!is_array($data)) {
    http_response_code(400);
    exit('Data JSON tidak valid');
}

/* ================= HELPER ================= */
function fillBgWithBorder($sheet, string $cell, string $color): void
{
    $sheet->getStyle($cell)->applyFromArray([
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => ['rgb' => $color],
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['rgb' => '000000'],
            ],
        ],
    ]);
}

function alignCenter($sheet, string $cell): void
{
    $sheet->getStyle($cell)->getAlignment()
        ->setHorizontal(Alignment::HORIZONTAL_CENTER)
        ->setVertical(Alignment::VERTICAL_CENTER);
}

/* ================= VESSEL CONFIG ================= */
$vesselOrder = ['Cargo', 'Fishing', 'Other', 'Passenger', 'Tanker', 'Tug'];

$vesselColors = [
    'Cargo'     => '90EE90',
    'Fishing'   => 'FFCC00',
    'Other'     => 'FFFF00',
    'Passenger' => '000BF6',
    'Tanker'    => 'FF0000',
    'Tug'       => 'FF902C'
];

/* ================= GROUP DATA ================= */
$groupedData = array_fill_keys($vesselOrder, []);
foreach ($data as $row) {
    $type = $row['vessel_type'] ?? 'Other';
    $groupedData[$type][] = $row;
}

/* ================= START ROW ================= */
$currentRow = 5;
$dateExport = date('Y-m-d');

/* ================= WRITE DATA ================= */
foreach ($vesselOrder as $vesselType) {

    if (empty($groupedData[$vesselType])) {
        continue;
    }

    /* ===== CATEGORY HEADER ===== */
    if ($vesselType !== 'Cargo') {
        $sheet->setCellValue("A{$currentRow}", strtoupper($vesselType) . ' SHIP');
        $sheet->mergeCells("A{$currentRow}:T{$currentRow}");
        $sheet->getStyle("A{$currentRow}:T{$currentRow}")->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => $vesselColors[$vesselType]],
            ],
            'font' => [
                'bold' => true,
                'size' => 16,
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical'   => Alignment::VERTICAL_CENTER,
            ],
        ]);
        $currentRow++;
    }


/* ===== DATA ROW ===== */
$no = 1;
foreach ($groupedData[$vesselType] as $row) {
    
    $sheet->setCellValue("A{$currentRow}", $no++);
    $sheet->setCellValue("B{$currentRow}", $row['date'] ?? $dateExport);
    
    // Time - SIMPLIFIED VERSION
    $timeValue = $row['time'] ?? '15:00'; // default ke 15:00 jika kosong
    $sheet->setCellValueExplicit("C{$currentRow}", $timeValue, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
    
    // Calculate AM/PM
    if (!empty($timeValue)) {
        list($hour) = explode(':', $timeValue);
        $period = ((int)$hour >= 12) ? 'PM' : 'AM';
    } else {
        $period = '';
    }
    $sheet->setCellValueExplicit("D{$currentRow}", $period, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
    
    $sheet->setCellValue("E{$currentRow}", $row['name'] ?? '');
    $sheet->setCellValue("F{$currentRow}", $row['callsign'] ?? '');
    $sheet->setCellValue("G{$currentRow}", $row['mmsi'] ?? '');
    $sheet->setCellValue("H{$currentRow}", $row['course'] ?? '');
    $sheet->setCellValue("I{$currentRow}", $row['speed'] ?? '');
    $sheet->setCellValue("J{$currentRow}", $row['imo'] ?? '');
    $sheet->setCellValue("K{$currentRow}", $row['flag'] ?? '');
    $sheet->setCellValue("L{$currentRow}", $row['destination'] ?? '');
    
    $sheet->setCellValue("M{$currentRow}", $row['latitude1'] ?? '');
    
 $sheet->getStyle("M{$currentRow}")->getNumberFormat()->setFormatCode('00"°"');

    $sheet->setCellValue("N{$currentRow}", $row['latitude2'] ?? '');
    $sheet->setCellValue("O{$currentRow}", $row['latitude3'] ?? '');
    $sheet->setCellValue("P{$currentRow}", $row['latitude4'] ?? '');
    
    $sheet->setCellValue("Q{$currentRow}", $row['longitude1'] ?? '');
    $sheet->setCellValue("R{$currentRow}", $row['longitude2'] ?? '');
    $sheet->setCellValue("S{$currentRow}", $row['longitude3'] ?? '');
    
    $sheet->setCellValue("T{$currentRow}", $row['status'] ?? '');
    
    /* ===== COLOR + BORDER ===== */
    fillBgWithBorder($sheet, "A{$currentRow}", "FFFFFF");
    fillBgWithBorder($sheet, "B{$currentRow}", "EEECE0");
    
    foreach (['C','D','G','H','J','K'] as $col) {
        fillBgWithBorder($sheet, "{$col}{$currentRow}", "C0DAF3");
    }
    
    foreach (['E','F','I','L','T'] as $col) {
        fillBgWithBorder($sheet, "{$col}{$currentRow}", "FFE8D7");
    }
    
    foreach (['M','N','O','P'] as $col) {
        fillBgWithBorder($sheet, "{$col}{$currentRow}", "E5DFED");
    }
    
    foreach (['Q','R','S'] as $col) {
        fillBgWithBorder($sheet, "{$col}{$currentRow}", "F7DBDA");
    }
    
    /* ===== ALIGN CENTER ALL DATA ===== */
    foreach (range('A', 'T') as $col) {
        alignCenter($sheet, "{$col}{$currentRow}");
    }
    
    $currentRow++;
}

    $currentRow++;
}
/* ================= FOOTER / WEATHER INFO ================= */
// Tambah baris kosong
$currentRow += 2;

// Weather information
// TINGGI GELOMBANG (A-D merged) dan value di E
$sheet->setCellValue("A{$currentRow}", "TINGGI GELOMBANG");
$sheet->mergeCells("A{$currentRow}:D{$currentRow}");
$sheet->setCellValue("E{$currentRow}", ": SEDANG, 1.9 M");
$currentRow++;

// ARAH ANGIN & KECEPATAN ANGIN (A-D merged) dan value di E
$sheet->setCellValue("A{$currentRow}", "ARAH ANGIN & KECEPATAN ANGIN");
$sheet->mergeCells("A{$currentRow}:D{$currentRow}");
$sheet->setCellValue("E{$currentRow}", ": BARAT - TIMUR, 8 - 16 KNOT");
$currentRow++;

// SUHU UDARA (A-D merged) dan value di E
$sheet->setCellValue("A{$currentRow}", "SUHU UDARA");
$sheet->mergeCells("A{$currentRow}:D{$currentRow}");
$sheet->setCellValue("E{$currentRow}", ": 29-31° C");
$currentRow++;

// CUACA (A-D merged) dan value di E
$sheet->setCellValue("A{$currentRow}", "CUACA");
$sheet->mergeCells("A{$currentRow}:D{$currentRow}");
$sheet->setCellValue("E{$currentRow}", ": BERAWAN TEBAL");
$currentRow++;

/* ================= TANDA TANGAN ================= */
// Tanda tangan dimulai dari baris yang sama dengan CUACA (atau baris setelahnya)
// Posisi di kolom M-S (kanan bawah), merged dan rata tengah
$signRow = $currentRow - 3; // Kembali ke baris pertama weather info untuk align dengan CUACA

// Tanggal dan tempat (kolom M-S merged, rata tengah)
$sheet->setCellValue("M{$signRow}", "Karangasem, 31 Desember 2025");
$sheet->mergeCells("M{$signRow}:S{$signRow}");
$sheet->getStyle("M{$signRow}:S{$signRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$signRow += 2;

// Nama jabatan (kolom M-S merged, rata tengah)
$sheet->setCellValue("M{$signRow}", "Piket AIS Stasiun Bakamla Bali");
$sheet->mergeCells("M{$signRow}:S{$signRow}");
$sheet->getStyle("M{$signRow}:S{$signRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$signRow += 2;

// TTD (kolom M-S merged, rata tengah)
$sheet->setCellValue("M{$signRow}", "TTD");
$sheet->mergeCells("M{$signRow}:S{$signRow}");
$sheet->getStyle("M{$signRow}:S{$signRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$signRow += 2;

// Nama (kolom M-S merged, rata tengah)
$sheet->setCellValue("M{$signRow}", "I NYOMAN GELGEL");
$sheet->mergeCells("M{$signRow}:S{$signRow}");
$sheet->getStyle("M{$signRow}:S{$signRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);



/* ================= OUTPUT ================= */
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
exit;
