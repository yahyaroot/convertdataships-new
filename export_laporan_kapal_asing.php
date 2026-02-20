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

require __DIR__ . '/vendor/autoload.php';

/* ================= HEADERS ================= */
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="laporan_harian_ais_kapal_asing.xlsx"');
header('Cache-Control: max-age=0');

/* ================= LOAD TEMPLATE ================= */
$templatePath = realpath(__DIR__ . '/templates/laporan_template2.xlsx');

if ($templatePath === false) {
    http_response_code(500);
    exit('Template Excel tidak ditemukan');
}

$spreadsheet = IOFactory::load($templatePath);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('DATA KAPAL ASING');

/* ================= READ INPUT JSON ================= */
$raw = file_get_contents('php://input');
$data = json_decode($raw, true);

if (!is_array($data)) {
    http_response_code(400);
    exit('Data JSON tidak valid');
}

/* ================= FILTER KAPAL IKAN ASING ================= */
$foreignVessels = [];
/* ================= FILTER KAPAL ASING (SEMUA JENIS) ================= */
$foreignVessels = [];
foreach ($data as $row) {
    // Skip baris kosong
    if (empty($row['mmsi']) && empty($row['name'])) {
        continue;
    }
    
    $flag = strtolower(trim($row['flag'] ?? ''));
    
    // ✅ Ambil SEMUA kapal asing (tidak peduli jenis kapal)
    if ($flag !== '' && 
        $flag !== 'indonesia' && 
        $flag !== 'idn' && 
        $flag !== 'id') {
        $foreignVessels[] = $row;
    }
}

/* ================= DATE EXPORT ================= */
$dateExport = date('d-m-Y');

/* ================= START ROW ================= */
$startRow = 19;
$currentRow = $startRow;

/* ================= COUNT FLAGS (KAPAL IKAN ASING SAJA) ================= */
$flagCount = [];
foreach ($foreignVessels as $row) {
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

/* ================= WRITE DATA TO SHEET 1 (KAPAL IKAN ASING) ================= */
$no = 1;

foreach ($foreignVessels as $row) {
    $sheet->setCellValue("A{$currentRow}", $no++);

    // Time & CTU
    $sheet->setCellValue("B{$currentRow}", $row['time'] ?? '15:00');
    $sheet->setCellValue("C{$currentRow}", $row['ctu'] ?? 'PM');

    // Ship Info
    $sheet->setCellValue("D{$currentRow}", $row['name'] ?? '');
    $sheet->setCellValue("E{$currentRow}", $row['callsign'] ?? '');
    $sheet->setCellValue("F{$currentRow}", $row['mmsi'] ?? '');

    /* ================= COURSE (°) ================= */
    $sheet->setCellValue("G{$currentRow}", (float)($row['course'] ?? 0));
    $sheet->getStyle("G{$currentRow}")
          ->getNumberFormat()
          ->setFormatCode('000"°"');

    /* ================= SPEED (kts) ================= */
    $sheet->setCellValue("H{$currentRow}", (float)($row['speed'] ?? 0));
    $sheet->getStyle("H{$currentRow}")
          ->getNumberFormat()
          ->setFormatCode('0" kts"');

    // Other info
    $sheet->setCellValue("I{$currentRow}", $row['imo'] ?? '');
    $sheet->setCellValue("J{$currentRow}", $row['flag'] ?? '');
    $sheet->setCellValue("K{$currentRow}", $row['vessel_type'] ?? '');
    $sheet->setCellValue("L{$currentRow}", $row['destination'] ?? '');

    /* ================= LATITUDE ================= */
    $sheet->setCellValue("M{$currentRow}", (float)($row['latitude1'] ?? 0));
    $sheet->setCellValue("N{$currentRow}", $row['latitude2'] ?? '');
    $sheet->setCellValue("O{$currentRow}", $row['latitude3'] ?? '');
    $sheet->setCellValue("P{$currentRow}", $row['latitude4'] ?? '');

    // >>> FORMAT DEGREE UNTUK CELL M <
    $sheet->getStyle("M{$currentRow}")
          ->getNumberFormat()
          ->setFormatCode('00"°";;00"°";');

    /* ================= LONGITUDE ================= */
    $sheet->setCellValue("Q{$currentRow}", $row['longitude1'] ?? '');
    $sheet->setCellValue("R{$currentRow}", $row['longitude2'] ?? '');
    $sheet->setCellValue("S{$currentRow}", $row['longitude3'] ?? '');

    // Status
    $sheet->setCellValue("T{$currentRow}", $row['status'] ?? '');

    /* ================= STYLE ================= */
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
                'color' => ['rgb' => '000000'],
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ]);
    };

    // Base colors
    $applyStyle("A{$currentRow}", 'FFFFFF');
    $applyStyle("B{$currentRow}", 'C0DAF3');
    $applyStyle("C{$currentRow}", 'C0DAF3');
    $applyStyle("D{$currentRow}", 'FFE8D7');
    $applyStyle("E{$currentRow}", 'FFE8D7');
    $applyStyle("F{$currentRow}", 'C0DAF3');
    $applyStyle("G{$currentRow}", 'C0DAF3'); // Course
    $applyStyle("H{$currentRow}", 'FFE8D7'); // Speed
    $applyStyle("I{$currentRow}", 'C0DAF3');
    $applyStyle("J{$currentRow}", 'C0DAF3');
    $applyStyle("K{$currentRow}", 'FFE8D7');
    $applyStyle("L{$currentRow}", 'FFE8D7');

    // Latitude (Purple)
    foreach (['M','N','O','P'] as $c) {
        $applyStyle("{$c}{$currentRow}", 'E5DFED');
    }

    // Longitude (Pink)
    foreach (['Q','R','S'] as $c) {
        $applyStyle("{$c}{$currentRow}", 'F7DBDA');
    }

    // Status
    $applyStyle("T{$currentRow}", 'C0DAF3');

    $currentRow++;
}

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
$writer->setIncludeCharts(true);
$writer->save('php://output');
exit;