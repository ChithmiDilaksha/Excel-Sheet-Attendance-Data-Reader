<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use ZipArchive;

class AttendanceController extends Controller
{
    /**
     * Show the upload form.
     */
    public function index()
    {
        return view('attendance.index');
    }

    /**
     * Handle the uploaded files, build one workbook with one sheet per day,
     * and return it as a download.
     */
    public function generate(Request $request)
    {
        $request->validate([
            'month'          => 'required|date_format:Y-m',
            'employee_list'  => 'nullable|file|mimes:xlsx,xls,ods',
            'daily_files'    => 'required|array|min:1',
            'daily_files.*'  => 'file|mimes:xlsx,xls,ods',
        ], [
            'daily_files.required' => 'Please upload at least one employee attendance file.',
        ]);

        // Only used to label the downloaded zip file, e.g. Attendance_2026-09.zip
        $monthLabel = $request->input('month');

        // 1) Optional master list -> [employee_code => employee_name]
        $employeeNames = [];
        if ($request->hasFile('employee_list')) {
            $employeeNames = $this->readEmployeeList($request->file('employee_list')->getRealPath());
        }

        // 2) Read every per-employee daily file and bucket rows by day number.
        //    Days come purely from what's actually in the uploaded files -
        //    if a file only has 28 rows, only days 1-28 get generated.
        $byDay = [];

        foreach ($request->file('daily_files') as $file) {
            $parsed = $this->readDailyFile($file->getRealPath());

            if (!$parsed) {
                continue; // skip unreadable file, don't fail the whole batch
            }

            $code = $parsed['code'];
            $name = $employeeNames[$code] ?? $parsed['name'];

            foreach ($parsed['days'] as $day => $times) {
                $byDay[$day][] = [
                    'code' => $code,
                    'name' => $name,
                    'in'   => $times['in'],
                    'out'  => $times['out'],
                ];
            }
        }

        if (empty($byDay)) {
            return back()->withErrors([
                'daily_files' => 'No attendance rows were found in the uploaded files.',
            ]);
        }

        ksort($byDay); // day 1, 2, 3 ... in order

        // Row order for every day's file: master list order first, then any
        // employee not in the master list, in the order their file appeared.
        $employeeOrder = [];
        foreach ($employeeNames as $code => $name) {
            $employeeOrder[$code] = count($employeeOrder);
        }
        foreach ($byDay as $rows) {
            foreach ($rows as $entry) {
                if (!isset($employeeOrder[$entry['code']])) {
                    $employeeOrder[$entry['code']] = count($employeeOrder);
                }
            }
        }

        // 3) Build one .xlsx per day, zipped together.
        $zipPath = $this->buildZip($byDay, $employeeOrder, $monthLabel);

        return response()
            ->download($zipPath, 'Attendance_' . $monthLabel . '.zip')
            ->deleteFileAfterSend(true);
    }

    /**
     * Read the master "Employee Code / Employee Name" list.
     * Expected columns (row 1 = header): A = Employee Code, B = Employee Name.
     */
    private function readEmployeeList(string $path): array
    {
        $map = [];

        try {
            $spreadsheet = IOFactory::load($path);
            $sheet       = $spreadsheet->getSheet(0);
            $highestRow  = $sheet->getHighestRow();

            for ($row = 2; $row <= $highestRow; $row++) {
                $code = trim((string) $sheet->getCell("A{$row}")->getValue());
                $name = trim((string) $sheet->getCell("B{$row}")->getValue());

                if ($code === '') {
                    continue;
                }

                $map[$code] = $name;
            }
        } catch (\Throwable $e) {
            // If the master list can't be read, we simply fall back to the
            // name embedded in each daily attendance file.
        }

        return $map;
    }

    /**
     * Read one per-employee attendance file, shaped like:
     *
     *   Payroll Num | 331209
     *   Name        | W.M.M.B.Jayathilaka
     *   (blank row)
     *   Date | Time In | Time Out
     *   1    | 08:00:00 | 20:05:00
     *   2    | 07:56:00 | 20:10:00
     *   ...
     *
     * A day with "0" in both time columns is treated as no clock-in (absent),
     * and is written back out as plain 0, matching the target workbook.
     */
    private function readDailyFile(string $path): ?array
    {
        try {
            $spreadsheet = IOFactory::load($path);
        } catch (\Throwable $e) {
            return null;
        }

        $sheet      = $spreadsheet->getSheet(0);
        $highestRow = $sheet->getHighestRow();

        $code = trim((string) $sheet->getCell('B1')->getValue());
        $name = trim((string) $sheet->getCell('B2')->getValue());

        if ($code === '') {
            return null;
        }

        // Find the "Date / Time In / Time Out" header row so this keeps
        // working even if a row or two is added/removed above it.
        $headerRow = null;
        for ($row = 1; $row <= $highestRow; $row++) {
            $a = strtolower(trim((string) $sheet->getCell("A{$row}")->getValue()));
            if ($a === 'date') {
                $headerRow = $row;
                break;
            }
        }
        $headerRow = $headerRow ?? 4; // fallback to the template's default position

        $days = [];

        for ($row = $headerRow + 1; $row <= $highestRow; $row++) {
            $dayCell = $sheet->getCell("A{$row}")->getValue();

            if ($dayCell === null || $dayCell === '') {
                continue;
            }

            $day = (int) $dayCell;
            if ($day < 1) {
                continue;
            }

            // getCalculatedValue() returns a numeric Excel time serial for a
            // properly time-typed cell. Some source files store the time as
            // plain text instead (e.g. "7:55 AM"), so fall back to parsing
            // that text rather than treating it as an absent day.
            $days[$day] = [
                'in'  => $this->readTimeValue($sheet, "B{$row}"),
                'out' => $this->readTimeValue($sheet, "C{$row}"),
            ];
        }

        return ['code' => $code, 'name' => $name, 'days' => $days];
    }

    /**
     * Read a Time In / Time Out cell as a fraction-of-a-day float (0.5 = noon),
     * the same representation Excel/ODS use internally for time values.
     * Handles both a real time-typed cell (returns a numeric serial already)
     * and a cell where the time was entered as plain text (e.g. "7:55 AM",
     * "07:55:00") by parsing that text instead of assuming it's absent.
     */
    private function readTimeValue($sheet, string $coordinate): float
    {
        $raw = $sheet->getCell($coordinate)->getCalculatedValue();

        if (is_numeric($raw)) {
            return (float) $raw;
        }

        $text = trim((string) $raw);
        if ($text === '' || $text === '0') {
            return 0.0;
        }

        $timestamp = strtotime($text);
        if ($timestamp === false) {
            return 0.0;
        }

        $secondsIntoDay = ((int) date('H', $timestamp)) * 3600
            + ((int) date('i', $timestamp)) * 60
            + ((int) date('s', $timestamp));

        return $secondsIntoDay / 86400;
    }

    /**
     * Write a Time In / Time Out value the same way the target workbook does:
     * an absent day (0) is left as a plain number 0, a real clock-in is
     * written as an actual time value formatted as "h:mm AM/PM".
     */
    private function writeTimeCell($sheet, string $coordinate, float $value): void
    {
        if ($value == 0) {
            $sheet->setCellValue($coordinate, 0);
            return;
        }

        $sheet->setCellValue($coordinate, $value);
        $sheet->getStyle($coordinate)->getNumberFormat()->setFormatCode('h:mm AM/PM');
    }

    /**
     * Build one .xlsx file per day ("Attendance Day 01.xlsx", "Attendance Day
     * 02.xlsx", ...), each with a single sheet of that day's records, and
     * bundle them all into one zip file for download.
     */
    private function buildZip(array $byDay, array $employeeOrder, string $monthLabel): string
    {
        $tmpDir = storage_path('app/tmp/attendance_' . uniqid());
        mkdir($tmpDir, 0775, true);

        $filePaths = [];

        foreach ($byDay as $day => $rows) {
            $sheetTitle = 'Attendance Day ' . str_pad((string) $day, 2, '0', STR_PAD_LEFT);

            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $sheet->setTitle($sheetTitle);

            $sheet->setCellValue('A1', 'Employee Code');
            $sheet->setCellValue('B1', 'Employee Name');
            $sheet->setCellValue('C1', 'In Time');
            $sheet->setCellValue('D1', 'Out Time');

            usort($rows, fn ($a, $b) => $employeeOrder[$a['code']] <=> $employeeOrder[$b['code']]);

            $r = 2;
            foreach ($rows as $entry) {
                $sheet->setCellValueExplicit('A' . $r, $entry['code'], DataType::TYPE_STRING);
                $sheet->setCellValue('B' . $r, $entry['name']);

                $this->writeTimeCell($sheet, 'C' . $r, $entry['in']);
                $this->writeTimeCell($sheet, 'D' . $r, $entry['out']);

                $r++;
            }

            $sheet->getColumnDimension('A')->setWidth(18.8);
            $sheet->getColumnDimension('B')->setWidth(36.6);
            $sheet->getColumnDimension('C')->setWidth(16.5);
            $sheet->getColumnDimension('D')->setWidth(12.8);

            $filePath = $tmpDir . '/' . $sheetTitle . '.xlsx';
            (new Xlsx($spreadsheet))->save($filePath);
            $filePaths[] = $filePath;

            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
        }

        $zipPath = storage_path('app/tmp/Attendance_' . $monthLabel . '_' . uniqid() . '.zip');

        $zip = new ZipArchive();
        $zip->open($zipPath, ZipArchive::CREATE | ZipArchive::OVERWRITE);
        foreach ($filePaths as $path) {
            $zip->addFile($path, basename($path));
        }
        $zip->close();

        // Clean up the individual day files now that they're inside the zip.
        foreach ($filePaths as $path) {
            @unlink($path);
        }
        @rmdir($tmpDir);

        return $zipPath;
    }
}
