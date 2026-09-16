<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Attendance Sheet Generator</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
    <div class="container py-5">
        <div class="row justify-content-center">
            <div class="col-lg-7">
                <div class="card shadow-sm">
                    <div class="card-body p-4">
                        <h4 class="mb-1">Attendance Sheet Generator</h4>
                        <p class="text-muted mb-4">
                            Upload each employee's daily attendance file (Payroll Num, Name, and Date / Time In / Time Out rows)
                            together with the employee list, and download a
                            <strong>zip file with one Excel sheet per day</strong> (Attendance Day 01.xlsx,
                            Attendance Day 02.xlsx ...) — only for the days that actually have data.
                        </p>

                        @if ($errors->any())
                            <div class="alert alert-danger">
                                <ul class="mb-0">
                                    @foreach ($errors->all() as $error)
                                        <li>{{ $error }}</li>
                                    @endforeach
                                </ul>
                            </div>
                        @endif

                        <form action="{{ route('attendance.generate') }}" method="POST" enctype="multipart/form-data">
                            @csrf

                            <div class="mb-3">
                                <label class="form-label">Month (used only to label the zip file)</label>
                                <input type="month" name="month" class="form-control"
                                       value="{{ old('month', now()->format('Y-m')) }}" required>
                            </div>

                            <div class="mb-3">
                                <label class="form-label">Employee list (optional)</label>
                                <input type="file" name="employee_list" class="form-control"
                                       accept=".xlsx,.xls,.ods">
                                <div class="form-text">
                                    A file with "Employee Code" in column A and "Employee Name" in column B.
                                    If skipped, names are taken from each daily file instead.
                                </div>
                            </div>

                            <div class="mb-4">
                                <label class="form-label">Daily attendance files (one per employee)</label>
                                <input type="file" name="daily_files[]" class="form-control"
                                       accept=".xlsx,.xls,.ods" multiple required>
                                <div class="form-text">
                                    Select all employee files together (you can multi-select in the file dialog).
                                </div>
                            </div>

                            <button type="submit" class="btn btn-primary w-100">
                                Generate &amp; Download Attendance ZIP (one file per day)
                            </button>
                        </form>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
