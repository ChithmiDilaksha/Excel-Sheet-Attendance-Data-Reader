<?php

use App\Http\Controllers\AttendanceController;
use Illuminate\Support\Facades\Route;

Route::get('/', [AttendanceController::class, 'index'])->name('attendance.index');
Route::post('/attendance/generate', [AttendanceController::class, 'generate'])->name('attendance.generate');
