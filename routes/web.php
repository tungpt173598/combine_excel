<?php

use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

Route::get('/', \App\Http\Controllers\ExcelController::class . '@index');
Route::get('/add', \App\Http\Controllers\ExcelController::class . '@add')->name('add');
Route::get('/check', \App\Http\Controllers\ExcelController::class . '@getCheck');
Route::post('/file', \App\Http\Controllers\ExcelController::class . '@mergeFile')->name('merge');
Route::post('/add-file', \App\Http\Controllers\ExcelController::class . '@addFile')->name('add-file');
Route::post('/check', \App\Http\Controllers\ExcelController::class . '@check')->name('check');
