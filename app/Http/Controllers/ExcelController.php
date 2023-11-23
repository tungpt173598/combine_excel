<?php

namespace App\Http\Controllers;

use App\Exports\Excel;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Storage;
use ZipArchive;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Illuminate\Filesystem\Filesystem;
use PhpOffice\PhpSpreadsheet\IOFactory;

class ExcelController extends Controller
{
    public function index()
    {
        return view('index');
    }

    public function add()
    {
        return view('add');
    }

    public function mergeFile(Request $request)
    {
        $zip = new ZipArchive();
        $file = $request->file('folder');
        $zip->open($file->getRealPath());
        $fileExp = $zip->statIndex(0)['name'];
        $folderName = explode('/', $fileExp)[0];
        $destination = storage_path('excels');
        if (!\File::exists($destination)) {
            \File::makeDirectory($destination, 0755, true);
        }
        $zip->extractTo($destination);
        $zip->close();
        $excels = Storage::disk('local')->allFiles($folderName);
        $results = [];
        $error = [];
        foreach ($excels as $key => $excel) {
            try {
                $inputFileName = storage_path('excels/' . $excel);
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
                $spreadsheet  = $spreadsheet->getActiveSheet();
                $data = $spreadsheet->toArray();
                if ($key != 0) {
                    unset($data[0]);
                }
                $results = array_merge($results, $data);
                unset($spreadsheet, $data);
            } catch (\Exception $exception) {
                $message = explode('->', $exception->getMessage())[0];
                $error[] = [
                    "file-{$key}" => $excel,
                    "line-{$key}" => explode('!', $message)[1] ?? ''
                ];
            }

        }
        $file = new Filesystem;
        $file->cleanDirectory(storage_path('excels'));
        if (empty($error)) {
            $fileName = $folderName . '.xlsx';
            return \Excel::download(new Excel($results), $fileName);
        }
        return redirect()->back()->withErrors($error);
    }

    public function addFile(Request $request)
    {
        $zip = new ZipArchive();
        $file = $request->file('folder');
        $zip->open($file->getRealPath());
        $fileExp = $zip->statIndex(0)['name'];
        $folderName = explode('/', $fileExp)[0];
        $destination = storage_path('add');
        if (!\File::exists($destination)) {
            \File::makeDirectory($destination, 0755, true);
        }
        $zip->extractTo($destination);
        $zip->close();
        $excels = Storage::disk('add')->allFiles($folderName);
        foreach ($excels as $key => $excel) {
            $inputFileName = storage_path('add/' . $excel);
            $spreadsheet = IOFactory::load($inputFileName);
            $spreadsheet1  = $spreadsheet->getActiveSheet();
            $rows = $spreadsheet1->getHighestRow();
            $fileName1 = explode('/', $excel);
            for ($i = 2; $i <= $rows; $i++) {
                $spreadsheet1->setCellValue('H' . $i, end($fileName1));
            }
            $fileName = explode('.', $excel);
            $writer = IOFactory::createWriter($spreadsheet, ucfirst(end($fileName)));
            $writer->save($inputFileName);
        }
        $zip2 = new ZipArchive();
        $zipName = $folderName . '.zip';
        if ($zip2->open(public_path('storage/' . $zipName), ZipArchive::CREATE) === TRUE)
        {
            $files = File::files(storage_path('add/' . $folderName));
            foreach ($files as $value) {
                $zip2->addFile($value, basename($value));
            }
            $zip2->close();
        }
        $folder = new Filesystem;
        $folder->cleanDirectory(storage_path('add'));
        return response()->download(public_path('storage/' . $zipName), $zipName)->deleteFileAfterSend();
    }

    public function getCheck()
    {
        return view('check');
    }

    public function check(Request $request)
    {
        $zip = new ZipArchive();
        $file = $request->file('folder');
        $zip->open($file->getRealPath());
        $fileExp = $zip->statIndex(0)['name'];
        $folderName = explode('/', $fileExp)[0];
        $destination = storage_path('excels');
        if (!\File::exists($destination)) {
            \File::makeDirectory($destination, 0755, true);
        }
        $zip->extractTo($destination);
        $zip->close();
        $excels = Storage::disk('local')->allFiles($folderName);
        $error = [];
        foreach ($excels as $key => $excel) {
            try {
                $inputFileName = storage_path('excels/' . $excel);
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
            } catch (\Exception $exception) {
                $message = explode('->', $exception->getMessage())[0];
                $error[] = [
                    "file-{$key}" => $excel,
                    "line-{$key}" => explode('!', $message)[1] ?? ''
                ];
            }
        }
        $file = new Filesystem;
        $file->cleanDirectory(storage_path('excels'));
        if (empty($error)) {
            return redirect()->back()->with(['success' => 'Không có file lỗi']);
        }
        return redirect()->back()->withErrors($error);
    }
}
