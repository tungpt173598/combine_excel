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
        $destination = storage_path('excels');
        if (!\File::exists($destination)) {
            \File::makeDirectory($destination, 0755, true);
        }
        $file = new Filesystem;
        $file->cleanDirectory(storage_path('excels'));
        $zip->extractTo($destination);
        $zip->close();
        if (str_contains($fileExp, '/')) {
            $folderName = explode('/', $fileExp)[0];
            $excels = Storage::disk('local')->allFiles($folderName);
        } else {
            $folderName = explode('.', $file->getClientOriginalName())[0];
            $excels = Storage::disk('local')->allFiles();
        }
        $results = [
            ['STT', 'UID', 'phone', 'email', 'name', 'gender', 'country', 'group_name']
        ];
        $error = [];
        foreach ($excels as $key => $excel) {
            try {
                $inputFileName = storage_path('excels/' . $excel);
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
                $spreadsheet  = $spreadsheet->getActiveSheet();
                $data = $spreadsheet->toArray();
                unset($data[0]);
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
        $folder = new Filesystem;
        $folder->cleanDirectory(storage_path('add'));
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
        if (str_contains($fileExp, '/')) {
            $folderName = explode('/', $fileExp)[0];
            $excels = Storage::disk('add')->allFiles($folderName);
        } else {
            $folderName = explode('.', $file->getClientOriginalName())[0];
            $excels = Storage::disk('add')->allFiles();
        }
        foreach ($excels as $excel) {
            $inputFileName = storage_path('add/' . $excel);
            $spreadsheet = IOFactory::load($inputFileName);
            $spreadsheet1  = $spreadsheet->getActiveSheet();
            $rows = $spreadsheet1->getHighestRow();
            $fileName1 = explode('/', $excel);
            $fileName2 = end($fileName1);
            $fileName3 = explode('.', $fileName2)[0];
            for ($i = 2; $i <= $rows; $i++) {
                if (!empty($spreadsheet1->getCell('A' . $i)->getValue())) {
                    $spreadsheet1->setCellValue('H' . $i, $fileName3);
                    $name = $spreadsheet1->getCell('E' . $i)->getValue();
                    if (empty($name)) {
                        $spreadsheet1->setCellValue('E' . $i, 'KH');
                    }
                }
            }
            $fileName = explode('.', $excel);
            $writer = IOFactory::createWriter($spreadsheet, ucfirst(strtolower(end($fileName))));
            $writer->save($inputFileName);
        }
        $zip2 = new ZipArchive();
        $zipName = $folderName . '.zip';
        if ($zip2->open(public_path('storage/' . $zipName), ZipArchive::CREATE) === TRUE)
        {
            if (str_contains($fileExp, '/')) {
                $files = File::files(storage_path('add/' . $folderName));
            } else {
                $files = File::files(storage_path('add'));
            }
            foreach ($files as $value) {
                $zip2->addFile($value, basename($value));
            }
            $zip2->close();
        }
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
                $spreadsheet  = $spreadsheet->getActiveSheet()->toArray();
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

    public function getSplit()
    {
        return view('split');
    }

    public function splitFile(Request $request)
    {
        $file = $request->file('folder');
        $fileName = explode('.', $file->getClientOriginalName())[0];
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file->getRealPath());
        $spreadsheet  = $spreadsheet->getActiveSheet()->toArray();
        $lines = array_chunk($spreadsheet, 1000);
        $count = count($lines);
        if (count(end($lines)) < 500) {
            $lines[$count -2] = array_merge($lines[$count -2], $lines[$count -1]);
            unset($lines[$count - 1]);
        }
        foreach ($lines as $key => $value) {
            \Excel::store(new Excel($value), 'split/' . $key + 1 . '.xlsx');
        }
        $zip = new ZipArchive();
        $zipName = $fileName . '.zip';
        if ($zip->open(public_path('storage/' . $zipName), ZipArchive::CREATE) === TRUE)
        {
            $files = File::files(storage_path('excels/split'));
            foreach ($files as $value) {
                $zip->addFile($value, basename($value));
            }
            $zip->close();
        }
        $file = new Filesystem;
        $file->cleanDirectory(storage_path('excels'));
        return response()->download(public_path('storage/' . $zipName), $zipName)->deleteFileAfterSend();

    }
}
