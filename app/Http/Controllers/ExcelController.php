<?php

namespace App\Http\Controllers;

use App\Exports\Excel;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use ZipArchive;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Illuminate\Filesystem\Filesystem;

class ExcelController extends Controller
{
    public function index()
    {
        return view('index');
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
        $excel = '';
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
}
