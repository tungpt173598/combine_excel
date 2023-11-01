<?php

namespace App\Console\Commands;

use App\Exports\Excel;
use Illuminate\Console\Command;
use Illuminate\Filesystem\Filesystem;
use Illuminate\Support\Facades\Storage;

class MergeExcel extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'app:merge-excel {folder_name}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Command description';

    /**
     * Execute the console command.
     */
    public function handle()
    {
        $folderName = $this->argument('folder_name');
        $excels = Storage::disk('local')->allFiles($folderName);
        $results = [];
        foreach ($excels as $key => $excel) {
            $inputFileName = storage_path('excels/' . $excel);
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
            $spreadsheet  = $spreadsheet->getActiveSheet();
            $data = $spreadsheet->toArray();
            if ($key != 0) {
                unset($data[0]);
            }
            $results = array_merge($results, $data);
        }
        $file = new Filesystem;
        $file->cleanDirectory('storage/excels');
        $export = new Excel(collect($results));
        return \Excel::download($export, 'lienxinh.xlsx');
    }
}
