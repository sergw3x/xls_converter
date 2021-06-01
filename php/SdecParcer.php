<?php


use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing;

class SdecParcer
{
    private array $map;
    private array $table;
    private string $catName;
    private string $catDesc;
    private object $reader;

    public function __construct($file)
    {
        $this->FILE = $file;
        $this->table   = [];
        $this->map     = [];
        $this->catName = '';
        $this->catDesc = '';

        $this->reader = IOFactory::load($this->FILE);
    }

    function readAll()
    {
        $sheets = $this->reader->getAllSheets();

        foreach ($sheets as $sheet) {
            $this->readContentSheet($sheet->getTitle());
            break;
        }
    }

    function readContentSheet(string $sheetName){

        $sheet = $this->reader->getSheetByName($sheetName);

        if ($sheet === null){
            throw new \PhpOffice\PhpSpreadsheet\Exception("sheet is null");
        }

        $prevCodeRange = '';
        $prevCode = '';

        $cells = $sheet->getCellCollection();
        for ($row = 1; $row <= $cells->getHighestRow(); $row++) {
            $tabRow = [];
            for ($col = 'A'; $col <= $cells->getHighestColumn(); $col++) {
                $cell = $cells->get($col . $row);
                if ($cell && $cell->getValue() !== null) {
                    $cellValue = $cell->getValue();
                    if (gettype($cellValue) == 'object') {
                        $cellValue = $cellValue->getPlainText();
                    }
                    if (!$this->catName) {
                        $this->catName = $cellValue;
                        continue;
                    }

                    if (!$this->catDesc) {
                        $this->catDesc = $cellValue;
                        continue;
                    }

                    if (str_contains($cellValue, "Group Number")) {
                        $this->map[$col] = 'GroupNumber';
                        continue;
                    } elseif (str_contains($cellValue, "Chinese Description")) {
                        $this->map[$col] = 'ChineseDescription';
                        continue;
                    } elseif (str_contains($cellValue, "English Description")) {
                        $this->map[$col] = 'EnglishDescription';
                        continue;
                    } elseif (str_contains($cellValue, "Code")) {
                        $this->map[$col] = 'Code';
                        continue;
                    }

                    if (array_key_exists($col, $this->map)) {

                        if ($cell->getMergeRange()) {
                            if ($prevCodeRange == '' || $prevCodeRange != $cell->getMergeRange()) {
                                $prevCodeRange = $cell->getMergeRange();
                                $prevCode = $cellValue;
                                $tabRow[$this->map[$col]] = $cellValue;
                            }
                        } else {
                            $tabRow[$this->map[$col]] = $cellValue;
                        }
                    } else {
                        echo $cell . $row . ' ' . $cellValue . PHP_EOL;
                    }
                } else {
                    if ($prevCodeRange && array_key_exists($col, $this->map)) {
                        $max = explode(':', $prevCodeRange)[1]; // G9
                        if ($max == $col . $row) { // G5 == G9
                            $tabRow[$this->map[$col]] = $prevCode;
                            $prevCode = '';
                        } else {
                            $tabRow[$this->map[$col]] = $prevCode;
                        }
                    }
                }
            }
            if (!$this->isEmptyArray($tabRow)) {
                $this->table[] = $tabRow;
            }
        }

        return $this->table;
    }

    function readSheet(string $sheetName){

        $sheet = $this->reader->getSheetByName($sheetName);

        if ($sheet === null){
            throw new \PhpOffice\PhpSpreadsheet\Exception("sheet is null");
        }

        $prevCodeRange = '';
        $prevCode = '';

        $cells = $sheet->getCellCollection();
        for ($row = 1; $row <= $cells->getHighestRow(); $row++) {
            $tabRow = [];
            for ($col = 'A'; $col <= $cells->getHighestColumn(); $col++) {
                $cell = $cells->get($col . $row);
                if ($cell && $cell->getValue() !== null) {
                    $cellValue = trim($cell->getValue());
                    if (!$cellValue) continue;
                    if (gettype($cellValue) == 'object') {
                        $cellValue = $cellValue->getPlainText();
                    }

                    if (str_contains($cellValue, "Ref")) {
                        $this->map[$col] = 'Ref';
                        continue;
                    } elseif (str_contains($cellValue, "Part No")) {
                        $this->map[$col] = 'PartNo';
                        continue;
                    } elseif (str_contains($cellValue, "Chinese Description")) {
                        $this->map[$col] = 'ChineseDescription';
                        continue;
                    } elseif (str_contains($cellValue, "English Description")) {
                        $this->map[$col] = 'EnglishDescription';
                        continue;
                    } elseif (str_contains($cellValue, "Quantity")) {
                        $this->map[$col] = 'Quantity';
                        continue;
                    } elseif (str_contains($cellValue, "Standard Fasteners Sign")) {
                        $this->map[$col] = 'StandardFastenersSign';
                        continue;
                    }

                    if (array_key_exists($col, $this->map)) {
                        if ($cell->getMergeRange()) {
                            if ($prevCodeRange == '' || $prevCodeRange != $cell->getMergeRange()) {
                                $prevCodeRange = $cell->getMergeRange();
                                $prevCode = $cellValue;
                                $tabRow[$this->map[$col]] = $cellValue;
                            }
                        } else {
                            if (str_contains($cellValue, 'Back to') && $cell->getHyperlink()){
                                continue;
                            }else{
                                $tabRow[$this->map[$col]] = $cellValue;
                            }
                        }
                    } else {
                        echo $cell . $row . ' ' . $cellValue . PHP_EOL;
                    }
                } else {
                    if ($prevCodeRange && array_key_exists($col, $this->map)) {
                        $max = explode(':', $prevCodeRange)[1]; // G9
                        if ($max == $col . $row) { // G5 == G9
                            $tabRow[$this->map[$col]] = $prevCode;
                            $prevCode = '';
                        } else {
                            $tabRow[$this->map[$col]] = $prevCode;
                        }
                    }
                }
            }
            if (!$this->isEmptyArray($tabRow)) {
                $this->table[] = $tabRow;
            }
        }


        return $this->table;
    }

    function readImg($sheetName){
        $sheet = $this->reader->getSheetByName($sheetName);

        if ($sheet === null){
            throw new \PhpOffice\PhpSpreadsheet\Exception("sheet is null");
        }

        $i = 0;
        foreach ($sheet->getDrawingCollection() as $drawing) {
            if ($drawing instanceof MemoryDrawing) {
                ob_start();
                call_user_func(
                    $drawing->getRenderingFunction(),
                    $drawing->getImageResource()
                );
                $imageContents = ob_get_contents();
                ob_end_clean();
                switch ($drawing->getMimeType()) {
                    case MemoryDrawing::MIMETYPE_PNG :
                        $extension = 'png';
                        break;
                    case MemoryDrawing::MIMETYPE_GIF:
                        $extension = 'gif';
                        break;
                    case MemoryDrawing::MIMETYPE_JPEG :
                        $extension = 'jpg';
                        break;
                }
            } else {
                $zipReader = fopen($drawing->getPath(),'r');
                $imageContents = '';
                while (!feof($zipReader)) {
                    $imageContents .= fread($zipReader,1024);
                }
                fclose($zipReader);
                $extension = $drawing->getExtension();
            }
            $myFileName = '00_Image_'.++$i.'.'.$extension;
            file_put_contents($myFileName,$imageContents);
        }
    }
    
    function isEmptyArray($tabRow): bool
    {
        foreach (array_keys($tabRow) as $arKey) {
            if (!empty($tabRow[$arKey])) {
                return false;
            }
        }
        return true;
    }


}