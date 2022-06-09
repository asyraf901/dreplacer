<?php
include_once 'lib/docx-replacer/Docx.php';
include_once 'lib/SimpleXLSX.php';

class Dreplacer
{
    var $fileDir = 'temp/';
    var $dataFilePath;
    var $templateFilePath;
    var $templateFileInfo;
    var $newFileDir;
    var $newFilePath;
    var $zipFilePath;
    var $dReplacerArr;

    var $zipArchiveOptions = array(
        'remove_all_path' => TRUE
    );

    private function downloadZip(){
        if(file_exists($this->zipFilePath)) {
            $fileInfo = pathinfo($this->zipFilePath);
    
            header("Content-type: application/zip"); 
            header("Content-Disposition: attachment; filename=".$fileInfo['basename']);
            header("Content-length: " . filesize($filepath));
            header("Pragma: no-cache"); 
            header("Expires: 0"); 
            readfile($this->zipFilePath);
        } else {
            $this->_error('Error download zip.');
        }
    }
    
    private function verifyFile(){
        if (!file_exists($this->templateFilePath)) $this->_error('Template file not exists!');
        if (!file_exists($this->dataFilePath)) $this->_error('Data file not exists!');
    }

    private function initVar(){
        $this->templateFileInfo = pathinfo($this->templateFilePath);
        
        $d = $this->fileDir.$this->templateFileInfo['filename'];
        $this->newFileDir =  $d.'/';
        $this->zipFilePath = $d.'.zip';
    }

    private function genNewFilePath($cnt){
        $this->checkDir();
        $this->newFilePath = $this->newFileDir.$this->templateFileInfo['filename'].'_'.$cnt.'.'.$this->templateFileInfo['extension'];
    }

    private function checkDir(){
        if(!file_exists($this->fileDir)) mkdir($this->fileDir, '0777');
        if(!file_exists($this->newFileDir)) mkdir($this->newFileDir, '0777');
    }
    
    private function replacer(){
        copy($this->templateFilePath, $this->newFilePath);
        $docx = new \IRebega\DocxReplacer\Docx($this->newFilePath);
        foreach ($this->dReplacerArr as $k => $v) {
            $docx->replaceTextInsensitive($k, $v);
        }
        return;
    }

    private function generateZip(){
        $zip = new ZipArchive();
        $zip->open($this->zipFilePath, ZipArchive::CREATE | ZipArchive::OVERWRITE);
        $options = array('remove_all_path' => TRUE);
        $zip->addGlob($this->newFileDir.'*', 0, $this->zipArchiveOptions);
        if ($zip->status != ZIPARCHIVE::ER_OK) $this->_error('Failed to write files to zip.');
        $zip->close();
    }

    private function procesFile(){
        if ($xlsx = SimpleXLSX::parse($this->dataFilePath) ) {
            $allRow = $xlsx->rows();
            foreach ($allRow  as $r => $row ) {
                if ($r === 0) continue;
                $this->dReplacerArr = array();
                foreach ($row as $m => $n) {
                    $this->dReplacerArr[$allRow[0][$m]] = $n;
                }
                $this->genNewFilePath($r);
                $this->replacer();
            }
            return;
        }else{
            $this->_error('Error parsing data file.');
            return;
        }
    }

    public function process($templateFilePath, $dataFilePath, $download = true){
        $this->dataFilePath = $dataFilePath;
        $this->templateFilePath = $templateFilePath;
        
        $this->verifyFile();
        $this->initVar();
        $this->procesFile();
        $this->generateZip();
        if($download) $this->downloadZip();
    }

    private function _error($msg){
        throw new Exception($msg);
        return;
    }
}
