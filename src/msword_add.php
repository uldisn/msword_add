<?php

class msword_add {
    
    const SETTING_FILE = 'word/settings.xml';
    
    private $zip;
    private $_path;
    
    public $settings;
    
    
    public function __construct() {
        $this->zip = new ZipArchive();
    }

    /**
     * read word file
     * @param string $path path to word document
     * @return boolean
     */
    public function open($path){
        if(!$this->zip->open($path)){
            return false;
        }
        $this->settings = $this->getFile(self::SETTING_FILE);
        if(!$this->settings){
            return false;
        }
        
        return true;
    }
    
    /**
     * read subfile (from archive)
     * @param string $file
     * @return boolean
     */
    public function getFile($file){
        return $this->zip->getFromName($file);
    }
    
    /**
     * update subfile
     * @param string $file subfile name
     * @param string $data data for saving in subfile
     * @return type
     */
    public function updateFile($file,$data){
        return $this->zip->deleteName($file)
                || $this->zip->addFromString($file, $data);        
    }
    
    /**
     * save doc file (includes ssaving settings)
     * @return boolean
     */
    public function save() {
        return $this->updateFile(self::SETTING_FILE, $this->settings)
            || $this->zip->close();
    }
    
    /**
     * set readonly word file
     */
    public function setReadOnly(){
        $tag_data = $this->getXmlTagData('w:documentProtection',$this->settings);
        if(!$tag_data){
            $tag_data = '';
        }
        $this->setXmlTagAtribute('w:edit','readOnly',$tag_data);
        $this->setXmlTagAtribute('w:enforcement','1',$tag_data);
        $this->setXmlTagData('w:settings','w:documentProtection',$tag_data,$this->settings);     
        
    }
    
    /**
     * read html tag without data
     * @param string $tag tag name
     * @param string $data xml data
     * @return boolean
     */
    public function getXmlTagData($tag, $data){
        if(!preg_match('#<'.$tag.' ([^>])+#',$data,$match)){
            return false;
        }
        
        return $match[1];
    }
    
    /**
     * if exist tag - update or add tag with attributes
     * @param string $insert_after parrent tag, where add tag, if noexist tag
     * @param string $tag tag for update
     * @param string $tag_data tag attributes
     * @param string $data xml
     * @return boolean
     */
    public function setXmlTagData($insert_after,$tag, $tag_data,&$data ){
        if(preg_match('#<'.$tag.' ([^>])+#',$data,$match)){
            $data = preg_replace('#<'.$tag.' [^>]+#','<' . $tag . ' ' . $tag_data,$data);
            return ;
        }
        $replace='$0<'.$tag. ' ' . $tag_data . '/>';
        var_dump($replace);
        $data = preg_replace('#<'.$insert_after.'[^>]+>#',$replace,$data);
        return true;     
        
    }
    
    /**
     * to tag add/update attributes
     * @param type $name
     * @param type $value
     * @param type $tag_data
     */
    public function setXmlTagAtribute($name,$value,&$tag_data){
        if(preg_match('#'.$name.'="[^"]+"#',$tag_data)){
            $tag_data = preg_replace('#'.$name.'="([^"])+"#',$name.'=".$value."',$tag_data);
        }else{
            $tag_data = trim($tag_data) . ' ' . $name . '="' . $value . '"';
        }
    }
    
}