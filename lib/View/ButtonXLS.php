<?php
/**
 *Created by Konstantin Kolodnitsky
 * Date: 02.10.13
 * Time: 10:34
 */
namespace KonstantinKolodnitsky\kk_xls;
class View_ButtonXLS extends \View_Button{
    public $data;
    public $label = 'Get XLS';
    public $fields = null;
    public $fields_width =null;
    public $count_totals = null;
    public $properties = array(
        'creator' => 'ATK4 addon kk_xls',
        'lastModifiedBy' => 'kk_xls',
        'title' => 'ATK4 addon kk_xls',
        'subject' => 'xls data',
        'description' => 'This file has been generated via kk_xls addon for ATK4',
        'keywords' => 'atk4, agiletoolkit, addons, kk_xls',
        'category' => 'data'
    );
    function init(){
        parent::init();

        $this->set($this->label);

        $xls = $this->add('KonstantinKolodnitsky/kk_xls/Controller_XLS');

        $this->js('click',$this->js()->univ()->redirect($this->api->url(null,array('action'=>'export'))));

        if($_GET['action'] == 'export'){
            $xls->getXLS($this->data, $this->properties, $this->fields, $this->fields_width, $this->count_totals);
        }
    }
}