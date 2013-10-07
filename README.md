kk_xls
======

Allows to download xls file from your Model or custom array

#Usage
You have two options of using this plugin.

* Simple - just add button to your page.
* Custom - you can specify how, where and when to use it.

##Data
1. You have to use associative array or Model(don\'t fetch it) as source:

        //custom array structure
        $m = array(
            array(
                'field_one' =>'qwe',
                'field_two' => 'asd'),
            array(
                'field_one' =>'zxc',
                'field_two' => 'cxz'
            )
            array(
                'field_one' =>'mmm',
                'field_two' => 'jjj'
            )
        );

2. You may specify custom properties otherwise it\'ll display defauld properties:

        //Default properties are:
        $properties = array(
            'creator'     => 'ATK4 addon kk_xls',
            'lastModifiedBy' => 'kk_xls',
            'title'       => 'ATK4 addon kk_xls',
            'subject'     => 'xls data',
            'description' => 'This file has been generated via kk_xls addon for ATK4',
            'keywords'    => 'atk4, agiletoolkit, addons, kk_xls',
            'category'    => 'data'
        );
    
3. You may specify which filelds to display otherwise it\'ll display all fields:

        $fields = array('field_one','field_one'); 

4. You may define width of columns:

        $fields_width = array(15, 30, 10, 14, 12, 12, 15);

5. You may add TOTALs by specifying names of fields:

        $count_totals = array('field_one');

## Simple usage

    $m = $this->add('Model_User');
    $this->add('kk_xls\View_ButtonXLS',array(
        'data'       => $m //<<<<< Your array or Model
    ));

This will produce simple button.

## Cutom usage

    $m = $this->add('Model_User');
    $xls = $this->add('kk_xls/Controller_XLS');
    $xls->getXLS($m, $properties, $fields, $fields_width, $count_totals);

This will generate a file.
    