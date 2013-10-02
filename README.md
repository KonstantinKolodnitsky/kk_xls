kk_xls
======

Allows to download xls file from your Model or custom array

#Usage
You have two options of using this plugin.

* Simple - just add button to your page.
* Custom - you can specify how, where and when to use it.

##Data
1. You have to use associative array or Model(don\'t fetch it) as source.
2. You may specify custom properties otherwise it\'ll display defauld properties.
3. You may specyfy which filelds to display otherwise it\'ll display all fields.

*Example*

    //custom array
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

    $fields = array('field_one','field_two'); 

    //Default properties are:
    $properties = array(
        'creator'     => 'ATK4 addon kk_xls',
        'title'       => 'ATK4 addon kk_xls',
        'subject'     => 'xls data',
        'description' => 'This file has been generated via kk_xls addon for ATK4',
        'keywords'    => 'atk4, agiletoolkit, addons, kk_xls',
        'category'    => 'data'
    );

## Simple usage

    $m = $this->add('Model_User');
    $this->add('kk_xls\View_ButtonXLS',array(
        'data'       => $m //<<<<< Your array or Model
        'properties' => $properties, // optional
        'fields'     => $fields // optional
    ));

## Cutom
This stuff is in progress