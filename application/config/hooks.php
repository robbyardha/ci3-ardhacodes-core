<?php
defined('BASEPATH') or exit('No direct script access allowed');

require FCPATH . 'vendor/autoload.php';

use Dotenv\Dotenv;
/*
| -------------------------------------------------------------------------
| Hooks
| -------------------------------------------------------------------------
| This file lets you define "hooks" to extend CI without hacking the core
| files.  Please see the user guide for info:
|
|	https://codeigniter.com/userguide3/general/hooks.html
|
*/


$hook['pre_system'][] = array(
    'class'    => 'maintenance_hook',
    'function' => 'offline_check',
    'filename' => 'maintenance_hook.php',
    'filepath' => 'hooks',
);


$hook['pre_system'] = function () {
    /**
     * APPPATH =>  'C:\xampp\htdocs\tutor-in\application\'
     * __DIR__ =>  'C:\xampp\htdocs\tutor-in\application\controllers'
     * getcwd() =>  'C:\xampp\htdocs\tutor-in'
     */

    //TODO : FROM .env Path /application
    // $dotenv = Dotenv::create(APPPATH);
    // $dotenv->load();

    //TODO : FROM .env Path /tutor-in
    // $dotenv = Dotenv::create(getcwd());
    // $dotenv->load();

    // $dotenv = Dotenv\Dotenv::createImmutable(getcwd());

    $dotenv = Dotenv::createImmutable(getcwd());
    $dotenv->load();
};
