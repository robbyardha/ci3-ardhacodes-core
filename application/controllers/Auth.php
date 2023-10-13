<?php

class Auth extends CI_Controller
{

    public function __construct()
    {
        parent::__construct();
    }

    public function index()
    {
        $data['title'] = "Auth - Login";
        $this->load->view('auth/layout/header', $data);
        $this->load->view('auth/content/login', $data);
        $this->load->view('auth/layout/footer', $data);
    }

    private function doLogin()
    {
    }

    public function register()
    {
        $data['title'] = "Auth - Register";
        $this->load->view('auth/layout/header', $data);
        $this->load->view('auth/content/register', $data);
        $this->load->view('auth/layout/footer', $data);
    }
}
