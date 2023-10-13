<?php

class Dashboard extends CI_Controller
{

    public function __construct()
    {
        parent::__construct();
    }

    public function index()
    {
        $data['title'] = "Dashboard";
        $this->load->view('be/layout/header');
        $this->load->view('be/layout/sidebar');
        $this->load->view('be/layout/navbar');
        $this->load->view('be/dashboard/index');
        $this->load->view('be/layout/footer');
        $this->load->view('be/layout/script');
    }
}
