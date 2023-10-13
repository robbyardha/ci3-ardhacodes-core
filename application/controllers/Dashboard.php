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
        $this->load->view('be/layout/header', $data);
        $this->load->view('be/layout/sidebar', $data);
        $this->load->view('be/layout/navbar', $data);
        $this->load->view('be/dashboard/index', $data);
        $this->load->view('be/layout/footer', $data);
        $this->load->view('be/layout/script', $data);
    }
}
