<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as ReaderXlsx;
use PhpOffice\PhpSpreadsheet\Reader\Xls as ReaderXls;


class Excel_library
{

    protected $CI;
    private $spreadsheet;
    private $export;
    private $import;

    public function __construct($data)
    {
        $this->CI = &get_instance();

        $this->spreadsheet = new Spreadsheet();

        $this->case($data);
    }

    public function case($data)
    {
        if (is_array($data) && !empty($data)) {

            switch ($data['func']) {
                case 'export':
                    $this->export($data);
                    break;
                case 'import':
                    $this->import($data);
                default:
                    return;
                    break;
            }
        } else {
            return;
        }
    }

    private function import($data)
    {

        if ($data['format_data'] == "xlsx") {

            $this->import = new ReaderXlsx;
        } else if ($data['format_data'] == "xls") {

            $this->import = new ReaderXls;
        }

        $spreadsheetReader = $this->import->load($data['realpath']);
        $result = $spreadsheetReader->getActiveSheet()->toArray();

        switch ($data['jenis']) {

            case "prapendaftar":

                $status_bayar = ["BAYAR", "BELUM"];

                if (!empty($result)) {

                    if (count($result[0]) == 8) {

                        $row            = [];
                        $row_error      = 0;
                        $cell           = 2;

                        for ($i = 1; $i < count($result); $i++) {

                            $col_error    = "";

                            $kelas_id = $this->CI->kelas->cekKelas($result[$i][1]);

                            // JALUR
                            if ($result[$i][1]) {

                                if ($kelas_id) {

                                    $result[$i][1] = $kelas_id["id"];
                                } else {

                                    $col_error .= "jalur " . $result[$i][1] . " tidak ada, ";
                                }
                            } else {

                                $col_error .= "jalur tidak boleh kosong, ";
                            }

                            // NAMA
                            if (!$result[$i][2]) {

                                $col_error .= "nama tidak boleh kosong, ";
                            }

                            // HP
                            if ($result[$i][3]) {

                                if (strlen($result[$i][3]) < 10) {

                                    $col_error .= "nomor hp minimal 10, ";
                                } else {

                                    if (!is_numeric($result[$i][3])) {

                                        $col_error .= "nomor hp harus berupa angka, ";
                                    }
                                }
                            } else {

                                $col_error .= "nomor hp tidak boleh kosong, ";
                            }

                            // TGL_LAHIR
                            if ($result[$i][4]) {

                                $result[$i][4] = date("Y-m-d", strtotime($result[$i][4]));
                            } else {

                                $col_error .= "tanggal lahir tidak boleh kosong, ";
                            }

                            // TGL DAFTAR
                            if ($result[$i][5]) {

                                $result[$i][5] = date("Y-m-d", strtotime($result[$i][5]));

                                if (!$this->CI->tagihan->ambilGelombangBerdasarkanTglSekarang($result[$i][5], $result[$i][1])) {

                                    $col_error .= "tagihan belum di set pada tanggal pendaftaran tersebut, ";
                                }
                            } else {

                                $col_error .= "tanggal daftar tidak boleh kosong, ";
                            }

                            // STATUS BAYAR
                            if ($result[$i][6]) {

                                if (in_array(strtoupper($result[$i][6]), $status_bayar)) {

                                    $result[$i][6]  = strtoupper($result[$i][6]);
                                } else {

                                    $col_error .= "status tidak tersedia, ";
                                }
                            } else {

                                $col_error .= "status tidak boleh kosong, ";
                            }

                            if ($col_error) {

                                $result[$i][8]  = $col_error;
                                $row_error++;
                            } else {

                                $result[$i][8]  = "valid";
                            }

                            array_push($row, $result[$i]);

                            $cell++;
                        }

                        if ($row_error) {

                            // set error
                            $this->CI->session->set_userdata("error_prapendaftaran", $row);
                            redirect("peserta/prapendaftaran", "refresh");
                        } else {

                            // unset error
                            $this->CI->session->unset_userdata("error_prapendaftaran");

                            // eksekusi semua baris success
                            if ($this->CI->prapendaftaran->daftarKanSemua($row, $data["wa"])) {

                                $this->CI->session->set_flashdata("success", "Data berhasil di import");
                                redirect("peserta/prapendaftaran");
                            } else {

                                $this->CI->session->set_flashdata("error", "Data gagal di import");
                                redirect("peserta/prapendaftaran");
                            }
                        }
                    } else {

                        $this->CI->session->set_flashdata("error", "Kolom excel tidak sesuai");
                        redirect("peserta/prapendaftaran", "refresh");
                    }
                } else {

                    $this->CI->session->set_flashdata("error", "Data kosong");
                    redirect("peserta/prapendaftaran", "refresh");
                }

                break;

            case "nilai":
                if (!empty($result)) {

                    if (array_key_exists(2, $result) !== FALSE) {

                        $row_error   = [];
                        $row_success = [];

                        for ($i = 2; $i < count($result); $i++) {

                            $col_error   = [];
                            $col_success = [];

                            if ($result[$i][1]) {
                                // cek nocust ada atau tidak 
                                $data_daftar = $this->CI->peserta->cekNocustDiDaftar($result[$i][1]);

                                if ($data_daftar) {

                                    $col_success["daftar_id"] = $data_daftar["daftar_id"];

                                    // jika ada maka 
                                    if (is_numeric($result[$i][4])) {

                                        $col_success["n_wawancara"] = $result[$i][4];
                                    } else {

                                        $col_error[] = "Nilai wawancara pada cell " . ($i + 1) . " hanya boleh angka";
                                    }

                                    if (is_numeric($result[$i][5])) {

                                        $col_success["n_bhs_ing"] = $result[$i][5];
                                    } else {

                                        $col_error[] = "Nilai bahasa inggris pada cell " . ($i + 1) . " hanya boleh angka";
                                    }

                                    if (is_numeric($result[$i][6])) {

                                        $col_success["n_bhs_indo"] = $result[$i][6];
                                    } else {

                                        $col_error[] = "Nilai bahasa indonesia pada cell " . ($i + 1) . " hanya boleh angka";
                                    }

                                    if (is_numeric($result[$i][7])) {

                                        $col_success["n_berhitung"] = $result[$i][7];
                                    } else {

                                        $col_error[] = "Nilai berhitung pada cell " . ($i + 1) . " hanya boleh angka";
                                    }

                                    if (is_numeric($result[$i][8])) {

                                        $col_success["n_mewarnai"] = $result[$i][8];
                                    } else {

                                        $col_error[] = "Nilai mewarnai pada cell " . ($i + 1) . " hanya boleh angka";
                                    }

                                    if ($result[$i][11]) {

                                        $col_success["catatan"] = $result[$i][11];
                                    }
                                } else {

                                    // nocust tidak ada 
                                    $col_error[] = "Nocust pada cell " . ($i + 1) . " tidak ditemukan";
                                }
                            } else {

                                // nocust kosong
                                $col_error[] = "Nocust pada cell " . ($i + 1) . " tidak boleh kosong";
                            }

                            if ($col_error) {

                                array_push($row_error, $col_error);
                            }

                            if ($col_success) {

                                array_push($row_success, $col_success);
                            }
                        }

                        if ($row_error) {

                            $this->CI->session->set_flashdata("error_array", $row_error);
                            redirect("peserta/terdaftar", "refresh");
                        } else {

                            if ($this->CI->nilai->eksekusiNilaiBerdasarkanDaftarId($row_success)) {

                                $this->CI->session->set_flashdata("success", "Data berhasil di perbarui");
                            } else {

                                $this->CI->session->set_flashdata("error", "Data gagal di perbarui");
                            }

                            redirect("peserta/terdaftar", "refresh");
                        }
                    } else {

                        $this->CI->session->set_flashdata("error", "Data kosong");
                        redirect("peserta/terdaftar", "refresh");
                    }
                } else {

                    $this->CI->session->set_flashdata("error", "Data kosong");
                    redirect("peserta/terdaftar", "refresh");
                }
                break;
        }
    }

    private function export($data)
    {

        $activeWorksheet = $this->spreadsheet->getActiveSheet();

        $activeWorksheet->getColumnDimension('A')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('B')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('C')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('D')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('E')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('F')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('G')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('H')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('I')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('J')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('K')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('L')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('M')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('N')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('O')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('P')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('Q')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('R')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('S')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('T')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('U')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('V')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('W')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('X')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('Y')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('Z')->setAutoSize(true);

        $activeWorksheet->getColumnDimension('AA')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AB')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AC')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AD')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AE')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AF')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AG')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AH')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AI')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AJ')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AK')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AL')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AM')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AN')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AO')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AP')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AQ')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AR')->setAutoSize(true);
        $activeWorksheet->getColumnDimension('AS')->setAutoSize(true);

        $styleArray = [

            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'FFE5E4' // Kode warna RGB (misalnya, kuning)
                ]
            ]
        ];

        $styleArray1 = [
            'font' => [
                'bold' => true,
                'size' => 13
            ],
        ];

        switch ($data["jenis"]) {

            case "tagihan":

                $activeWorksheet->setCellValue("A1", "NO");
                $activeWorksheet->setCellValue("B1", "NIS");
                $activeWorksheet->setCellvalue("C1", "POS_TAGIHAN");
                $activeWorksheet->setCellvalue("D1", "NAMA_TAGIHAN");
                $activeWorksheet->setCellvalue("E1", "NOMINAL");
                $activeWorksheet->setCellvalue("F1", "POTONGAN");
                $activeWorksheet->setCellvalue("G1", "JENIS_TAGIHAN");
                $activeWorksheet->setCellvalue("H1", "TAHUN_AJARAN");
                $activeWorksheet->setCellvalue("I1", "BULAN");
                $activeWorksheet->setCellvalue("J1", "TGL_BAYAR");
                $activeWorksheet->setCellvalue("K1", "LOKET");

                $activeWorksheet->getStyle("A1:K1")->applyFromArray($styleArray);

                $urutan = 1;
                $cell = 2;
                foreach ($data["data_tagihan"] as $dt) {

                    $activeWorksheet->setCellValue("A$cell", $urutan);
                    $activeWorksheet->setCellValue("B$cell", $dt["nocust"]);
                    $activeWorksheet->setCellValue("C$cell", $dt["pos_tagihan"]);
                    $activeWorksheet->setCellValue("D$cell", $dt["billnm"]);
                    $activeWorksheet->setCellValue("E$cell", str_replace(".", "", $dt["billam"]));
                    $activeWorksheet->setCellValue("F$cell", 0);
                    $activeWorksheet->setCellValue("G$cell", $dt["jenis"]);
                    $activeWorksheet->setCellValue("H$cell", $dt["periode"]);
                    $activeWorksheet->setCellValue("I$cell", 0);
                    $activeWorksheet->setCellValue("J$cell", date("d/m/Y", strtotime($dt["paiddt"])));
                    $activeWorksheet->setCellValue("K$cell", $dt["user_input"]);

                    $urutan++;
                    $cell++;
                }

                break;

            case "nilai":

                $activeWorksheet->setCellValue("A1", "NO");
                $activeWorksheet->setCellValue("B1", "No formulir calon siswa");
                $activeWorksheet->setCellValue("C1", "Nama calon siswa");
                $activeWorksheet->setCellvalue("D1", "Program");
                $activeWorksheet->setCellvalue("E1", "Rekap nilai");
                $activeWorksheet->setCellValue("E2", "Wawancara (18)");
                $activeWorksheet->setCellValue("F2", "Bahasa inggris(20)");
                $activeWorksheet->setCellValue("G2", "Bahasa indonesia(30)");
                $activeWorksheet->setCellValue("H2", "Berhitung(15)");
                $activeWorksheet->setCellValue("I2", "Mewarnai(17)");
                $activeWorksheet->setCellvalue("J1", "Total nilai");
                $activeWorksheet->setCellvalue("K1", "Program rekomendasi");
                $activeWorksheet->setCellvalue("L1", "Catatan");

                $activeWorksheet->mergeCells("A1:A2");
                $activeWorksheet->mergeCells("B1:B2");
                $activeWorksheet->mergeCells("C1:C2");
                $activeWorksheet->mergeCells("D1:D2");
                $activeWorksheet->mergeCells("E1:I1");
                $activeWorksheet->mergeCells("J1:J2");
                $activeWorksheet->mergeCells("K1:K2");
                $activeWorksheet->mergeCells("L1:L2");

                $activeWorksheet->getStyle("A1:L1")->applyFromArray($styleArray);
                $activeWorksheet->getStyle("E2:I2")->applyFromArray($styleArray);
                $activeWorksheet->getStyle("A1:L1")->applyFromArray($styleArray1);
                $activeWorksheet->getStyle("E2:I2")->applyFromArray($styleArray1);
                $activeWorksheet->getStyle('A:L')->getAlignment()->setHorizontal('center');
                $activeWorksheet->getStyle('A:L')->getAlignment()->setVertical('center');

                $urutan = 1;
                $cell = 3;

                foreach ($data["data_nilai"] as $dn) {

                    $nilaiTotal = (($dn["n_wawancara"]) ? $dn["n_wawancara"] : 0) + (($dn["n_bhs_ing"]) ? $dn["n_bhs_ing"] : 0) + (($dn["n_bhs_indo"]) ? $dn["n_bhs_indo"] : 0) + (($dn["n_berhitung"]) ? $dn["n_berhitung"] : 0) + (($dn["n_mewarnai"]) ? $dn["n_mewarnai"] : 0);

                    $activeWorksheet->setCellValue("A$cell", $urutan);
                    $activeWorksheet->setCellValue("B$cell", $dn["nocust"]);
                    $activeWorksheet->setCellvalue("C$cell", $dn["nama"]);
                    $activeWorksheet->setCellvalue("D$cell", $dn["kelas"]);
                    $activeWorksheet->setCellValue("E$cell", ($dn["n_wawancara"] !== NULL) ? $dn["n_wawancara"] : 0);
                    $activeWorksheet->setCellValue("F$cell", ($dn["n_bhs_ing"] !== NULL) ? $dn["n_bhs_ing"] : 0);
                    $activeWorksheet->setCellValue("G$cell", ($dn["n_bhs_indo"] !== NULL) ? $dn["n_bhs_indo"] : 0);
                    $activeWorksheet->setCellValue("H$cell", ($dn["n_berhitung"] !== NULL) ? $dn["n_berhitung"] : 0);
                    $activeWorksheet->setCellValue("I$cell", ($dn["n_mewarnai"] !== NULL) ? $dn["n_mewarnai"] : 0);
                    $activeWorksheet->setCellvalue("J$cell", $nilaiTotal);
                    $activeWorksheet->setCellvalue("K$cell", ($dn["n_bhs_ing"] !== NULL && $dn["n_bhs_ing"] >= 11) ? "CIP" : "REGULER");
                    $activeWorksheet->setCellvalue("L$cell", $dn["catatan"]);

                    $urutan++;
                    $cell++;
                }

                break;

            case "peserta":

                $activeWorksheet->setCellValue("A1", "Data diri");
                $activeWorksheet->setCellValue("Y1", "Data ayah");
                $activeWorksheet->setCellValue("AF1", "Data ibu");
                $activeWorksheet->setCellValue("AM1", "Data wali");

                $activeWorksheet->setCellValue("A2", "No");
                $activeWorksheet->setCellValue("B2", "Nocust");
                $activeWorksheet->setCellValue("C2", "Nova");
                $activeWorksheet->setCellValue("D2", "Nama");
                $activeWorksheet->setCellValue("E2", "Tahun angkatan");
                $activeWorksheet->setCellValue("F2", "Nik");
                $activeWorksheet->setCellValue("G2", "Nokk");
                $activeWorksheet->setCellValue("H2", "Nisn");
                $activeWorksheet->setCellValue("I2", "Jenis kelamin");
                $activeWorksheet->setCellValue("J2", "Kategori sekolah");
                $activeWorksheet->setCellValue("K2", "Asal sekolah");
                $activeWorksheet->setCellValue("L2", "Alamat sekolah");
                $activeWorksheet->setCellValue("M2", "Tempat lahir");
                $activeWorksheet->setCellValue("N2", "Tanggal lahir");
                $activeWorksheet->setCellValue("O2", "Alamat domisili");
                $activeWorksheet->setCellValue("P2", "Alamat KK");
                $activeWorksheet->setCellValue("Q2", "Kecamatan");
                $activeWorksheet->setCellValue("R2", "Desa");
                $activeWorksheet->setCellValue("S2", "Kabupaten");
                $activeWorksheet->setCellValue("T2", "Email");
                $activeWorksheet->setCellValue("U2", "Hp");
                $activeWorksheet->setCellValue("V2", "Anakke");
                $activeWorksheet->setCellValue("W2", "Keberadaan anak");
                $activeWorksheet->setCellValue("X2", "Jumlah saudara");

                $activeWorksheet->setCellValue("Y2", "Nik");
                $activeWorksheet->setCellValue("Z2", "Nama");
                $activeWorksheet->setCellValue("AA2", "Hp");
                $activeWorksheet->setCellValue("AB2", "Kerja");
                $activeWorksheet->setCellValue("AC2", "Gaji");
                $activeWorksheet->setCellValue("AD2", "Pendidikan terakhir");
                $activeWorksheet->setCellValue("AE2", "Tahun lahir");

                $activeWorksheet->setCellValue("AF2", "Nik");
                $activeWorksheet->setCellValue("AG2", "Nama");
                $activeWorksheet->setCellValue("AH2", "Hp");
                $activeWorksheet->setCellValue("AI2", "Kerja");
                $activeWorksheet->setCellValue("AJ2", "Gaji");
                $activeWorksheet->setCellValue("AK2", "Pendidikan terakhir");
                $activeWorksheet->setCellValue("AL2", "Tahun lahir");

                $activeWorksheet->setCellValue("AM2", "Nik");
                $activeWorksheet->setCellValue("AN2", "Nama");
                $activeWorksheet->setCellValue("AO2", "Hp");
                $activeWorksheet->setCellValue("AP2", "Kerja");
                $activeWorksheet->setCellValue("AQ2", "Gaji");
                $activeWorksheet->setCellValue("AR2", "Pendidikan terakhir");
                $activeWorksheet->setCellValue("AS2", "Tahun lahir");

                $activeWorksheet->mergeCells("A1:X1");
                $activeWorksheet->mergeCells("Y1:AE1");
                $activeWorksheet->mergeCells("AF1:AL1");
                $activeWorksheet->mergeCells("AM1:AS1");

                $activeWorksheet->getStyle("A1:AS1")->applyFromArray($styleArray);
                $activeWorksheet->getStyle("A2:AS2")->applyFromArray($styleArray);
                $activeWorksheet->getStyle("A1:AS1")->applyFromArray($styleArray1);
                $activeWorksheet->getStyle("A2:AS2")->applyFromArray($styleArray1);
                $activeWorksheet->getStyle('A:AS')->getAlignment()->setHorizontal('center');
                $activeWorksheet->getStyle('A:AS')->getAlignment()->setVertical('center');

                $urutan = 1;
                $cell = 3;

                foreach ($data["data_peserta"] as $dp) {

                    $activeWorksheet->setCellValue("A$cell", $urutan);
                    $activeWorksheet->setCellValue("B$cell", $dp["nocust"]);
                    $activeWorksheet->setCellValueExplicit("C$cell", $dp["nova"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValue("D$cell", $dp["nama"]);
                    $activeWorksheet->setCellValue("E$cell", $dp["th_angkatan"]);
                    $activeWorksheet->setCellValueExplicit("F$cell", $dp["nik"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValueExplicit("G$cell", $dp["nokk"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValueExplicit("H$cell", $dp["nisn"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValue("I$cell", $dp["nocust"]);
                    $activeWorksheet->setCellValue("J$cell", $dp["kategori_skl"]);
                    $activeWorksheet->setCellValue("K$cell", $dp["asal_skl"]);
                    $activeWorksheet->setCellValue("L$cell", $dp["alamat_skl"]);
                    $activeWorksheet->setCellValue("M$cell", $dp["tmp_lhr"]);
                    $activeWorksheet->setCellValue("N$cell", $dp["tgl_lhr"]);
                    $activeWorksheet->setCellValue("O$cell", $dp["alamat"]);
                    $activeWorksheet->setCellValue("P$cell", $dp["alamat_kk"]);
                    $activeWorksheet->setCellValue("Q$cell", $dp["kecamatan"]);
                    $activeWorksheet->setCellValue("R$cell", $dp["desa"]);
                    $activeWorksheet->setCellValue("S$cell", $dp["kabupaten"]);
                    $activeWorksheet->setCellValue("T$cell", $dp["email"]);
                    $activeWorksheet->setCellValueExplicit("U$cell", $dp["hp"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValue("V$cell", $dp["anakke"]);
                    $activeWorksheet->setCellValue("W$cell", $dp["keberadaan_anak"]);
                    $activeWorksheet->setCellValue("X$cell", $dp["saudara"]);

                    $activeWorksheet->setCellValueExplicit("Y$cell", $dp["nik_ayah"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValue("Z$cell", $dp["nama_ayah"]);
                    $activeWorksheet->setCellValueExplicit("AA$cell", $dp["hp_ayah"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValue("AB$cell", $dp["kerja_ayah"]);
                    $activeWorksheet->setCellValue("AC$cell", $dp["gaji_ayah"]);
                    $activeWorksheet->setCellValue("AD$cell", $dp["pendidikan_ayah"]);
                    $activeWorksheet->setCellValue("AE$cell", ($dp["th_lhr_ayah"]) ? $dp["th_lhr_ayah"] : "-");

                    $activeWorksheet->setCellValueExplicit("AF$cell", $dp["nik_ibu"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValue("AG$cell", $dp["nama_ibu"]);
                    $activeWorksheet->setCellValueExplicit("AH$cell", $dp["hp_ibu"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValue("AI$cell", $dp["kerja_ibu"]);
                    $activeWorksheet->setCellValue("AJ$cell", $dp["gaji_ibu"]);
                    $activeWorksheet->setCellValue("AK$cell", $dp["pendidikan_ibu"]);
                    $activeWorksheet->setCellValue("AL$cell", ($dp["th_lhr_ibu"]) ? $dp["th_lhr_ibu"] : "-");

                    $activeWorksheet->setCellValueExplicit("AM$cell", $dp["nik_wali"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValue("AN$cell", $dp["nama_wali"]);
                    $activeWorksheet->setCellValueExplicit("AO$cell", $dp["hp_wali"], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                    $activeWorksheet->setCellValue("AP$cell", $dp["kerja_wali"]);
                    $activeWorksheet->setCellValue("AQ$cell", $dp["gaji_wali"]);
                    $activeWorksheet->setCellValue("AR$cell", $dp["pendidikan_wali"]);
                    $activeWorksheet->setCellValue("AS$cell", ($dp["th_lhr_wali"]) ? $dp["th_lhr_wali"] : "-");

                    $urutan++;
                    $cell++;
                }

                break;
        }

        $this->export = new PhpOffice\PhpSpreadsheet\Writer\Xlsx($this->spreadsheet);

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . rawurlencode($data['filename']) . '"');

        ob_end_clean();
        $this->export->save('php://output');
        exit;
    }
}
