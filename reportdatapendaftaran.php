<?php
    include 'koneksi.php';
    require 'vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A1', 'No');
    $sheet->setCellValue('B1', 'Jenis Pendaftaran');
    $sheet->setCellValue('C1', 'Tgl Masuk Sekolah');
    $sheet->setCellValue('D1', 'NIS');
    $sheet->setCellValue('E1', 'No. Peserta Ujian');
    $sheet->setCellValue('F1', 'Pernah Paud');
    $sheet->setCellValue('G1', 'Pernah Tk');
    $sheet->setCellValue('H1', 'No.Skhun');
    $sheet->setCellValue('I1', 'No.Ijazah');
    $sheet->setCellValue('J1', 'Hobi');
    $sheet->setCellValue('K1', 'Cita-Cita');
    $sheet->setCellValue('L1', 'Nama Lengkap');
    $sheet->setCellValue('M1', 'Jk');
    $sheet->setCellValue('N1', 'NISN');
    $sheet->setCellValue('O1', 'NIK');
    $sheet->setCellValue('P1', 'Tempat Lahir');
    $sheet->setCellValue('Q1', 'Tanggal Lahir');
    $sheet->setCellValue('R1', 'Agama');
    $sheet->setCellValue('S1', 'Berkebutuhan Khusus');
    $sheet->setCellValue('T1', 'Alamat');
    $sheet->setCellValue('U1', 'RT');
    $sheet->setCellValue('V1', 'RW');
    $sheet->setCellValue('W1', 'Dusun');
    $sheet->setCellValue('X1', 'Kelurahan');
    $sheet->setCellValue('Y1', 'Kecamatan');
    $sheet->setCellValue('Z1', 'Kode Pos');
    $sheet->setCellValue('AA1', 'Tempat Tinggal');
    $sheet->setCellValue('AB1', 'Transportasi');
    $sheet->setCellValue('AC1', 'No.Hp');
    $sheet->setCellValue('AD1', 'No.Tlp');
    $sheet->setCellValue('AE1', 'Email');
    $sheet->setCellValue('AF1', 'Penerima Kps');
    $sheet->setCellValue('AG1', 'No.Kps');
    $sheet->setCellValue('AH1', 'Kewarganegaraan');
    $sheet->setCellValue('AI1', 'Negara');
    $sheet->setCellValue('AJ1', 'Nama Ayah Kandung');
    $sheet->setCellValue('AK1', 'Tahun Lahir');
    $sheet->setCellValue('AL1', 'Pendidikan');
    $sheet->setCellValue('AM1', 'Pekerjaan');
    $sheet->setCellValue('AN1', 'Penghasilan Bulanan');
    $sheet->setCellValue('AO1', 'Berkebutuhan Khusus');
    $sheet->setCellValue('AP1', 'Nama Ibu Kandung');
    $sheet->setCellValue('AQ1', 'Tahun Lahir');
    $sheet->setCellValue('AR1', 'Pendidikan');
    $sheet->setCellValue('AS1', 'Pekerjaan');
    $sheet->setCellValue('AT1', 'Penghasilan');
    $sheet->setCellValue('AU1', 'Berkebutuhan Khusus');
    
    $sql = mysqli_query($conn, "SELECT * FROM tbl_regis, tbl_data_pribadi, tbl_data_ayah, tbl_data_ibu");
    $i = 2;
    $no = 1;
    while ($row = mysqli_fetch_array($sql)) {
        $sheet->setCellValue('A'.$i, $no++);
        $sheet->setCellValue('B'.$i, $row['jenis_pendaftaran']);
        $sheet->setCellValue('C'.$i, $row['tanggal_masuk_sekolah']);
        $sheet->setCellValue('D'.$i, $row['nis']);
        $sheet->setCellValue('E'.$i, $row['no_peserta_ujian']);
        $sheet->setCellValue('F'.$i, $row['paud']);
        $sheet->setCellValue('G'.$i, $row['tk']);
        $sheet->setCellValue('H'.$i, $row['no_skhun']);
        $sheet->setCellValue('I'.$i, $row['no_ijazah']);
        $sheet->setCellValue('J'.$i, $row['hobi']);
        $sheet->setCellValue('K'.$i, $row['cita_cita']);
        $sheet->setCellValue('L'.$i, $row['nama_lengkap']);
        $sheet->setCellValue('M'.$i, $row['jenis_kelamin']);
        $sheet->setCellValue('N'.$i, $row['nisn']);
        $sheet->setCellValue('O'.$i, $row['nik']);
        $sheet->setCellValue('P'.$i, $row['tempat_lahir']);
        $sheet->setCellValue('Q'.$i, $row['tanggal_lahir']);
        $sheet->setCellValue('R'.$i, $row['agama']);
        $sheet->setCellValue('S'.$i, $row['berkebutuhan_khusus']);
        $sheet->setCellValue('T'.$i, $row['alamat']);
        $sheet->setCellValue('U'.$i, $row['rt']);
        $sheet->setCellValue('V'.$i, $row['rw']);
        $sheet->setCellValue('W'.$i, $row['dusun']);
        $sheet->setCellValue('X'.$i, $row['kelurahan']);
        $sheet->setCellValue('Y'.$i, $row['kecamatan']);
        $sheet->setCellValue('Z'.$i, $row['kode_pos']);
        $sheet->setCellValue('AA'.$i, $row['tempat_tinggal']);
        $sheet->setCellValue('AB'.$i, $row['transportasi']);
        $sheet->setCellValue('AC'.$i, $row['no_hp']);
        $sheet->setCellValue('AD'.$i, $row['no_tlp']);
        $sheet->setCellValue('AE'.$i, $row['email']);
        $sheet->setCellValue('AF'.$i, $row['penerima_kps']);
        $sheet->setCellValue('AG'.$i, $row['no_kps']);
        $sheet->setCellValue('AH'.$i, $row['kewarganegaraan']);
        $sheet->setCellValue('AI'.$i, $row['negara']);
        $sheet->setCellValue('AJ'.$i, $row['nama']);
        $sheet->setCellValue('AK'.$i, $row['tahun_lahir']);
        $sheet->setCellValue('AL'.$i, $row['pendidikan']);
        $sheet->setCellValue('AM'.$i, $row['pekerjaan']);
        $sheet->setCellValue('AN'.$i, $row['penghasilan_bulanan']);
        $sheet->setCellValue('AO'.$i, $row['berkebutuhan_khusus']);
        $sheet->setCellValue('AP'.$i, $row['nama']);
        $sheet->setCellValue('AQ'.$i, $row['tahun_lahir']);
        $sheet->setCellValue('AR'.$i, $row['pendidikan']);
        $sheet->setCellValue('AS'.$i, $row['pekerjaan']);
        $sheet->setCellValue('AT'.$i, $row['penghasilan_bulanan']);
        $sheet->setCellValue('AU'.$i, $row['berkebutuhan_khusus']);
        $i++;
    }
    $styleArray = [
        'borders'=>[
            'allBorders'=>[
                'borderStyle'=> \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];

    $i = $i - 1;
    $sheet->getStyle('A1:AU1'.$i)->applyFromArray($styleArray);
    $writer = new Xlsx($spreadsheet);
    $writer->save('Report Data Pendaftaran Siswa.xlsx');
?>