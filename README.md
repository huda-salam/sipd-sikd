# sipd-sikd
Upload LRA, DTH, RTH hasil dari SIPD via Sinergi v4

 
I) SIMPAN DATA DARI EXCEL KE DB MYSQL
1. Instal driver ODBC mysql versi 8.0. Sesuaikan versinya (32bit/64bit) dengan versi Office di tempat Bapak/Ibu. Dalam hal sudah ada driver versi lain bisa disesuaikan pada modul vba Module1 pada function getCon
2. Buka file "import excel sipd ke mysql.xlsm", dan aktifkan macro dengan klik tombol "Enable macros" 
3. Buka file hasil download LRA / DTH-RTH pada aplikasi SIPD Aklap yang telah dibenahi datanya
4. Dengan kondisi worksheet LRA / DTH-RTH terbuka dan aktif, akses Developor - Macros.... pilih macro sesuai dokumen yang mau diimport, Run
