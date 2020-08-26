/*
 Navicat Premium Data Transfer

 Source Server         : localhost_3108
 Source Server Type    : MySQL
 Source Server Version : 50562
 Source Host           : localhost:3108
 Source Schema         : penggajian

 Target Server Type    : MySQL
 Target Server Version : 50562
 File Encoding         : 65001

 Date: 27/08/2020 03:15:51
*/

SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS = 0;

-- ----------------------------
-- Table structure for biaya
-- ----------------------------
DROP TABLE IF EXISTS `biaya`;
CREATE TABLE `biaya`  (
  `kode` varchar(15) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL DEFAULT '',
  `nama` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  PRIMARY KEY (`kode`) USING BTREE
) ENGINE = MyISAM CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Dynamic;

-- ----------------------------
-- Records of biaya
-- ----------------------------
INSERT INTO `biaya` VALUES ('BADM', 'BIAYA ADMINISTRASI');
INSERT INTO `biaya` VALUES ('BAL', 'ALAT TULIS');
INSERT INTO `biaya` VALUES ('BANG', 'ANGSURAN KREDIT');
INSERT INTO `biaya` VALUES ('BBP', 'BUNGA PINJAM');
INSERT INTO `biaya` VALUES ('BDM', 'DISKON MEMBER');
INSERT INTO `biaya` VALUES ('BG', 'GAJI');
INSERT INTO `biaya` VALUES ('BI', 'IKLAN');
INSERT INTO `biaya` VALUES ('BIN', 'BIAYA INVENTARIS');
INSERT INTO `biaya` VALUES ('BK', 'BIAYA KOMISI');
INSERT INTO `biaya` VALUES ('BCC', 'BIAYA CONTROL CAB');
INSERT INTO `biaya` VALUES ('BL', 'BIAYA LISTRIK');
INSERT INTO `biaya` VALUES ('BO', 'BIAYA OPERASIONAL');
INSERT INTO `biaya` VALUES ('BPJ', 'BIAYA PAJAK');
INSERT INTO `biaya` VALUES ('BPD', 'PDAM');
INSERT INTO `biaya` VALUES ('BP', 'PAKET');
INSERT INTO `biaya` VALUES ('BPS', 'BIAYA PENGEMBALIAN');
INSERT INTO `biaya` VALUES ('BCTK', 'PERCETAKAN');
INSERT INTO `biaya` VALUES ('BS', 'BIAYA SEWA TOKO BENGKEL');
INSERT INTO `biaya` VALUES ('BAP', 'BIAYA AKMULASI PENYUSUTAN');
INSERT INTO `biaya` VALUES ('BIM', 'BIAYA ILMU MOTIVASI JIWA');
INSERT INTO `biaya` VALUES ('BR', 'BIAYA REKREASI');

-- ----------------------------
-- Table structure for departemen
-- ----------------------------
DROP TABLE IF EXISTS `departemen`;
CREATE TABLE `departemen`  (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `keterangan` varchar(20) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `autokode_karyawan` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE = InnoDB AUTO_INCREMENT = 8 CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of departemen
-- ----------------------------
INSERT INTO `departemen` VALUES (1, 'ALL DEPARTEMEN', '00001');
INSERT INTO `departemen` VALUES (2, 'WORK SHOP', '10001');
INSERT INTO `departemen` VALUES (3, 'HVAC', '20001');
INSERT INTO `departemen` VALUES (4, 'CIVIL', '30001');
INSERT INTO `departemen` VALUES (5, 'KIK', '40001');
INSERT INTO `departemen` VALUES (6, 'ALL AREA', '50001');
INSERT INTO `departemen` VALUES (7, 'DRIVER', '60001');

-- ----------------------------
-- Table structure for departemen_karyawan
-- ----------------------------
DROP TABLE IF EXISTS `departemen_karyawan`;
CREATE TABLE `departemen_karyawan`  (
  `id` varchar(15) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `departemen_id` int(11) NULL DEFAULT NULL
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of departemen_karyawan
-- ----------------------------
INSERT INTO `departemen_karyawan` VALUES ('10001', 2);
INSERT INTO `departemen_karyawan` VALUES ('10002', 2);
INSERT INTO `departemen_karyawan` VALUES ('10003', 2);
INSERT INTO `departemen_karyawan` VALUES ('10004', 2);
INSERT INTO `departemen_karyawan` VALUES ('10005', 2);
INSERT INTO `departemen_karyawan` VALUES ('10006', 2);
INSERT INTO `departemen_karyawan` VALUES ('10007', 2);
INSERT INTO `departemen_karyawan` VALUES ('10008', 2);
INSERT INTO `departemen_karyawan` VALUES ('10009', 2);
INSERT INTO `departemen_karyawan` VALUES ('10010', 2);
INSERT INTO `departemen_karyawan` VALUES ('20001', 3);
INSERT INTO `departemen_karyawan` VALUES ('20002', 3);
INSERT INTO `departemen_karyawan` VALUES ('20003', 3);
INSERT INTO `departemen_karyawan` VALUES ('20004', 3);
INSERT INTO `departemen_karyawan` VALUES ('20005', 3);
INSERT INTO `departemen_karyawan` VALUES ('20006', 3);
INSERT INTO `departemen_karyawan` VALUES ('20007', 3);
INSERT INTO `departemen_karyawan` VALUES ('20008', 3);
INSERT INTO `departemen_karyawan` VALUES ('30001', 4);
INSERT INTO `departemen_karyawan` VALUES ('30002', 4);
INSERT INTO `departemen_karyawan` VALUES ('30003', 4);
INSERT INTO `departemen_karyawan` VALUES ('30004', 4);
INSERT INTO `departemen_karyawan` VALUES ('30005', 4);
INSERT INTO `departemen_karyawan` VALUES ('30006', 4);
INSERT INTO `departemen_karyawan` VALUES ('30007', 4);
INSERT INTO `departemen_karyawan` VALUES ('30008', 4);
INSERT INTO `departemen_karyawan` VALUES ('30009', 4);
INSERT INTO `departemen_karyawan` VALUES ('30010', 5);
INSERT INTO `departemen_karyawan` VALUES ('40001', 5);
INSERT INTO `departemen_karyawan` VALUES ('40002', 5);
INSERT INTO `departemen_karyawan` VALUES ('40003', 5);
INSERT INTO `departemen_karyawan` VALUES ('40004', 5);
INSERT INTO `departemen_karyawan` VALUES ('40005', 5);
INSERT INTO `departemen_karyawan` VALUES ('40006', 5);
INSERT INTO `departemen_karyawan` VALUES ('40007', 5);
INSERT INTO `departemen_karyawan` VALUES ('40008', 5);
INSERT INTO `departemen_karyawan` VALUES ('40009', 5);
INSERT INTO `departemen_karyawan` VALUES ('40010', 5);
INSERT INTO `departemen_karyawan` VALUES ('40011', 5);
INSERT INTO `departemen_karyawan` VALUES ('40012', 5);
INSERT INTO `departemen_karyawan` VALUES ('40013', 5);
INSERT INTO `departemen_karyawan` VALUES ('50001', 6);
INSERT INTO `departemen_karyawan` VALUES ('50002', 6);
INSERT INTO `departemen_karyawan` VALUES ('50003', 6);
INSERT INTO `departemen_karyawan` VALUES ('50004', 6);
INSERT INTO `departemen_karyawan` VALUES ('60001', 7);

-- ----------------------------
-- Table structure for inventaris
-- ----------------------------
DROP TABLE IF EXISTS `inventaris`;
CREATE TABLE `inventaris`  (
  `id` varchar(15) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `tanggal` date NULL DEFAULT NULL,
  `nama` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `qty` int(11) NULL DEFAULT NULL,
  `perolehan` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `kategori` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `operator` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of inventaris
-- ----------------------------
INSERT INTO `inventaris` VALUES ('INT000001', '2017-01-01', 'MEJA', 10, 'BELI BARU', 'ALAT KANTOR', 'Administrator');
INSERT INTO `inventaris` VALUES ('INT000002', '2017-02-14', 'KURSI', 10, 'BELI BARU', 'ALAT KANTOR', 'Administrator');
INSERT INTO `inventaris` VALUES ('INT000003', '2017-05-01', 'SEPATU SAFETY', 10, 'BELI BARU', 'ALAT LAPANGAN', 'Administrator');
INSERT INTO `inventaris` VALUES ('INT000004', '2017-08-01', 'SARUNG TANGAN', 10, 'BELI BARU', 'ALAT LAPANGAN', 'Administrator');
INSERT INTO `inventaris` VALUES ('INT000005', '2017-09-01', 'MODEM SPEEDY', 10, 'BELI BARU', 'ALAT KANTOR', 'Administrator');

-- ----------------------------
-- Table structure for jenis_inventaris
-- ----------------------------
DROP TABLE IF EXISTS `jenis_inventaris`;
CREATE TABLE `jenis_inventaris`  (
  `id` int(5) NOT NULL,
  `nama` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of jenis_inventaris
-- ----------------------------
INSERT INTO `jenis_inventaris` VALUES (1, 'ALAT TULIS');
INSERT INTO `jenis_inventaris` VALUES (2, 'ALAT KANTOR');
INSERT INTO `jenis_inventaris` VALUES (3, 'ALAT LAPANGAN');
INSERT INTO `jenis_inventaris` VALUES (4, 'LAIN-LAIN');

-- ----------------------------
-- Table structure for jenis_potongan
-- ----------------------------
DROP TABLE IF EXISTS `jenis_potongan`;
CREATE TABLE `jenis_potongan`  (
  `id` int(5) NOT NULL,
  `nama` varchar(200) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of jenis_potongan
-- ----------------------------
INSERT INTO `jenis_potongan` VALUES (1, 'SOSIAL');
INSERT INTO `jenis_potongan` VALUES (2, 'PINJAMAN');
INSERT INTO `jenis_potongan` VALUES (3, 'TAB. SUKARELA');
INSERT INTO `jenis_potongan` VALUES (4, 'IURAN WAJIB');
INSERT INTO `jenis_potongan` VALUES (5, 'POTONGAN');
INSERT INTO `jenis_potongan` VALUES (6, 'LAIN-LAIN');

-- ----------------------------
-- Table structure for jenis_tunjangan
-- ----------------------------
DROP TABLE IF EXISTS `jenis_tunjangan`;
CREATE TABLE `jenis_tunjangan`  (
  `id` int(11) NULL DEFAULT NULL,
  `nama` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of jenis_tunjangan
-- ----------------------------
INSERT INTO `jenis_tunjangan` VALUES (1, 'HARIAN');
INSERT INTO `jenis_tunjangan` VALUES (2, 'BULANAN');
INSERT INTO `jenis_tunjangan` VALUES (3, 'TAHUNAN');
INSERT INTO `jenis_tunjangan` VALUES (4, 'ANAK');
INSERT INTO `jenis_tunjangan` VALUES (5, 'ISTRI');
INSERT INTO `jenis_tunjangan` VALUES (6, 'JABATAN');
INSERT INTO `jenis_tunjangan` VALUES (7, 'LAIN-LAIN');
INSERT INTO `jenis_tunjangan` VALUES (8, 'BPJS+');
INSERT INTO `jenis_tunjangan` VALUES (9, 'OVER TIME');
INSERT INTO `jenis_tunjangan` VALUES (10, 'Bonus HK FULL');

-- ----------------------------
-- Table structure for karyawan
-- ----------------------------
DROP TABLE IF EXISTS `karyawan`;
CREATE TABLE `karyawan`  (
  `id` varchar(15) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `nama` varchar(100) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `alamat` varchar(150) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `kontak` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `status` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `basic_hk` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `basic_gaji` decimal(15, 0) NULL DEFAULT NULL
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of karyawan
-- ----------------------------
INSERT INTO `karyawan` VALUES ('10001', 'ANDRIYAN', 'NULL', '-', '-', '23', 2356000);
INSERT INTO `karyawan` VALUES ('10002', 'SYAFRIL', 'NULL', '-', '-', '23', 2356000);
INSERT INTO `karyawan` VALUES ('10003', 'ALDI YOHANNAS', 'NULL', '-', '-', '23', 2356000);
INSERT INTO `karyawan` VALUES ('10004', 'INDRI YUDI', 'NULL', '-', '-', '23', 2356000);
INSERT INTO `karyawan` VALUES ('10005', 'ADI MUKHLIS', 'NULL', '-', '-', '23', 2356000);
INSERT INTO `karyawan` VALUES ('10006', 'NURUL ARIFIN', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('10007', 'MUSLIADI', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('10008', 'PADRI YANSA', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('10009', 'RUDI SAPUTRA', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('10010', 'TAUFIK HIDAYAT', 'NULL', '-', '-', '23', 0);
INSERT INTO `karyawan` VALUES ('20001', 'ADRIANTO', 'NULL', '-', '-', '23', 2356000);
INSERT INTO `karyawan` VALUES ('20002', 'FIRMAN ANDI PUTRA', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('20003', 'KHAIRAT', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('20004', 'ARDI SUPRIADI', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('20005', 'SYAHRUDIN AMIN', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('20006', 'BUDIARTO', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('20007', 'ILHAM MAULANA ISHAK', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('20008', 'ALEX MANSAH', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('30001', 'SYAMSIDAR', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('30002', 'JEFRY', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('30003', 'KIDUT', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('30004', 'NAZARUDIN', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('30005', 'RIYAN DAYAT', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('30006', 'ALI', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('30007', 'ZUBAND', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('30008', 'SOPIAN', 'NULL', '-', '-', '23', 1955000);
INSERT INTO `karyawan` VALUES ('30009', 'RITO', 'NULL', '-', '-', '23', 2500000);
INSERT INTO `karyawan` VALUES ('30010', 'M. SYAHPUTRA', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('40001', 'ERWAN', 'NULL', '-', '-', '23', 2175000);
INSERT INTO `karyawan` VALUES ('40002', 'KHODRAD', 'NULL', '-', '-', '23', 2075000);
INSERT INTO `karyawan` VALUES ('40003', 'HARMADI', 'NULL', '-', '-', '23', 2075000);
INSERT INTO `karyawan` VALUES ('40004', 'CHANDRA ARIF', 'NULL', '-', '-', '23', 2075000);
INSERT INTO `karyawan` VALUES ('40005', 'SARI RAMADAN', 'NULL', '-', '-', '23', 2075000);
INSERT INTO `karyawan` VALUES ('40006', 'RIKI BASTIAN', 'NULL', '-', '-', '23', 2075000);
INSERT INTO `karyawan` VALUES ('40007', 'SAIPUL ANWAR', 'NULL', '-', '-', '23', 2075000);
INSERT INTO `karyawan` VALUES ('40008', 'RIKO', 'NULL', '-', '-', '23', 2075000);
INSERT INTO `karyawan` VALUES ('40009', 'RIKA', 'NULL', '-', '-', '23', 1800000);
INSERT INTO `karyawan` VALUES ('40010', 'IGA KUMALASARI', 'NULL', '-', '-', '23', 1800000);
INSERT INTO `karyawan` VALUES ('40011', 'RIRIN ARISKA', 'NULL', '-', '-', '23', 1800000);
INSERT INTO `karyawan` VALUES ('40012', 'FEBRY ANDRIYANI', 'NULL', '-', '-', '23', 1400000);
INSERT INTO `karyawan` VALUES ('40013', 'MARWAN', 'NULL', '-', '-', '23', 2350000);
INSERT INTO `karyawan` VALUES ('50001', 'HABIBI', 'NULL', '-', '-', '7,5', 675000);
INSERT INTO `karyawan` VALUES ('50002', 'ONA', 'NULL', '-', '-', '8,5', 765000);
INSERT INTO `karyawan` VALUES ('50003', 'EDO', 'NULL', '-', '-', '8', 720000);
INSERT INTO `karyawan` VALUES ('50004', 'UCOK', 'NULL', '-', '-', '0', 0);
INSERT INTO `karyawan` VALUES ('60001', 'TOPIK', 'NULL', '-', '-', '23', 2500000);

-- ----------------------------
-- Table structure for karyawan_gaji
-- ----------------------------
DROP TABLE IF EXISTS `karyawan_gaji`;
CREATE TABLE `karyawan_gaji`  (
  `id` varchar(15) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL,
  `Kode` varchar(10) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL,
  `nama` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `basic_hk` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `basic_gaji` decimal(16, 3) NULL DEFAULT NULL,
  `hk_miss` int(5) NULL DEFAULT NULL,
  `hk_potongan` decimal(16, 0) NULL DEFAULT NULL,
  `hk_totpotong` decimal(16, 0) NULL DEFAULT NULL,
  `tunjangan` decimal(16, 0) NULL DEFAULT NULL,
  `potongan` decimal(16, 0) NULL DEFAULT NULL,
  `total_gaji` decimal(16, 0) NULL DEFAULT NULL,
  `tanggal` date NULL DEFAULT NULL,
  `gaji_bulan` varchar(10) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `operator` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  PRIMARY KEY (`Kode`, `id`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of karyawan_gaji
-- ----------------------------
INSERT INTO `karyawan_gaji` VALUES ('GK1709300001', '10001', 'ANDRIYAN', '23', 2356000.000, 0, 102434, 0, 100000, 100000, 2356000, '2017-09-30', '09', 'Administrator');

-- ----------------------------
-- Table structure for kas
-- ----------------------------
DROP TABLE IF EXISTS `kas`;
CREATE TABLE `kas`  (
  `kode` varchar(15) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL DEFAULT '',
  `nama` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `saldo` decimal(15, 0) NULL DEFAULT 0,
  PRIMARY KEY (`kode`) USING BTREE
) ENGINE = MyISAM CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Dynamic;

-- ----------------------------
-- Records of kas
-- ----------------------------
INSERT INTO `kas` VALUES ('KP', 'KAS PERUSAHAAN', 10400000);
INSERT INTO `kas` VALUES ('KG', 'KAS LAIN-LAIN', 1000000);
INSERT INTO `kas` VALUES ('KD', 'KAS DASAR', 1800000);
INSERT INTO `kas` VALUES ('KT', 'KAS TUNJANGAN', 1000000);

-- ----------------------------
-- Table structure for pemasukan
-- ----------------------------
DROP TABLE IF EXISTS `pemasukan`;
CREATE TABLE `pemasukan`  (
  `kode` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL DEFAULT '',
  `tanggal` date NULL DEFAULT NULL,
  `keterangan` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `kode_kas` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `jumlah` decimal(15, 2) NULL DEFAULT 0.00,
  `operator` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  PRIMARY KEY (`kode`) USING BTREE
) ENGINE = MyISAM CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Dynamic;

-- ----------------------------
-- Records of pemasukan
-- ----------------------------
INSERT INTO `pemasukan` VALUES ('PM14091701', '2017-07-01', 'UANG JATUH DARI LANGIT', 'KP', 1000000.00, 'Administrator');
INSERT INTO `pemasukan` VALUES ('PM14091702', '2017-08-01', 'UANG JATUH DARI LANGIT', 'KP', 1000000.00, 'Administrator');
INSERT INTO `pemasukan` VALUES ('PM22091701', '2017-09-22', 'UANG JATUH DARI LANGIT', 'KP', 1000000.00, 'Administrator');
INSERT INTO `pemasukan` VALUES ('PM22091702', '2017-09-22', 'Pembayaran Hutang Pinjaman Karyawan', 'KP', 800000.00, 'Administrator');

-- ----------------------------
-- Table structure for pengaturan
-- ----------------------------
DROP TABLE IF EXISTS `pengaturan`;
CREATE TABLE `pengaturan`  (
  `Kode` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL DEFAULT '',
  `Nama` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `Alamat` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `Struk` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `Link` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `Gambar` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `KodeReg` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `Kontak` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  PRIMARY KEY (`Kode`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of pengaturan
-- ----------------------------
INSERT INTO `pengaturan` VALUES ('1', 'SIMPEG LITE', 'JL.Lintas Timur no.288 - Pangkalan Kerinci - Riau - 28300', 'JL.Lintas Timur no.288, Pangkalan Kerinc', '-', '-', '-', '085222644014');

-- ----------------------------
-- Table structure for pengeluaran
-- ----------------------------
DROP TABLE IF EXISTS `pengeluaran`;
CREATE TABLE `pengeluaran`  (
  `id` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL DEFAULT '',
  `tanggal` date NULL DEFAULT NULL,
  `kd_biaya` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `keterangan` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `kode_kas` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `jumlah` decimal(15, 2) NULL DEFAULT 0.00,
  `operator` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE = MyISAM CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Dynamic;

-- ----------------------------
-- Records of pengeluaran
-- ----------------------------
INSERT INTO `pengeluaran` VALUES ('PG13091701', '2017-08-13', 'BADM', 'TEST BADM', 'KP', 100000.00, 'Administrator');
INSERT INTO `pengeluaran` VALUES ('PG14091701', '2017-08-14', 'BAL', 'PENA', 'KP', 100000.00, 'Administrator');
INSERT INTO `pengeluaran` VALUES ('PG20091701', '2017-08-14', 'BADM', 'ADM', 'KP', 100000.00, 'Administrator');
INSERT INTO `pengeluaran` VALUES ('PG20091702', '2017-08-20', 'BAL', 'PENA', 'KP', 1000000.00, 'Administrator');
INSERT INTO `pengeluaran` VALUES ('PG20091703', '2017-09-20', 'BAL', 'APA SAJA', 'KP', 1000000.00, 'Administrator');
INSERT INTO `pengeluaran` VALUES ('PG20091704', '2017-09-20', 'BAL', 'ADA', 'KP', 1000000.00, 'Administrator');
INSERT INTO `pengeluaran` VALUES ('PG20091705', '2017-09-20', 'BAL', 'PENA', 'KD', 400000.00, 'Administrator');
INSERT INTO `pengeluaran` VALUES ('PG20091706', '2017-09-20', 'BAL', 'AS', 'KP', 5000000.00, 'Administrator');
INSERT INTO `pengeluaran` VALUES ('PG20091707', '2017-10-03', 'BAL', 'PENA', 'KP', 2000000.00, 'Administrator');

-- ----------------------------
-- Table structure for pos_akses
-- ----------------------------
DROP TABLE IF EXISTS `pos_akses`;
CREATE TABLE `pos_akses`  (
  `id` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `uname` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `pword` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `namauser` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `jabatan` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `kode` varchar(35) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `registered` date NULL DEFAULT NULL,
  `gender` varchar(10) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of pos_akses
-- ----------------------------
INSERT INTO `pos_akses` VALUES ('00000', 'carlesneo70', 'uZqySm@k@p\\xi_Y;SySMtKEW', 'DeveloperSIP', 'Programmer', 'SU', '2017-09-22', 'LAKI-LAKI');
INSERT INTO `pos_akses` VALUES ('00001', 'admin', 'JW?jWkAkFo`xj;', 'Administrator', 'Kepala Utama', 'ADM', '2017-09-22', 'LAKI-LAKI');
INSERT INTO `pos_akses` VALUES ('00002', 'user', ';WqiXvAkFpw', 'Krani', 'Krani', 'SKT', '2017-09-22', 'PEREMPUAN');
INSERT INTO `pos_akses` VALUES ('00003', 'test', '', 'TEST1', 'TEST', 'SKT', '2017-10-05', 'PEREMPUAN');

-- ----------------------------
-- Table structure for pos_level
-- ----------------------------
DROP TABLE IF EXISTS `pos_level`;
CREATE TABLE `pos_level`  (
  `id` varchar(3) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL,
  `kode` varchar(25) CHARACTER SET latin1 COLLATE latin1_swedish_ci NOT NULL,
  `nama` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `karyawan_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `karyawan_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `karyawan_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `karyawan_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `potongan_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `potongan_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `potongan_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `potongan_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `tunjangan_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `tunjangan_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `tunjangan_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `tunjangan_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `gaji_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `gaji_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `gaji_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `gaji_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `gaji_generate` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `pengeluaran_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `pengeluaran_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `pengeluaran_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `pengeluaran_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `pemasukan_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `pemasukan_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `pemasukan_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `pemasukan_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `kas_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `kas_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `kas_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `kas_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `biaya_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `biaya_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `biaya_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `biaya_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `inventaris_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `inventaris_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `inventaris_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `inventaris_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `laporan_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `user_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `user_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `user_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `user_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `user_change` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `level_view` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `level_create` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `level_update` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `level_delete` varchar(5) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  PRIMARY KEY (`kode`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of pos_level
-- ----------------------------
INSERT INTO `pos_level` VALUES ('001', 'ADM', 'ADMIN', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True');
INSERT INTO `pos_level` VALUES ('002', 'SKT', 'SEKRETARIS', 'True', 'True', 'True', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'True', 'True', 'False', 'False', 'False', 'True', 'True', 'False', 'False', 'True', 'True', 'False', 'False', 'True', 'False', 'False', 'False', 'True', 'True', 'False', 'False', 'True', 'True', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'True', 'False', 'False', 'True', 'False');
INSERT INTO `pos_level` VALUES ('000', 'SU', 'SUPER USER', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True', 'True');
INSERT INTO `pos_level` VALUES ('003', 'USR', 'USER', 'True', 'True', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'True', 'True', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'True', 'True', 'False', 'False', 'False', 'False', 'False', 'False', 'False', 'True', 'False', 'False', 'False', 'False');

-- ----------------------------
-- Table structure for potongan
-- ----------------------------
DROP TABLE IF EXISTS `potongan`;
CREATE TABLE `potongan`  (
  `id` int(5) NOT NULL,
  `tanggal` date NULL DEFAULT NULL,
  `keterangan` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `jumlah` decimal(50, 0) NULL DEFAULT NULL,
  `kode` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of potongan
-- ----------------------------
INSERT INTO `potongan` VALUES (10001, '2017-09-14', 'PINJAMAN', 100000, '00000001');
INSERT INTO `potongan` VALUES (10002, '2017-07-11', 'PINJAMAN', 2000000, '00000006');
INSERT INTO `potongan` VALUES (10001, '2017-01-01', 'PINJAMAN', 500000, '00000002');
INSERT INTO `potongan` VALUES (10001, '2017-02-01', 'SOSIAL', 500000, '00000004');
INSERT INTO `potongan` VALUES (10002, '2017-03-01', 'PINJAMAN', 2000000, '00000003');
INSERT INTO `potongan` VALUES (10005, '2017-04-01', 'PINJAMAN', 500000, '00000007');
INSERT INTO `potongan` VALUES (10005, '2017-05-01', 'LAIN-LAIN', 100000, '00000008');
INSERT INTO `potongan` VALUES (10005, '2017-06-01', 'LAIN-LAIN', 100000, '00000009');
INSERT INTO `potongan` VALUES (10005, '2017-07-01', 'LAIN-LAIN', 100000, '00000006');
INSERT INTO `potongan` VALUES (10003, '2017-08-01', 'LAIN-LAIN', 100000, '00000010');
INSERT INTO `potongan` VALUES (10004, '2017-09-01', 'LAIN-LAIN', 1000000, '00000011');
INSERT INTO `potongan` VALUES (10001, '2017-10-01', 'PINJAMAN', 200000, '00000012');
INSERT INTO `potongan` VALUES (10004, '2017-09-15', 'IURAN WAJIB', 100000, '00000013');
INSERT INTO `potongan` VALUES (10002, '2017-10-02', 'POTONGAN', 100000, '00000014');
INSERT INTO `potongan` VALUES (10003, '2017-10-02', 'IURAN WAJIB', 10000, '00000015');
INSERT INTO `potongan` VALUES (10005, '2017-10-02', 'POTONGAN', 100000, '00000016');
INSERT INTO `potongan` VALUES (10004, '2017-10-02', 'POTONGAN', 100000, '00000017');
INSERT INTO `potongan` VALUES (10010, '2017-10-02', 'LAIN-LAIN', 200000, '00000018');
INSERT INTO `potongan` VALUES (10009, '2017-10-02', 'LAIN-LAIN', 200000, '00000019');
INSERT INTO `potongan` VALUES (10008, '2017-10-02', 'PINJAMAN', 200000, '00000020');
INSERT INTO `potongan` VALUES (10007, '2017-10-02', 'POTONGAN', 200000, '00000021');
INSERT INTO `potongan` VALUES (10006, '2017-10-02', 'SOSIAL', 200000, '00000022');

-- ----------------------------
-- Table structure for tunjangan
-- ----------------------------
DROP TABLE IF EXISTS `tunjangan`;
CREATE TABLE `tunjangan`  (
  `id` int(5) NOT NULL,
  `tanggal` date NULL DEFAULT NULL,
  `keterangan` varchar(255) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL,
  `jumlah` decimal(50, 0) NULL DEFAULT NULL,
  `kode` varchar(50) CHARACTER SET latin1 COLLATE latin1_swedish_ci NULL DEFAULT NULL
) ENGINE = InnoDB CHARACTER SET = latin1 COLLATE = latin1_swedish_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of tunjangan
-- ----------------------------
INSERT INTO `tunjangan` VALUES (10003, '2017-04-01', 'JABATAN', 100000, '00000005');
INSERT INTO `tunjangan` VALUES (10002, '2017-05-01', 'JABATAN', 200000, '00000002');
INSERT INTO `tunjangan` VALUES (10003, '2017-06-01', 'JABATAN', 100000, '00000003');
INSERT INTO `tunjangan` VALUES (10004, '2017-07-01', 'JABATAN', 500000, '00000004');
INSERT INTO `tunjangan` VALUES (10003, '2017-08-01', 'BULANAN', 100000, '00000001');
INSERT INTO `tunjangan` VALUES (10001, '2017-08-14', 'ANAK', 100000, '00000006');
INSERT INTO `tunjangan` VALUES (10001, '2017-09-30', 'BPJS+', 100000, '00000007');

SET FOREIGN_KEY_CHECKS = 1;
