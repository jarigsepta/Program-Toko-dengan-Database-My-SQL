-- phpMyAdmin SQL Dump
-- version 4.5.1
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: 20 Okt 2017 pada 12.48
-- Versi Server: 10.1.9-MariaDB
-- PHP Version: 5.5.30

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `ohssh`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `customer`
--

CREATE TABLE `customer` (
  `uid` varchar(10) NOT NULL,
  `email` varchar(20) NOT NULL,
  `name` varchar(30) NOT NULL,
  `pwd` varchar(100) NOT NULL,
  `phone` varchar(14) NOT NULL,
  `wa` varchar(14) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Truncate table before insert `customer`
--

TRUNCATE TABLE `customer`;
--
-- Dumping data untuk tabel `customer`
--

INSERT INTO `customer` (`uid`, `email`, `name`, `pwd`, `phone`, `wa`) VALUES
('OH2017001', 'ianz@gmail.com', 'Ian Acoustic', '202cb962ac59075b964b07152d234b70', '083832152700', 'X'),
('OH2017002', 'jarig@gmail.com', 'Jarig Septa', '202cb962ac59075b964b07152d234b70', '12345678', 'X'),
('OH2017003', 'ijul@gmail.com', 'Mas Izuel', '202cb962ac59075b964b07152d234b70', '0987654321', '0987654321'),
('OH2017004', 'setnov@gmail.com', 'setnov', '289dff07669d7a23de0ef88d2f7129e7', '89589589589', '89589589589');

-- --------------------------------------------------------

--
-- Struktur dari tabel `customer_pinfo`
--

CREATE TABLE `customer_pinfo` (
  `uid` varchar(10) NOT NULL,
  `address` varchar(50) NOT NULL,
  `gender` varchar(10) NOT NULL,
  `dob` date NOT NULL,
  `occupation` varchar(20) NOT NULL,
  `religion` varchar(10) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Truncate table before insert `customer_pinfo`
--

TRUNCATE TABLE `customer_pinfo`;
--
-- Dumping data untuk tabel `customer_pinfo`
--

INSERT INTO `customer_pinfo` (`uid`, `address`, `gender`, `dob`, `occupation`, `religion`) VALUES
('OH2017001', 'Jln Kabung AG, PP - G02', 'laki-laki', '1996-07-19', 'Musician', 'Islam'),
('OH2017002', 'Lamongan', 'laki-laki', '2017-09-01', 'Musisi', 'Islam'),
('OH2017003', 'Surabaya', 'laki-laki', '2017-09-06', 'Mahasiswa', 'Islam'),
('OH2017004', 'sbya', 'laki-laki', '2016-09-14', 'wiraswasta', 'Islam');

-- --------------------------------------------------------

--
-- Struktur dari tabel `customer_wallet`
--

CREATE TABLE `customer_wallet` (
  `uid` varchar(15) NOT NULL,
  `balance` int(10) NOT NULL,
  `payment_method` varchar(20) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Truncate table before insert `customer_wallet`
--

TRUNCATE TABLE `customer_wallet`;
--
-- Dumping data untuk tabel `customer_wallet`
--

INSERT INTO `customer_wallet` (`uid`, `balance`, `payment_method`) VALUES
('OH2017001', 23196000, 'Pulsa'),
('OH2017002', 126500, 'Indomaret'),
('OH2017003', 4885000, 'Indomaret'),
('OH2017004', 600000, 'None');

-- --------------------------------------------------------

--
-- Struktur dari tabel `ssh_item_active`
--

CREATE TABLE `ssh_item_active` (
  `item_id` varchar(15) NOT NULL,
  `uid` varchar(15) NOT NULL,
  `server_cuser` varchar(30) NOT NULL,
  `host` varchar(30) NOT NULL,
  `port` varchar(4) NOT NULL,
  `user_ssh` varchar(15) NOT NULL,
  `pass_ssh` varchar(15) NOT NULL,
  `active_date` date NOT NULL,
  `exp_date` date NOT NULL,
  `days_duration` int(1) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Truncate table before insert `ssh_item_active`
--

TRUNCATE TABLE `ssh_item_active`;
--
-- Dumping data untuk tabel `ssh_item_active`
--

INSERT INTO `ssh_item_active` (`item_id`, `uid`, `server_cuser`, `host`, `port`, `user_ssh`, `pass_ssh`, `active_date`, `exp_date`, `days_duration`) VALUES
('BRZ01', 'OH2017001', '[ianme] - Bluemix IBM', '200.9.221.198', '404', 'ianme', 'ian@p4ssW0Rd', '2017-10-05', '2017-10-07', 2),
('IND01', 'OH2017002', '[jarig] - ServerWala PVT LTD', '36.66.76.177', '143', 'jarig', 'jarigjunanto', '2017-10-05', '2017-10-07', 2),
('BRZ01', 'OH2017002', '[omjarig] - Bluemix IBM', '200.9.221.198', '404', 'omjarig', 'oomjrg', '2017-10-05', '2017-10-11', 6),
('USA02', 'OH2017002', '[jarig@123] - OVH DD Hosting ', '74.201.86.22', '443', 'jarig@123', 'OM@Jarig', '2017-10-05', '2017-10-06', 1),
('USA01', 'OH2017002', '[masterkoding] - MassiveGrid', '13.65.101.4', '143', 'masterkoding', 'master', '2017-10-05', '2017-10-06', 1),
('IND02', 'OH2017001', '[ianZXC] - Niagahoster', '190.202.41.138', '143', 'ianZXC', 'ZXC@net', '2017-10-05', '2017-10-07', 2),
('USA01', 'OH2017003', '[ijul] - MassiveGrid', '13.65.101.4', '143', 'ijul', 'ijoel', '2017-10-05', '2017-10-07', 2),
('JPD01', 'OH2017001', '[Prof] - Naze XG Japan', 'jp.fullssh.com', '22', 'Prof', 'PFes', '2017-10-11', '2017-10-13', 2),
('IND03', 'OH2017001', '[FSSH] - FastSSH ID', '10.152.224.94', '143', 'FSSH', 'IAN', '2017-10-12', '2017-10-13', 1),
('USA04', 'OH2017001', '[xKw@Wakak] - CyberSSH - Premi', '172.16.89.4', '443', 'xKw@Wakak', 'PROfesor', '2017-10-12', '2017-10-14', 2),
('IND01', 'OH2017001', '[Vanizhed] - ServerWala PVT LT', '36.66.76.177', '143', 'Vanizhed', '@abcdE', '2017-10-12', '2017-10-18', 6),
('JPD01', 'OH2017002', '[Omjrg] - Naze XG Japan', 'jp.fullssh.com', '22', 'Omjrg', 'OOMja@123', '2017-10-12', '2017-10-13', 1),
('IND02', 'OH2017001', '[yan] - Niagahoster', '190.202.41.138', '143', 'yan', 'iyyan', '2017-10-12', '2017-10-14', 2),
('SGD01', 'OH2017001', '[ianzME] - SG Digitalocean', '202.73.51.102', '443', 'ianzME', 'AAA', '2017-10-12', '2017-10-16', 4),
('JPD01', 'OH2017001', '[KN] - Naze XG Japan', 'jp.fullssh.com', '22', 'KN', 'XO', '2017-10-12', '2017-10-18', 6),
('BRZ01', 'OH2017001', '[m] - Bluemix IBM', '200.9.221.198', '404', 'm', 'l', '2017-10-20', '2017-10-22', 2);

-- --------------------------------------------------------

--
-- Struktur dari tabel `ssh_item_menu`
--

CREATE TABLE `ssh_item_menu` (
  `item_id` varchar(15) NOT NULL,
  `ssh_server` varchar(30) NOT NULL,
  `host` varchar(30) NOT NULL,
  `port` varchar(4) NOT NULL,
  `price` int(13) NOT NULL,
  `stock` int(3) NOT NULL,
  `country` varchar(12) NOT NULL,
  `description` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Truncate table before insert `ssh_item_menu`
--

TRUNCATE TABLE `ssh_item_menu`;
--
-- Dumping data untuk tabel `ssh_item_menu`
--

INSERT INTO `ssh_item_menu` (`item_id`, `ssh_server`, `host`, `port`, `price`, `stock`, `country`, `description`) VALUES
('BRZ01', 'Bluemix IBM', '200.9.221.198', '404', 35000, 1, 'Brazil', '* Proximity badge + Biometric\r\n* n+1 UPS Battery Backup Units\r\n* Multi-Level Access Control\r\n* 24/7 on-site security'),
('IND01', 'ServerWala PVT LTD', '36.66.76.177', '143', 10000, 103, 'Indonesia', '* 5TB Bandwidth\r\n* Intel® Xeon® E3-1230V3 \r\n* 8 GB DDR3 ECC \r\n* 1 TB HHD SATA'),
('IND02', 'Niagahoster', '190.202.41.138', '143', 15000, 0, 'Indonesia', '* Unlimited Bandwidth\r\n* Nano\r\n* RAM 1024 MB\r\n* 20GB Disk Space\r\n'),
('IND03', 'FastSSH ID', '10.152.224.94', '143', 5000, 398, 'Indonesia', '* Cepat\r\n* Cloud ssh server\r\n* Debian SSH Installation\r\n* Automation'),
('ITA01', 'Velia.NET', '95.110.189.185', '143', 45000, 74, 'ITA', '* Intel Xeon E3-1230 v3\r\n* RAM: 16 GB DDR3 ECC\r\n* High-Anonymous'),
('JPC01', 'Clara Online [JP Server]', '160.202.41.138', '22', 24000, 86, 'Japan', '* Unlimited Bandwidth\r\n* Unmetered Speed\r\n* HTTPS Protocol\r\n* Dedicated Server\r\n* Anti-DDoSing Protection'),
('JPD01', 'Naze XG Japan', 'jp.fullssh.com', '22', 30000, 127, 'Japan', '* Three Month subscription\r\n* Fast Cloud Data\r\n* Support DDOS, Torrent\r\n* Allow P2P'),
('SGD01', 'SG Digitalocean', '202.73.51.102', '443', 25000, 64, 'Singapore', '* Unlimited Bandwidth\n*Unlimited User(s)\n* 99% Uptime\n* Fast'),
('SWH01', 'Switzerland Hostpoints.CH', 'swz2.serverip.co', '22', 32500, 195, 'Switzerland', '* 98% Uptime\r\n* Reliable\r\n* Work With Bitvise\r\n* Cloud'),
('USA01', 'MassiveGrid', '13.65.101.4', '143', 57500, 19, 'USA', '* 40GB Bandwidth\r\n* 1MB/s Speed\r\n* XEON E7 CPU Cores\r\n* H/A SSD Disk\r\n* High Availability\r\n* Instant Activation'),
('USA02', 'OVH DD Hosting ', '74.201.86.22', '443', 56000, 33, 'USA', '* Intel  Xeon D-1520\r\n* 4c/8t - 2.4GHz /2.7GHz\r\n* 64GB DDR4 ECC 2133 MHz\r\n* SoftRaid  2x2TB \r\n* 250 Mbps  bandwidth\r\n* vRack: 10 Mbps'),
('USA03', 'CloudSSH.US', '5.189.161.89', '808', 100000, 84, 'Berlin', '* Fast Data Transfer \r\n* High Speed Servers\r\n* Exclusive Secure Shell\r\n* No DDOS'),
('USA04', 'CyberSSH - Premium', '172.16.89.4', '443', 20000, 99, 'USA', '* 1024 VRAM \r\n* SSD Up to 3TB\r\n* Intel Xeon R');

-- --------------------------------------------------------

--
-- Struktur dari tabel `trx`
--

CREATE TABLE `trx` (
  `tx_id` varchar(15) NOT NULL,
  `total_tx` int(11) NOT NULL,
  `tx_date` date NOT NULL,
  `tx_time` time NOT NULL,
  `uid` varchar(15) NOT NULL,
  `name` varchar(50) NOT NULL,
  `item_id` varchar(15) NOT NULL,
  `ssh_server` varchar(30) NOT NULL,
  `active_date` date NOT NULL,
  `exp_date` date NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Truncate table before insert `trx`
--

TRUNCATE TABLE `trx`;
--
-- Dumping data untuk tabel `trx`
--

INSERT INTO `trx` (`tx_id`, `total_tx`, `tx_date`, `tx_time`, `uid`, `name`, `item_id`, `ssh_server`, `active_date`, `exp_date`) VALUES
('TX201701', 70000, '2017-10-05', '03:14:06', 'OH2017001', 'Ian Acoustic\r', 'BRZ01', 'Bluemix IBM', '2017-10-05', '2017-10-07');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `customer`
--
ALTER TABLE `customer`
  ADD PRIMARY KEY (`uid`);

--
-- Indexes for table `customer_wallet`
--
ALTER TABLE `customer_wallet`
  ADD PRIMARY KEY (`uid`),
  ADD KEY `uid` (`uid`);

--
-- Indexes for table `ssh_item_active`
--
ALTER TABLE `ssh_item_active`
  ADD KEY `item_id` (`item_id`),
  ADD KEY `uid` (`uid`);

--
-- Indexes for table `ssh_item_menu`
--
ALTER TABLE `ssh_item_menu`
  ADD PRIMARY KEY (`item_id`);

--
-- Indexes for table `trx`
--
ALTER TABLE `trx`
  ADD PRIMARY KEY (`tx_id`),
  ADD KEY `uid` (`uid`),
  ADD KEY `item_id` (`item_id`);

--
-- Ketidakleluasaan untuk tabel pelimpahan (Dumped Tables)
--

--
-- Ketidakleluasaan untuk tabel `customer_wallet`
--
ALTER TABLE `customer_wallet`
  ADD CONSTRAINT `FK_UID_Wallet` FOREIGN KEY (`uid`) REFERENCES `customer` (`uid`);

--
-- Ketidakleluasaan untuk tabel `ssh_item_active`
--
ALTER TABLE `ssh_item_active`
  ADD CONSTRAINT `FK_ItemID_Active` FOREIGN KEY (`item_id`) REFERENCES `ssh_item_menu` (`item_id`) ON UPDATE CASCADE,
  ADD CONSTRAINT `Fk_UID_Active` FOREIGN KEY (`uid`) REFERENCES `customer` (`uid`);

--
-- Ketidakleluasaan untuk tabel `trx`
--
ALTER TABLE `trx`
  ADD CONSTRAINT `FK_UID_Customer` FOREIGN KEY (`uid`) REFERENCES `customer` (`uid`) ON UPDATE CASCADE;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
