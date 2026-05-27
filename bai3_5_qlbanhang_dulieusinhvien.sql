--bai 3.5: quản lí bán hàng
-- Tạo CSDL
-- Tạo CSDL
CREATE DATABASE QLBH_3_5;
GO
USE QLBH;
GO

-- Tạo bảng KHACHHG
CREATE TABLE KHACHHG (
    MAKH VARCHAR(10) PRIMARY KEY,
    TENKH NVARCHAR(100),
    DCHI NVARCHAR(200),
    DTHOAI VARCHAR(20)
);

-- Tạo bảng NHASX
CREATE TABLE NHASX (
    MANSX VARCHAR(10) PRIMARY KEY,
    TENNSX NVARCHAR(100),
    DCHI NVARCHAR(200),
    DTHOAI VARCHAR(20)
);

-- Tạo bảng NHACC
CREATE TABLE NHACC (
    MANCC VARCHAR(10) PRIMARY KEY,
    TENNCC NVARCHAR(100),
    DCHI NVARCHAR(200),
    DTHOAI VARCHAR(20)
);

-- Tạo bảng HANG
CREATE TABLE HANG (
    MAHG VARCHAR(10) PRIMARY KEY,
    TENHG NVARCHAR(100),
    DVT NVARCHAR(20),
    SOLUONGTON INT,
    MANSX VARCHAR(10),
    TINHTRANG NVARCHAR(50),
    FOREIGN KEY (MANSX) REFERENCES NHASX(MANSX)
);

-- Tạo bảng PHIEUNHAP
CREATE TABLE PHIEUNHAP (
    MAPN VARCHAR(10) PRIMARY KEY,
    NGAYNHAP DATE,
    MANCC VARCHAR(10),
    TIENNHAP MONEY,
    FOREIGN KEY (MANCC) REFERENCES NHACC(MANCC)
);

-- Tạo bảng CHITIETPN
CREATE TABLE CHITIETPN (
    MAPN VARCHAR(10),
    MAHG VARCHAR(10),
    SOLUONG INT,
    GIANHAP MONEY,
    THANHTIEN MONEY,
    PRIMARY KEY (MAPN, MAHG),
    FOREIGN KEY (MAPN) REFERENCES PHIEUNHAP(MAPN),
    FOREIGN KEY (MAHG) REFERENCES HANG(MAHG)
);

-- Tạo bảng HOADON
CREATE TABLE HOADON (
    MAHD VARCHAR(10) PRIMARY KEY,
    NGAYBAN DATE,
    MAKH VARCHAR(10),
    TIENBAN MONEY,
    GIAMGIA NVARCHAR(10),
    THANHTOAN MONEY,
    FOREIGN KEY (MAKH) REFERENCES KHACHHG(MAKH)
);

-- Tạo bảng CHITIETHD
CREATE TABLE CHITIETHD (
    MAHD VARCHAR(10),
    MAHG VARCHAR(10),
    SOLUONG INT,
    GIABAN MONEY,
    THANHTIEN MONEY,
    PRIMARY KEY (MAHD, MAHG),
    FOREIGN KEY (MAHD) REFERENCES HOADON(MAHD),
    FOREIGN KEY (MAHG) REFERENCES HANG(MAHG)
);

-- Tạo bảng DONGIA
CREATE TABLE DONGIA (
    MAHG VARCHAR(10) PRIMARY KEY,
    NGAYCN DATE,
    GIA MONEY,
    FOREIGN KEY (MAHG) REFERENCES HANG(MAHG)
);


-- Dữ liệu cho bảng KHACHHG
INSERT INTO KHACHHG (MAKH, TENKH, DCHI, DTHOAI) VALUES
('KH001', N'Nguyễn Văn An', N'123 Nguyễn Huệ, Quận 1', '0912345678'),
('KH002', N'Trần Thị Bình', N'456 Lê Lợi, Quận 2', '0912345679'),
('KH003', N'Lê Văn Cường', N'789 Hai Bà Trưng, Quận 3', '0912345680'),
('KH004', N'Phạm Thị Duyên', N'1011 Võ Văn Kiệt, Quận 4', '0912345681'),
('KH005', N'Hoàng Văn Hưng', N'1213 Lý Tự Trọng, Quận 1', '0912345682'),
('KH006', N'Đinh Thị Kiều', N'1415 Trần Hưng Đạo, Quận 5', '0912345683'),
('KH007', N'Vũ Văn Lâm', N'1617 CMT8, Quận 10', '0912345684'),
('KH008', N'Mai Thị Loan', N'1819 Hậu Giang, Quận 6', '0912345685'),
('KH009', N'Bùi Văn Minh', N'2021 Tôn Đức Thắng, Quận 1', '0912345686'),
('KH010', N'Ngô Thị Thảo', N'2223 Nguyễn Trãi, Quận 5', '0912345687');

-- Dữ liệu cho bảng NHASX
INSERT INTO NHASX (MANSX, TENNSX, DCHI, DTHOAI) VALUES
('NSX01', N'Công ty CP ABC', N'KCN Sóng Thần, Bình Dương', '0987654321'),
('NSX02', N'Công ty TNHH XYZ', N'KCN Tân Tạo, TP.HCM', '0987654322'),
('NSX03', N'Tập đoàn DEF', N'KCN Phố Nối, Hưng Yên', '0987654323'),
('NSX04', N'Nhà máy GHI', N'KCN Biên Hòa, Đồng Nai', '0987654324'),
('NSX05', N'Thương hiệu JKL', N'Lô C, KCN Hiệp Phước, TP.HCM', '0987654325'),
('NSX06', N'Đơn vị MNO', N'KCN Vĩnh Lộc, TP.HCM', '0987654326'),
('NSX07', N'Công ty PQR', N'KCN Amata, Đồng Nai', '0987654327'),
('NSX08', N'Công ty RST', N'KCN Vsip, Bình Dương', '0987654328'),
('NSX09', N'Công ty UVW', N'KCN Long Đức, Trà Vinh', '0987654329'),
('NSX10', N'Nhà máy XYZ', N'KCN Cát Lái, TP.HCM', '0987654330');

-- Dữ liệu cho bảng NHACC
INSERT INTO NHACC (MANCC, TENNCC, DCHI, DTHOAI) VALUES
('NCC01', N'Công ty Phân phối 1', N'111 Võ Văn Tần, Q.3', '0281111111'),
('NCC02', N'Công ty Thương mại 2', N'222 Nguyễn Đình Chiểu, Q.3', '0282222222'),
('NCC03', N'Đại lý Giao hàng nhanh', N'333 Điện Biên Phủ, Q.3', '0283333333'),
('NCC04', N'Công ty Xuất nhập khẩu', N'444 Cách Mạng Tháng 8, Q.3', '0284444444'),
('NCC05', N'Nhà cung cấp Lớn', N'555 Hai Bà Trưng, Q.3', '0285555555'),
('NCC06', N'Nhà cung cấp A', N'123 Đường A, Quận 1', '0901234567'),
('NCC07', N'Nhà cung cấp B', N'456 Đường B, Quận 2', '0902345678'),
('NCC08', N'Nhà cung cấp C', N'789 Đường C, Quận 3', '0903456789'),
('NCC09', N'Nhà cung cấp D', N'101 Đường D, Quận 4', '0904567890'),
('NCC10', N'Nhà cung cấp E', N'112 Đường E, Quận 5', '0905678901');

-- Dữ liệu cho bảng HANG
INSERT INTO HANG (MAHG, TENHG, DVT, SOLUONGTON, MANSX, TINHTRANG) VALUES
('HG001', N'Sữa tươi', N'Hộp', 150, 'NSX01', N'Còn hàng'),
('HG002', N'Bánh quy', N'Gói', 200, 'NSX02', N'Còn hàng'),
('HG003', N'Nước ngọt', N'Chai', 100, 'NSX03', N'Còn hàng'),
('HG004', N'Mì gói', N'Thùng', 50, 'NSX04', N'Còn hàng'),
('HG005', N'Kem đánh răng', N'Tuýp', 300, 'NSX05', N'Còn hàng'),
('HG006', N'Xà phòng', N'Bánh', 80, 'NSX06', N'Còn hàng'),
('HG007', N'Dầu ăn', N'Chai', 75, 'NSX07', N'Còn hàng'),
('HG008', N'Gạo', N'Kg', 120, 'NSX08', N'Còn hàng'),
('HG009', N'Bột giặt', N'Túi', 90, 'NSX09', N'Còn hàng'),
('HG010', N'Bia', N'Thùng', 60, 'NSX10', N'Còn hàng');

-- Dữ liệu cho bảng PHIEUNHAP
INSERT INTO PHIEUNHAP (MAPN, NGAYNHAP, MANCC, TIENNHAP) VALUES
('PN001', '2025-09-01', 'NCC01', 0),
('PN002', '2025-09-02', 'NCC02', 0),
('PN003', '2025-09-03', 'NCC03', 0),
('PN004', '2025-09-04', 'NCC04', 0),
('PN005', '2025-09-05', 'NCC05', 0),
('PN006', '2025-09-06', 'NCC06', 0),
('PN007', '2025-09-07', 'NCC07', 0),
('PN008', '2025-09-08', 'NCC08', 0),
('PN009', '2025-09-09', 'NCC09', 0),
('PN010', '2025-09-10', 'NCC10', 0);

-- Dữ liệu cho bảng CHITIETPN
INSERT INTO CHITIETPN (MAPN, MAHG, SOLUONG, GIANHAP, THANHTIEN) VALUES
('PN001', 'HG001', 50, 10000, 0),
('PN002', 'HG002', 40, 5000, 0),
('PN003', 'HG003', 60, 8000, 0),
('PN004', 'HG004', 30, 70000, 0),
('PN005', 'HG005', 100, 12000, 0),
('PN006', 'HG006', 50, 15000, 0),
('PN007', 'HG007', 25, 45000, 0),
('PN008', 'HG008', 50, 20000, 0),
('PN009', 'HG009', 40, 35000, 0),
('PN010', 'HG010', 20, 150000, 0);

-- Dữ liệu cho bảng HOADON
INSERT INTO HOADON (MAHD, NGAYBAN, MAKH, TIENBAN, GIAMGIA, THANHTOAN) VALUES
('HD001', '2025-09-11', 'KH001', 0, NULL, 0),
('HD002', '2025-09-12', 'KH002', 0, NULL, 0),
('HD003', '2025-09-13', 'KH003', 0, NULL, 0),
('HD004', '2025-09-14', 'KH004', 0, NULL, 0),
('HD005', '2025-09-15', 'KH005', 0, NULL, 0),
('HD006', '2025-09-16', 'KH006', 0, NULL, 0),
('HD007', '2025-09-17', 'KH007', 0, NULL, 0),
('HD008', '2025-09-18', 'KH008', 0, NULL, 0),
('HD009', '2025-09-19', 'KH009', 0, NULL, 0),
('HD010', '2025-09-20', 'KH010', 0, NULL, 0);

-- Dữ liệu cho bảng CHITIETHD
INSERT INTO CHITIETHD (MAHD, MAHG, SOLUONG, GIABAN, THANHTIEN) VALUES
('HD001', 'HG001', 5, 0, 0),
('HD002', 'HG002', 3, 0, 0),
('HD003', 'HG003', 10, 0, 0),
('HD004', 'HG004', 2, 0, 0),
('HD005', 'HG005', 8, 0, 0),
('HD006', 'HG006', 5, 0, 0),
('HD007', 'HG007', 2, 0, 0),
('HD008', 'HG008', 4, 0, 0),
('HD009', 'HG009', 3, 0, 0),
('HD010', 'HG010', 1, 0, 0);

-- Dữ liệu cho bảng DONGIA
INSERT INTO DONGIA (MAHG, NGAYCN, GIA) VALUES
('HG001', '2025-09-10', 15000),
('HG002', '2025-09-10', 8000),
('HG003', '2025-09-10', 10000),
('HG004', '2025-09-10', 100000),
('HG005', '2025-09-10', 18000),
('HG006', '2025-09-10', 20000),
('HG007', '2025-09-10', 60000),
('HG008', '2025-09-10', 25000),
('HG009', '2025-09-10', 45000),
('HG010', '2025-09-10', 200000);