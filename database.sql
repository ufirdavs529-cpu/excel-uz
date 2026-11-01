-- ============================================================================
-- MICROSOFT EXCEL WEB DATABASE SCHEMA (MySQL)
-- Baza nomi: excel_db
-- ============================================================================

CREATE DATABASE IF NOT EXISTS `excel_db` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `excel_db`;

-- 1. Kitoblar (Workbooks) jadvali
CREATE TABLE IF NOT EXISTS `workbooks` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `name` VARCHAR(255) NOT NULL UNIQUE,
    `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 2. Yacheykalar (Cells) jadvali
CREATE TABLE IF NOT EXISTS `cells` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `workbook_name` VARCHAR(255) NOT NULL,
    `cell_id` VARCHAR(10) NOT NULL,
    `val` LONGTEXT,
    `formula` TEXT,
    `style_json` TEXT,
    `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    UNIQUE KEY `wb_cell` (`workbook_name`, `cell_id`),
    FOREIGN KEY (`workbook_name`) REFERENCES `workbooks`(`name`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- ============================================================================
-- NAMUNAVIY MA'LUMOTLAR (Sample Data for Recent Files)
-- ============================================================================

-- 1-kitob: 2-maktab_baholari
INSERT INTO `workbooks` (`name`) VALUES ('2-maktab_baholari')
ON DUPLICATE KEY UPDATE `updated_at` = CURRENT_TIMESTAMP;

INSERT INTO `cells` (`workbook_name`, `cell_id`, `val`, `formula`, `style_json`) VALUES
('2-maktab_baholari', 'A1', '2-Maktab O\'quvchilari Baholar Jurnali', '', '{"bold":true,"fontSize":"15px","color":"#107c41"}'),
('2-maktab_baholari', 'A3', '№', '', '{"bold":true,"bg":"#f3f2f1"}'),
('2-maktab_baholari', 'B3', 'O\'quvchi F.I.Sh', '', '{"bold":true,"bg":"#f3f2f1"}'),
('2-maktab_baholari', 'C3', 'Matematika', '', '{"bold":true,"bg":"#f3f2f1"}'),
('2-maktab_baholari', 'D3', 'Informatika', '', '{"bold":true,"bg":"#f3f2f1"}'),
('2-maktab_baholari', 'E3', 'Ingliz tili', '', '{"bold":true,"bg":"#f3f2f1"}'),
('2-maktab_baholari', 'F3', 'O\'rtacha baho', '', '{"bold":true,"bg":"#f3f2f1"}'),
('2-maktab_baholari', 'A4', '1', '', '{}'),
('2-maktab_baholari', 'B4', 'Aliyev Dovud', '', '{}'),
('2-maktab_baholari', 'C4', '5', '', '{}'),
('2-maktab_baholari', 'D4', '5', '', '{}'),
('2-maktab_baholari', 'E4', '5', '', '{}'),
('2-maktab_baholari', 'F4', '5', '=AVERAGE(C4:E4)', '{}'),
('2-maktab_baholari', 'A5', '2', '', '{}'),
('2-maktab_baholari', 'B5', 'Karimov Jasur', '', '{}'),
('2-maktab_baholari', 'C5', '4', '', '{}'),
('2-maktab_baholari', 'D5', '5', '', '{}'),
('2-maktab_baholari', 'E5', '4', '', '{}'),
('2-maktab_baholari', 'F5', '4.33', '=AVERAGE(C5:E5)', '{}'),
('2-maktab_baholari', 'A9', 'Sinf o\'rtacha ko\'rsatkichi:', '', '{"bold":true}'),
('2-maktab_baholari', 'F9', '4.67', '=AVERAGE(F4:F5)', '{"bold":true,"bg":"#d4edda","color":"#155724"}')
ON DUPLICATE KEY UPDATE `val` = VALUES(`val`), `formula` = VALUES(`formula`), `style_json` = VALUES(`style_json`);

-- 2-kitob: sirdaryo_11-sinf_lotin
INSERT INTO `workbooks` (`name`) VALUES ('sirdaryo_11-sinf_lotin')
ON DUPLICATE KEY UPDATE `updated_at` = CURRENT_TIMESTAMP;

INSERT INTO `cells` (`workbook_name`, `cell_id`, `val`, `formula`, `style_json`) VALUES
('sirdaryo_11-sinf_lotin', 'A1', 'Sirdaryo viloyati 11-sinf reyting jadvali', '', '{"bold":true,"fontSize":"15px","color":"#004e8c"}'),
('sirdaryo_11-sinf_lotin', 'A3', 'T/r', '', '{"bold":true,"bg":"#e1dfdd"}'),
('sirdaryo_11-sinf_lotin', 'B3', 'Tuman / Shahar', '', '{"bold":true,"bg":"#e1dfdd"}'),
('sirdaryo_11-sinf_lotin', 'C3', 'Bitiruvchilar', '', '{"bold":true,"bg":"#e1dfdd"}'),
('sirdaryo_11-sinf_lotin', 'D3', 'Oliy ta\'limga kirish', '', '{"bold":true,"bg":"#e1dfdd"}'),
('sirdaryo_11-sinf_lotin', 'E3', 'Ulush (%)', '', '{"bold":true,"bg":"#e1dfdd"}'),
('sirdaryo_11-sinf_lotin', 'A4', '1', '', '{}'),
('sirdaryo_11-sinf_lotin', 'B4', 'Guliston shahar', '', '{}'),
('sirdaryo_11-sinf_lotin', 'C4', '1250', '', '{}'),
('sirdaryo_11-sinf_lotin', 'D4', '980', '', '{}'),
('sirdaryo_11-sinf_lotin', 'E4', '78.4', '=(D4/C4)*100', '{}'),
('sirdaryo_11-sinf_lotin', 'A9', 'JAMI VILOYAT:', '', '{"bold":true}'),
('sirdaryo_11-sinf_lotin', 'C9', '1250', '=SUM(C4:C4)', '{"bold":true}'),
('sirdaryo_11-sinf_lotin', 'D9', '980', '=SUM(D4:D4)', '{"bold":true}')
ON DUPLICATE KEY UPDATE `val` = VALUES(`val`), `formula` = VALUES(`formula`), `style_json` = VALUES(`style_json`);
