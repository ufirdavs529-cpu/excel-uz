<?php
/**
 * ============================================================================
 * MICROSOFT EXCEL WEB - PHP BACKEND API (MySQL Integration)
 * ============================================================================
 */

header('Content-Type: application/json; charset=UTF-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit();
}

// 1. Database Configuration
$db_host = 'localhost';
$db_name = 'excel_db';
$db_user = 'root';
$db_pass = '';

try {
    $pdo = new PDO("mysql:host={$db_host};charset=utf8mb4", $db_user, $db_pass, [
        PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC
    ]);

    // Auto-create database & tables if not exists
    $pdo->exec("CREATE DATABASE IF NOT EXISTS `{$db_name}` CHARACTER SET utf8mb4;");
    $pdo->exec("USE `{$db_name}`;");

    $pdo->exec("
        CREATE TABLE IF NOT EXISTS `workbooks` (
            `id` INT AUTO_INCREMENT PRIMARY KEY,
            `name` VARCHAR(255) NOT NULL UNIQUE,
            `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    ");

    $pdo->exec("
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
    ");

} catch (PDOException $e) {
    echo json_encode([
        'status' => 'error',
        'message' => 'MySQL ulanish xatosi: ' . $e->getMessage(),
        'hint' => 'XAMPP, Laragon yoki OpenServerda MySQL yoqilganligini tekshiring.'
    ]);
    exit();
}

// 2. Routing Actions
$action = isset($_GET['action']) ? $_GET['action'] : '';

switch ($action) {
    case 'save':
        handleSave($pdo);
        break;

    case 'load':
        handleLoad($pdo);
        break;

    case 'list':
        handleList($pdo);
        break;

    default:
        echo json_encode([
            'status' => 'success',
            'message' => 'Excel API ishlamoqda!',
            'available_actions' => ['list', 'load', 'save']
        ]);
        break;
}

/**
 * Save Workbook & Cells to MySQL
 */
function handleSave($pdo) {
    $rawInput = file_get_contents('php://input');
    $data = json_decode($rawInput, true);

    if (!$data || !isset($data['name']) || !isset($data['data'])) {
        echo json_encode(['status' => 'error', 'message' => 'Noto\'g\'ri ma\'lumot uzatildi']);
        return;
    }

    $name = trim($data['name']);
    $cells = $data['data'];

    try {
        $pdo->beginTransaction();

        // Insert or update workbook record
        $stmtWb = $pdo->prepare("INSERT INTO `workbooks` (`name`) VALUES (?) ON DUPLICATE KEY UPDATE `updated_at` = CURRENT_TIMESTAMP");
        $stmtWb->execute([$name]);

        // Insert / Update each cell
        $stmtCell = $pdo->prepare("
            INSERT INTO `cells` (`workbook_name`, `cell_id`, `val`, `formula`, `style_json`)
            VALUES (?, ?, ?, ?, ?)
            ON DUPLICATE KEY UPDATE
                `val` = VALUES(`val`),
                `formula` = VALUES(`formula`),
                `style_json` = VALUES(`style_json`)
        ");

        foreach ($cells as $cellId => $cellInfo) {
            $val = isset($cellInfo['val']) ? (string)$cellInfo['val'] : '';
            $formula = isset($cellInfo['formula']) ? (string)$cellInfo['formula'] : '';
            $style = isset($cellInfo['style']) ? json_encode($cellInfo['style']) : '{}';

            $stmtCell->execute([$name, $cellId, $val, $formula, $style]);
        }

        $pdo->commit();

        echo json_encode([
            'status' => 'success',
            'message' => "'{$name}' kitobi MySQL bazasiga muvaffaqiyatli saqlandi!",
            'cell_count' => count($cells)
        ]);
    } catch (Exception $e) {
        $pdo->rollBack();
        echo json_encode(['status' => 'error', 'message' => 'Saqlashda xatolik: ' . $e->getMessage()]);
    }
}

/**
 * Load Workbook & Cells from MySQL
 */
function handleLoad($pdo) {
    $name = isset($_GET['name']) ? trim($_GET['name']) : '';
    if (!$name) {
        echo json_encode(['status' => 'error', 'message' => 'Kitob nomi kiritilmadi']);
        return;
    }

    $stmt = $pdo->prepare("SELECT * FROM `cells` WHERE `workbook_name` = ?");
    $stmt->execute([$name]);
    $rows = $stmt->fetchAll();

    $cellsData = [];
    foreach ($rows as $row) {
        $cellsData[$row['cell_id']] = [
            'val' => $row['val'],
            'formula' => $row['formula'],
            'style' => json_decode($row['style_json'] ?: '{}', true)
        ];
    }

    echo json_encode([
        'status' => 'success',
        'name' => $name,
        'data' => $cellsData
    ]);
}

/**
 * List all saved workbooks
 */
function handleList($pdo) {
    $stmt = $pdo->query("SELECT `name`, `updated_at` FROM `workbooks` ORDER BY `updated_at` DESC");
    $list = $stmt->fetchAll();

    echo json_encode([
        'status' => 'success',
        'workbooks' => $list
    ]);
}
?>
