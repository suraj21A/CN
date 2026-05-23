-- database.sql

CREATE DATABASE IF NOT EXISTS Hundai_export_cn_db;
USE Hundai_export_cn_db;

CREATE TABLE IF NOT EXISTS cn_master (
  id INT AUTO_INCREMENT PRIMARY KEY,
  
  `CN No` VARCHAR(50) UNIQUE NOT NULL,
  `CN Type` VARCHAR(20),
  `CN Date` DATE,
  `Office List` VARCHAR(100),
  `Billing Office List` VARCHAR(100),
  `Consignor` VARCHAR(255),
  `Actual Route` VARCHAR(255),
  `Charged Route` VARCHAR(255),
  `Consignee` VARCHAR(255),
  `Vehicle No` VARCHAR(50),
  `Rate Chart` VARCHAR(100),
  `Load Type` VARCHAR(100),
  `Lr No` VARCHAR(100),
  `MM Invoice No` VARCHAR(50),
  `MM Invoice Date` DATE,
  `MM Material` VARCHAR(100),
  `MM Actual Weight` DECIMAL(10,2),
  `Rate` DECIMAL(10,2),
  `Freight` DECIMAL(10,2),
  `Other Charges` DECIMAL(10,2),
  `MM Chassis No` VARCHAR(100),
  `MM Engine No` VARCHAR(100),
  `MM NVRR No` VARCHAR(50),
  `MM Remark` VARCHAR(255),
  `No of Vehicles In Trailer` INT,
  `Pod Date` DATE,
  
  `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  INDEX idx_cn_no (`CN No`),
  INDEX idx_created_at (`created_at`)
);