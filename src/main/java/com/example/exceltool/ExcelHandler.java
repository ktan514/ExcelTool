package com.example.exceltool;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.security.GeneralSecurityException;

public class ExcelHandler {
    private final String filePath;
    private final Workbook workbook;

    public ExcelHandler(String filePath, String password) throws IOException {
        this.filePath = filePath;
        try {
            this.workbook = openWorkbook(filePath, password);
        } catch (GeneralSecurityException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public Workbook openWorkbook(String filePath, String password) throws IOException, GeneralSecurityException, InvalidFormatException {
        Workbook workbook;
        try (FileInputStream fis = new FileInputStream(filePath);
             BufferedInputStream bis = new BufferedInputStream(fis)) {
            if (DocumentFactoryHelper.hasOOXMLHeader(bis)) {
                // XSSF (.xlsx)形式のファイルを処理する
                bis.reset();
                workbook = openXSSFWorkbook(fis, password);
            } else {
                // HSSF (.xls)形式のファイルを処理する
                bis.reset();
                workbook = openHSSFWorkbook(fis, password);
            }
        }
        return workbook;
    }

    private Workbook openXSSFWorkbook(FileInputStream fis, String password) throws IOException, GeneralSecurityException, InvalidFormatException {
        Workbook workbook = null;
        POIFSFileSystem fs = new POIFSFileSystem(fis);
        EncryptionInfo info = new EncryptionInfo(fs);
        Decryptor d = Decryptor.getInstance(info);

        if (d.verifyPassword(password)) {
            try (InputStream dataStream = d.getDataStream(fs)) {
                OPCPackage opcPackage = OPCPackage.open(dataStream);
                workbook = new XSSFWorkbook(opcPackage);
            }
        } else {
            System.err.println("Incorrect password");
        }

        return workbook;
    }

    private Workbook openHSSFWorkbook(FileInputStream fis, String password) throws IOException, GeneralSecurityException {
        Workbook workbook = null;
        POIFSFileSystem fs = new POIFSFileSystem(fis);
        EncryptionInfo info = new EncryptionInfo(fs);
        Decryptor d = Decryptor.getInstance(info);

        if (d.verifyPassword(password)) {
            try (InputStream dataStream = d.getDataStream(fs)) {
                workbook = new HSSFWorkbook(dataStream);
            }
        } else {
            System.err.println("Incorrect password");
        }

        return workbook;
    }


    public void saveWorkbook() throws IOException {
        FileOutputStream file = new FileOutputStream(this.filePath);
        this.workbook.write(file);
        file.close();
    }

    public void closeWorkbook() throws IOException {
        this.workbook.close();
    }

    public Sheet getSheet(String sheetName) {
        return this.workbook.getSheet(sheetName);
    }

    public Cell getCell(Sheet sheet, int rowNumber, int columnNumber) {
        Row row = sheet.getRow(rowNumber);
        if (row == null) {
            row = sheet.createRow(rowNumber);
        }
        return row.getCell(columnNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
    }

    public void setCellValue(Cell cell, Object value) {
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value == null) {
            cell.setCellValue("");
        } else {
            cell.setCellValue(value.toString());
        }
    }
}
