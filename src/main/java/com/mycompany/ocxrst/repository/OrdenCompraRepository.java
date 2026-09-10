package com.mycompany.ocxrst.repository;

import com.mycompany.ocxrst.config.AppConfig;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OrdenCompraRepository {

    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    public File obtenerArchivo(String nombreBase) {
        return AppConfig.resolveBaseFile(nombreBase);
    }

    public Workbook abrirWorkbook(File archivo) throws IOException {
        try (FileInputStream fis = new FileInputStream(archivo)) {
            return new XSSFWorkbook(fis);
        }
    }

    public Workbook abrirOCrearWorkbook(File archivo, String nombreHoja) throws IOException {
        if (archivo.exists()) {
            return abrirWorkbook(archivo);
        }
        Workbook workbook = new XSSFWorkbook();
        workbook.createSheet(nombreHoja);
        return workbook;
    }

    public void guardarWorkbook(Workbook workbook, File archivo) throws IOException {
        try (java.io.FileOutputStream fos = new java.io.FileOutputStream(archivo)) {
            workbook.write(fos);
        }
    }

    public List<String[]> cargarCatalogo(File archivo, int columnaCodigo, int columnaDescripcion) throws IOException {
        List<String[]> resultado = new ArrayList<>();
        try (Workbook workbook = abrirWorkbook(archivo)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                String codigo = leerCelda(row.getCell(columnaCodigo));
                String descripcion = leerCelda(row.getCell(columnaDescripcion));
                if (!codigo.isEmpty()) {
                    resultado.add(new String[]{codigo, descripcion});
                }
            }
        }
        return resultado;
    }

    public String leerCelda(Cell celda) {
        if (celda == null) {
            return "";
        }
        return DATA_FORMATTER.formatCellValue(celda).trim();
    }
}
