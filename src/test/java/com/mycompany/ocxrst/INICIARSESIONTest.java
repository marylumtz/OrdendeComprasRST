package com.mycompany.ocxrst;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class INICIARSESIONTest {

    @Test
    void obtenerNombreUsuarioExcelDevuelveNombreDeLaColumnaTres() throws IOException {
        try (Workbook libro = new XSSFWorkbook()) {
            Sheet hoja = libro.createSheet("USUARIOS");
            Row fila = hoja.createRow(1);
            fila.createCell(0).setCellValue("admin");
            fila.createCell(1).setCellValue("123");
            fila.createCell(2).setCellValue("Maria Luisa Martinez");

            assertEquals("Maria Luisa Martinez", INICIARSESION.obtenerNombreUsuarioExcel(libro, "admin", "123"));
        }
    }
}
