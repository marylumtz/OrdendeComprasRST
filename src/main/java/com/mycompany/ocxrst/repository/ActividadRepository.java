package com.mycompany.ocxrst.repository;

import com.mycompany.ocxrst.config.AppConfig;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ActividadRepository {

    public static final String[] ENCABEZADOS = {
        "FECHA", "MODULO", "ACTIVIDAD", "REFERENCIA", "USUARIO"
    };
    private static final DateTimeFormatter FORMATO_FECHA = DateTimeFormatter.ofPattern("dd/MM/yyyy");

    public File obtenerArchivoActividades() {
        return AppConfig.actividadesFile();
    }

    public synchronized void registrar(String modulo, String actividad, String referencia, String usuario)
            throws IOException {
        File archivo = obtenerArchivoActividades();
        try (Workbook libro = abrirLibro(archivo)) {
            Sheet hoja = libro.getSheetAt(0);
            asegurarEncabezados(hoja);
            if (hoja.getLastRowNum() >= 1) {
                hoja.shiftRows(1, hoja.getLastRowNum(), 1);
            }
            Row fila = hoja.createRow(1);
            escribir(fila, 0, LocalDate.now().format(FORMATO_FECHA));
            escribir(fila, 1, modulo);
            escribir(fila, 2, actividad);
            escribir(fila, 3, referencia);
            escribir(fila, 4, usuario);
            try (FileOutputStream salida = new FileOutputStream(archivo)) {
                libro.write(salida);
            }
        }
    }

    public Workbook abrirLibro(File archivo) throws IOException {
        if (!archivo.exists()) {
            Workbook libro = new XSSFWorkbook();
            libro.createSheet("ACTIVIDADES");
            asegurarEncabezados(libro.getSheetAt(0));
            try (FileOutputStream salida = new FileOutputStream(archivo)) {
                libro.write(salida);
            }
            libro.close();
        }
        try (FileInputStream entrada = new FileInputStream(archivo)) {
            return new XSSFWorkbook(entrada);
        }
    }

    private void asegurarEncabezados(Sheet hoja) {
        Row encabezado = hoja.getRow(0);
        if (encabezado == null) {
            encabezado = hoja.createRow(0);
        }
        for (int columna = 0; columna < ENCABEZADOS.length; columna++) {
            encabezado.createCell(columna).setCellValue(ENCABEZADOS[columna]);
        }
    }

    private void escribir(Row fila, int columna, String valor) {
        fila.createCell(columna).setCellValue(valor == null ? "" : valor.trim());
    }
}