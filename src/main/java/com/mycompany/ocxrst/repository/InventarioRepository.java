package com.mycompany.ocxrst.repository;

import com.mycompany.ocxrst.config.AppConfig;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class InventarioRepository {

    public static final String[] ENCABEZADOS = {
        "CODIGO", "DESCRIPCION", "UNIDAD", "PRECIO_UNITARIO", "CANTIDAD", "ACTIVO",
        "CATEGORIA", "MARCA", "STOCK_MINIMO", "OBSERVACIONES", "IMAGEN_RUTA", "SKU"
    };
    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    public File obtenerArchivoInventarios() {
        return AppConfig.resolveBaseFile("INVENTARIOS.xlsx");
    }

    public void asegurarArchivoInventarios() {
        File archivo = obtenerArchivoInventarios();
        if (!archivo.exists()) {
            try (Workbook libro = new XSSFWorkbook()) {
                Sheet hoja = libro.createSheet("INVENTARIOS");
                asegurarEncabezados(hoja);
                guardarLibro(libro, archivo);
            } catch (IOException ex) {
                throw new IllegalStateException("No se pudo crear el archivo de inventarios.", ex);
            }
            return;
        }

        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaInventarios(libro);
            if (asegurarEncabezados(hoja)) {
                guardarLibro(libro, archivo);
            }
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudo preparar el archivo de inventarios.", ex);
        }
    }

    public Workbook cargarLibro(File archivo) throws IOException {
        if (!archivo.exists()) {
            Workbook libroNuevo = new XSSFWorkbook();
            Sheet hojaNueva = libroNuevo.createSheet("INVENTARIOS");
            asegurarEncabezados(hojaNueva);
            guardarLibro(libroNuevo, archivo);
            return libroNuevo;
        }
        try (FileInputStream fis = new FileInputStream(archivo)) {
            return new XSSFWorkbook(fis);
        }
    }

    public void guardarLibro(Workbook libro, File archivo) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(archivo)) {
            libro.write(fos);
        }
    }

    public Sheet obtenerHojaInventarios(Workbook libro) {
        if (libro.getNumberOfSheets() == 0) {
            return libro.createSheet("INVENTARIOS");
        }
        return libro.getSheetAt(0);
    }

    public boolean asegurarEncabezados(Sheet hoja) {
        boolean huboCambios = false;
        Row filaEncabezado = hoja.getRow(0);
        if (filaEncabezado == null) {
            filaEncabezado = hoja.createRow(0);
            huboCambios = true;
        }
        for (int columna = 0; columna < ENCABEZADOS.length; columna++) {
            Cell celda = filaEncabezado.getCell(columna, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            String valorEsperado = ENCABEZADOS[columna];
            if (celda == null) {
                celda = filaEncabezado.createCell(columna);
                celda.setCellValue(valorEsperado);
                huboCambios = true;
                continue;
            }
            String valorActual = DATA_FORMATTER.formatCellValue(celda).trim();
            if (!valorEsperado.equalsIgnoreCase(valorActual)) {
                celda.setCellValue(valorEsperado);
                huboCambios = true;
            }
        }
        return huboCambios;
    }

    public String leerCelda(Row fila, int columna) {
        if (fila == null) {
            return "";
        }
        Cell celda = fila.getCell(columna, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (celda == null) {
            return "";
        }
        return DATA_FORMATTER.formatCellValue(celda).trim();
    }

    public void escribirCelda(Row fila, int columna, String valor) {
        Cell celda = fila.getCell(columna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        celda.setCellValue(valor == null ? "" : valor.trim());
    }

    public void guardarProducto(String codigo, String descripcion, String unidad,
            String precioUnitario, String cantidad, String categoria, String marca,
            String stockMinimo, String observaciones, String imagenRuta, String sku) throws IOException {
        String codigoNormalizado = codigo == null ? "" : codigo.trim();
        if (codigoNormalizado.isBlank()) {
            throw new IllegalArgumentException("El código del producto es obligatorio.");
        }
        try (Workbook libro = cargarLibro(obtenerArchivoInventarios())) {
            Sheet hoja = obtenerHojaInventarios(libro);
            asegurarEncabezados(hoja);
            int filaDestino = -1;
            for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                if (codigoNormalizado.equalsIgnoreCase(leerCelda(hoja.getRow(indice), 0))) {
                    filaDestino = indice;
                    break;
                }
            }
            if (filaDestino < 0) {
                filaDestino = hoja.getLastRowNum() + 1;
            }
            Row fila = hoja.getRow(filaDestino);
            if (fila == null) {
                fila = hoja.createRow(filaDestino);
            }
            escribirCelda(fila, 0, codigoNormalizado);
            escribirCelda(fila, 1, descripcion);
            escribirCelda(fila, 2, unidad);
            escribirCelda(fila, 3, precioUnitario);
            escribirCelda(fila, 4, cantidad);
            escribirCelda(fila, 5, "1");
            escribirCelda(fila, 6, categoria);
            escribirCelda(fila, 7, marca);
            escribirCelda(fila, 8, stockMinimo);
            escribirCelda(fila, 9, observaciones);
            escribirCelda(fila, 10, imagenRuta);
            escribirCelda(fila, 11, sku);
            guardarLibro(libro, obtenerArchivoInventarios());
        }
    }
}
