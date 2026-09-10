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

public class ProveedorRepository {

    public static final String[] ENCABEZADOS = {
        "ID_PROVEEDOR", "NOMBRE_RAZON_SOCIAL", "RFC", "TELEFONO", "CORREO",
        "DIRECCION", "CONTACTO", "PAGO", "METODO_PAGO", "TIEMPO_ENTREGA", "FORMA_PAGO", "ACTIVO"
    };
    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    public File obtenerArchivoProveedores() {
        return AppConfig.resolveBaseFile("PROVEEDORES.xlsx");
    }

    public void asegurarArchivoProveedores() {
        File archivo = obtenerArchivoProveedores();
        if (!archivo.exists()) {
            try (Workbook libro = new XSSFWorkbook()) {
                Sheet hoja = libro.createSheet("PROVEEDORES");
                asegurarEncabezados(hoja);
                guardarLibro(libro, archivo);
            } catch (IOException ex) {
                throw new IllegalStateException("No se pudo crear el archivo de proveedores.", ex);
            }
            return;
        }

        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaProveedores(libro);
            if (asegurarEncabezados(hoja)) {
                guardarLibro(libro, archivo);
            }
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudo preparar el archivo de proveedores.", ex);
        }
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

    public Sheet obtenerHojaProveedores(Workbook libro) {
        if (libro.getNumberOfSheets() == 0) {
            return libro.createSheet("PROVEEDORES");
        }
        return libro.getSheetAt(0);
    }

    public Workbook cargarLibro(File archivo) throws IOException {
        if (!archivo.exists()) {
            Workbook libroNuevo = new XSSFWorkbook();
            Sheet hojaNueva = libroNuevo.createSheet("PROVEEDORES");
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
}
