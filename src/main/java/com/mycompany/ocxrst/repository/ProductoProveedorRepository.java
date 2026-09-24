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

/**
 * Relación muchos-a-muchos entre productos (INVENTARIOS) y proveedores
 * (PROVEEDORES). Un mismo producto puede tener varios proveedores asociados,
 * cada uno con su propio código, precio y tiempo de entrega.
 */
public class ProductoProveedorRepository {

    public static final String[] ENCABEZADOS = {
        "CODIGO_PRODUCTO", "ID_PROVEEDOR", "CODIGO_PROVEEDOR", "PRECIO", "TIEMPO_ENTREGA", "PREDETERMINADO"
    };
    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    public File obtenerArchivoProductoProveedor() {
        return AppConfig.productoProveedorFile();
    }

    public void asegurarArchivoProductoProveedor() {
        File archivo = obtenerArchivoProductoProveedor();
        if (!archivo.exists()) {
            try (Workbook libro = new XSSFWorkbook()) {
                Sheet hoja = libro.createSheet("PRODUCTOPROVEEDORES");
                asegurarEncabezados(hoja);
                guardarLibro(libro, archivo);
            } catch (IOException ex) {
                throw new IllegalStateException("No se pudo crear el archivo de proveedores por producto.", ex);
            }
            return;
        }

        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaProductoProveedor(libro);
            if (asegurarEncabezados(hoja)) {
                guardarLibro(libro, archivo);
            }
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudo preparar el archivo de proveedores por producto.", ex);
        }
    }

    public Workbook cargarLibro(File archivo) throws IOException {
        if (!archivo.exists()) {
            Workbook libroNuevo = new XSSFWorkbook();
            Sheet hojaNueva = libroNuevo.createSheet("PRODUCTOPROVEEDORES");
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

    public Sheet obtenerHojaProductoProveedor(Workbook libro) {
        if (libro.getNumberOfSheets() == 0) {
            return libro.createSheet("PRODUCTOPROVEEDORES");
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

    /**
     * Devuelve todas las filas (proveedores) asociadas a un producto, en el
     * mismo orden en que aparecen en el archivo. Cada fila viene como:
     * {numeroFilaExcel, idProveedor, codigoProveedor, precio, tiempoEntrega, predeterminado}
     */
    public java.util.List<String[]> obtenerProveedoresDeProducto(String codigoProducto) {
        java.util.List<String[]> resultado = new java.util.ArrayList<>();
        asegurarArchivoProductoProveedor();
        File archivo = obtenerArchivoProductoProveedor();
        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaProductoProveedor(libro);
            int ultimaFila = hoja.getLastRowNum();
            for (int numeroFila = 1; numeroFila <= ultimaFila; numeroFila++) {
                Row fila = hoja.getRow(numeroFila);
                if (fila == null) {
                    continue;
                }
                String codigo = leerCelda(fila, 0);
                if (codigo.isEmpty() || !codigo.equalsIgnoreCase(codigoProducto)) {
                    continue;
                }
                resultado.add(new String[]{
                    String.valueOf(numeroFila),
                    leerCelda(fila, 1),
                    leerCelda(fila, 2),
                    leerCelda(fila, 3),
                    leerCelda(fila, 4),
                    leerCelda(fila, 5)
                });
            }
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudieron leer los proveedores del producto.", ex);
        }
        return resultado;
    }

    public void agregarProveedorAProducto(String codigoProducto, String idProveedor, String codigoProveedor,
            String precio, String tiempoEntrega, boolean predeterminado) {
        asegurarArchivoProductoProveedor();
        File archivo = obtenerArchivoProductoProveedor();
        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaProductoProveedor(libro);
            if (predeterminado) {
                quitarPredeterminado(hoja, codigoProducto);
            }
            int nuevaFila = hoja.getLastRowNum() + 1;
            Row fila = hoja.createRow(nuevaFila);
            escribirCelda(fila, 0, codigoProducto);
            escribirCelda(fila, 1, idProveedor);
            escribirCelda(fila, 2, codigoProveedor);
            escribirCelda(fila, 3, precio);
            escribirCelda(fila, 4, tiempoEntrega);
            escribirCelda(fila, 5, predeterminado ? "SI" : "NO");
            guardarLibro(libro, archivo);
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudo agregar el proveedor al producto.", ex);
        }
    }

    public void actualizarProveedorDeProducto(int numeroFilaExcel, String codigoProducto, String idProveedor,
            String codigoProveedor, String precio, String tiempoEntrega, boolean predeterminado) {
        asegurarArchivoProductoProveedor();
        File archivo = obtenerArchivoProductoProveedor();
        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaProductoProveedor(libro);
            if (predeterminado) {
                quitarPredeterminado(hoja, codigoProducto);
            }
            Row fila = hoja.getRow(numeroFilaExcel);
            if (fila == null) {
                fila = hoja.createRow(numeroFilaExcel);
            }
            escribirCelda(fila, 0, codigoProducto);
            escribirCelda(fila, 1, idProveedor);
            escribirCelda(fila, 2, codigoProveedor);
            escribirCelda(fila, 3, precio);
            escribirCelda(fila, 4, tiempoEntrega);
            escribirCelda(fila, 5, predeterminado ? "SI" : "NO");
            guardarLibro(libro, archivo);
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudo actualizar el proveedor del producto.", ex);
        }
    }

    public void marcarProveedorPredeterminado(int numeroFilaExcel, String codigoProducto) {
        asegurarArchivoProductoProveedor();
        File archivo = obtenerArchivoProductoProveedor();
        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaProductoProveedor(libro);
            quitarPredeterminado(hoja, codigoProducto);
            Row fila = hoja.getRow(numeroFilaExcel);
            if (fila != null) {
                escribirCelda(fila, 5, "SI");
            }
            guardarLibro(libro, archivo);
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudo marcar el proveedor como predeterminado.", ex);
        }
    }

    public void eliminarProveedorDeProducto(int numeroFilaExcel) {
        asegurarArchivoProductoProveedor();
        File archivo = obtenerArchivoProductoProveedor();
        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaProductoProveedor(libro);
            Row fila = hoja.getRow(numeroFilaExcel);
            if (fila != null) {
                hoja.removeRow(fila);
            }
            guardarLibro(libro, archivo);
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudo eliminar el proveedor del producto.", ex);
        }
    }

    public void eliminarTodosLosProveedoresDeProducto(String codigoProducto) {
        asegurarArchivoProductoProveedor();
        File archivo = obtenerArchivoProductoProveedor();
        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaProductoProveedor(libro);
            int ultimaFila = hoja.getLastRowNum();
            for (int numeroFila = ultimaFila; numeroFila >= 1; numeroFila--) {
                Row fila = hoja.getRow(numeroFila);
                if (fila == null) {
                    continue;
                }
                String codigo = leerCelda(fila, 0);
                if (codigo.equalsIgnoreCase(codigoProducto)) {
                    hoja.removeRow(fila);
                }
            }
            guardarLibro(libro, archivo);
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudieron eliminar los proveedores del producto.", ex);
        }
    }

    private void quitarPredeterminado(Sheet hoja, String codigoProducto) {
        int ultimaFila = hoja.getLastRowNum();
        for (int numeroFila = 1; numeroFila <= ultimaFila; numeroFila++) {
            Row fila = hoja.getRow(numeroFila);
            if (fila == null) {
                continue;
            }
            String codigo = leerCelda(fila, 0);
            if (codigo.equalsIgnoreCase(codigoProducto)) {
                escribirCelda(fila, 5, "NO");
            }
        }
    }
}
