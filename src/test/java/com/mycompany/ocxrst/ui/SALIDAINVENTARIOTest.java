package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.model.Usuario;
import com.mycompany.ocxrst.repository.ActividadRepository;
import com.mycompany.ocxrst.repository.InventarioRepository;
import com.mycompany.ocxrst.service.SesionActual;
import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class SALIDAINVENTARIOTest {
    @TempDir
    Path directorio;

    @AfterEach
    void cerrarSesion() {
        SesionActual.cerrar();
    }

    @Test
    void salidaRestaExistenciaYRegistraMovimientoYActividad() throws Exception {
        InventarioRepository inventario = crearInventario();
        ActividadRepository actividades = crearActividades();
        SesionActual.iniciar(new Usuario("operador", "Usuario de prueba"));
        SALIDAINVENTARIO salida = crearVista(inventario, actividades);
        DefaultTableModel modelo = campo(salida, "modelo", DefaultTableModel.class);
        JTextField busqueda = campo(salida, "producto", JTextField.class);

        SwingUtilities.invokeAndWait(() -> busqueda.setText("SKU-1"));
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(1, modelo.getRowCount());
            busqueda.setText("SKU-1");
        });
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(1, modelo.getRowCount());
            assertEquals("2", modelo.getValueAt(0, 4));
            modelo.addRow(new Object[]{"COD-2", "Producto 2", "SKU-2", "PZA", "5", "99"});
            try {
                salida.guardarEntrada();
            } catch (IOException ex) {
                throw new AssertionError(ex);
            }
        });

        try (Workbook libro = inventario.cargarLibro(inventario.obtenerArchivoInventarios())) {
            Sheet hoja = libro.getSheetAt(0);
            assertEquals("3", inventario.leerCelda(hoja.getRow(1), 4));
            assertEquals("0", inventario.leerCelda(hoja.getRow(2), 4));
            assertEquals("10", inventario.leerCelda(hoja.getRow(2), 3));
            assertEquals("SKU-2", inventario.leerCelda(hoja.getRow(2), 11));
        }
        try (Workbook libro = inventario.cargarLibro(directorio.resolve("movimientos.xlsx").toFile())) {
            Sheet hoja = libro.getSheetAt(0);
            assertEquals(2, hoja.getLastRowNum());
            assertEquals("SI000001", inventario.leerCelda(hoja.getRow(1), 0));
            assertEquals("2", inventario.leerCelda(hoja.getRow(1), 4));
            assertEquals("Usuario de prueba", inventario.leerCelda(hoja.getRow(1), 6));
            assertEquals("SI000001", inventario.leerCelda(hoja.getRow(2), 0));
        }
        try (Workbook libro = actividades.abrirLibro(actividades.obtenerArchivoActividades())) {
            Row actividad = libro.getSheetAt(0).getRow(1);
            assertEquals("Inventarios", actividad.getCell(1).getStringCellValue());
            assertEquals("Salida de inventario", actividad.getCell(2).getStringCellValue());
            assertEquals("SI000001", actividad.getCell(3).getStringCellValue());
            assertEquals("Usuario de prueba", actividad.getCell(4).getStringCellValue());
        }
        DefaultTableModel historial = new DefaultTableModel(0, 7);
        INVENTARIOS.agregarMovimientosManuales(historial, "COD-2",
                directorio.resolve("movimientos.xlsx").toFile(), inventario,
                directorio.resolve("actividades.xlsx").toFile());
        assertEquals(1, historial.getRowCount());
        assertEquals("Salida", historial.getValueAt(0, 1));
        assertEquals("SI000001", historial.getValueAt(0, 2));
        assertEquals("N/A", historial.getValueAt(0, 3));
        assertEquals("Usuario de prueba", historial.getValueAt(0, 5));
    }

    @Test
    void historialDistingueEntradasAnterioresYFoliosDeSalida() throws Exception {
        InventarioRepository inventario = crearInventario();
        File archivo = directorio.resolve("movimientos.xlsx").toFile();
        try (Workbook libro = new XSSFWorkbook()) {
            Sheet hoja = libro.createSheet("INVENTARIOS");
            Row encabezado = hoja.createRow(0);
            String[] columnas = {"FOLIO", "FECHA", "CODIGO", "SKU", "CANTIDAD", "COSTO"};
            for (int indice = 0; indice < columnas.length; indice++) {
                encabezado.createCell(indice).setCellValue(columnas[indice]);
            }
            Row entrada = hoja.createRow(1);
            entrada.createCell(0).setCellValue("EI000012");
            entrada.createCell(2).setCellValue("COD-1");
            entrada.createCell(4).setCellValue("1");
            inventario.guardarLibro(libro, archivo);
        }
        crearActividades().registrar("Inventarios", "Entrada de inventario",
                "EI000012", "Usuario anterior");
        SALIDAINVENTARIO salida = crearVista(inventario, crearActividades());
        DefaultTableModel modelo = campo(salida, "modelo", DefaultTableModel.class);
        SwingUtilities.invokeAndWait(() -> {
            modelo.addRow(new Object[]{"COD-1", "Producto 1", "SKU-1", "PZA", "1", "10"});
            try {
                salida.guardarEntrada();
            } catch (IOException ex) {
                throw new AssertionError(ex);
            }
        });
        DefaultTableModel historial = new DefaultTableModel(0, 7);
        INVENTARIOS.agregarMovimientosManuales(historial, "COD-1", archivo, inventario,
                directorio.resolve("actividades.xlsx").toFile());
        assertEquals(2, historial.getRowCount());
        assertEquals("Entrada", historial.getValueAt(0, 1));
        assertEquals("N/A", historial.getValueAt(0, 3));
        assertEquals("Usuario anterior", historial.getValueAt(0, 5));
        assertEquals("Salida", historial.getValueAt(1, 1));
        assertEquals("SI000001", historial.getValueAt(1, 2));
        SALIDAINVENTARIO siguiente = crearVista(inventario, crearActividades());
        assertEquals("SI000002", campo(siguiente, "movimiento", JTextField.class).getText());
    }

    @Test
    void insuficienciaEnCualquierRenglonBloqueaTodoSinRegistro() throws Exception {
        InventarioRepository inventario = crearInventario();
        ActividadRepository actividades = crearActividades();
        SALIDAINVENTARIO salida = crearVista(inventario, actividades);
        DefaultTableModel modelo = campo(salida, "modelo", DefaultTableModel.class);
        SwingUtilities.invokeAndWait(() -> {
            modelo.addRow(new Object[]{"COD-1", "Producto 1", "SKU-1", "PZA", "2", "10"});
            modelo.addRow(new Object[]{"COD-2", "Producto 2", "SKU-2", "PZA", "6", "10"});
            IllegalArgumentException error = assertThrows(IllegalArgumentException.class, salida::guardarEntrada);
            assertTrue(error.getMessage().contains("COD-2"));
            assertTrue(error.getMessage().contains("Disponible: 5"));
        });
        try (Workbook libro = inventario.cargarLibro(inventario.obtenerArchivoInventarios())) {
            assertEquals("5", inventario.leerCelda(libro.getSheetAt(0).getRow(1), 4));
            assertEquals("5", inventario.leerCelda(libro.getSheetAt(0).getRow(2), 4));
        }
        assertFalse(directorio.resolve("movimientos.xlsx").toFile().exists());
        assertFalse(actividades.obtenerArchivoActividades().exists());
    }

    @Test
    void guardaLaCantidadEditadaYRechazaCantidadNoPositiva() throws Exception {
        InventarioRepository inventario = crearInventario();
        SALIDAINVENTARIO salida = crearVista(inventario, crearActividades());
        DefaultTableModel modelo = campo(salida, "modelo", DefaultTableModel.class);
        JTable tabla = campo(salida, "tabla", JTable.class);
        SwingUtilities.invokeAndWait(() -> {
            modelo.addRow(new Object[]{"COD-1", "Producto 1", "SKU-1", "PZA", "1", "10"});
            assertTrue(tabla.editCellAt(0, 4));
            ((JTextField) tabla.getEditorComponent()).setText("0");
            assertThrows(IllegalArgumentException.class, salida::guardarEntrada);
            modelo.setValueAt("5", 0, 4);
            try {
                salida.guardarEntrada();
            } catch (IOException ex) {
                throw new AssertionError(ex);
            }
        });
        try (Workbook libro = inventario.cargarLibro(inventario.obtenerArchivoInventarios())) {
            assertEquals("0", inventario.leerCelda(libro.getSheetAt(0).getRow(1), 4));
        }
    }

    private InventarioRepository crearInventario() throws IOException {
        File archivo = directorio.resolve("inventario.xlsx").toFile();
        InventarioRepository inventario = new InventarioRepository() {
            @Override
            public File obtenerArchivoInventarios() {
                return archivo;
            }
        };
        try (Workbook libro = new XSSFWorkbook()) {
            Sheet hoja = libro.createSheet("INVENTARIOS");
            for (int indice = 1; indice <= 2; indice++) {
                Row fila = hoja.createRow(indice);
                fila.createCell(0).setCellValue("COD-" + indice);
                fila.createCell(1).setCellValue("Producto " + indice);
                fila.createCell(2).setCellValue("PZA");
                fila.createCell(3).setCellValue("10");
                fila.createCell(4).setCellValue("5");
                fila.createCell(11).setCellValue("SKU-" + indice);
            }
            inventario.guardarLibro(libro, archivo);
        }
        return inventario;
    }

    private ActividadRepository crearActividades() {
        return new ActividadRepository() {
            @Override
            public File obtenerArchivoActividades() {
                return directorio.resolve("actividades.xlsx").toFile();
            }
        };
    }

    private SALIDAINVENTARIO crearVista(InventarioRepository inventario, ActividadRepository actividades)
            throws Exception {
        SALIDAINVENTARIO[] vista = new SALIDAINVENTARIO[1];
        SwingUtilities.invokeAndWait(() -> vista[0] = new SALIDAINVENTARIO(
                () -> {}, inventario, actividades, directorio.resolve("movimientos.xlsx").toFile()));
        return vista[0];
    }

    private <T> T campo(SALIDAINVENTARIO vista, String nombre, Class<T> tipo)
            throws ReflectiveOperationException {
        java.lang.reflect.Field campo = ENTRADAINVENTARIO.class.getDeclaredField(nombre);
        campo.setAccessible(true);
        return tipo.cast(campo.get(vista));
    }
}
