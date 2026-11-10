package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.model.Usuario;
import com.mycompany.ocxrst.repository.ActividadRepository;
import com.mycompany.ocxrst.repository.InventarioRepository;
import com.mycompany.ocxrst.service.SesionActual;
import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import javax.swing.SwingUtilities;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ActividadEntradaInventarioTest {
    @TempDir
    Path directorio;

    @AfterEach
    void cerrarSesion() {
        SesionActual.cerrar();
    }

    @Test
    void entradaConfirmadaRegistraUnaActividadConFolioYUsuario() throws Exception {
        InventarioRepository inventario = crearInventario();
        ActividadRepository actividades = crearActividades();
        SesionActual.iniciar(new Usuario("operador", "Usuario de prueba"));
        ENTRADAINVENTARIO vista = crearVista(inventario, actividades);
        DefaultTableModel modelo = modelo(vista);

        SwingUtilities.invokeAndWait(() -> {
            modelo.addRow(new Object[]{"COD-1", "Producto 1", "SKU-1", "PZA", "2", "10"});
            modelo.addRow(new Object[]{"COD-2", "Producto 2", "SKU-2", "PZA", "1", "20"});
            try {
                vista.guardarEntrada();
            } catch (IOException ex) {
                throw new AssertionError(ex);
            }
        });

        try (Workbook libro = actividades.abrirLibro(actividades.obtenerArchivoActividades())) {
            Sheet hoja = libro.getSheetAt(0);
            assertEquals(1, hoja.getLastRowNum());
            Row actividad = hoja.getRow(1);
            assertEquals(LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")),
                    actividad.getCell(0).getStringCellValue());
            assertEquals("Inventarios", actividad.getCell(1).getStringCellValue());
            assertEquals("Entrada de inventario", actividad.getCell(2).getStringCellValue());
            assertEquals("EI000001", actividad.getCell(3).getStringCellValue());
            assertEquals("Usuario de prueba", actividad.getCell(4).getStringCellValue());
        }
        try (Workbook libro = inventario.cargarLibro(inventario.obtenerArchivoInventarios())) {
            assertEquals("7", inventario.leerCelda(libro.getSheetAt(0).getRow(1), 4));
            assertEquals("6", inventario.leerCelda(libro.getSheetAt(0).getRow(2), 4));
        }
        try (Workbook libro = inventario.cargarLibro(directorio.resolve("movimientos.xlsx").toFile())) {
            assertEquals(2, libro.getSheetAt(0).getLastRowNum());
        }
    }

    @Test
    void entradaVaciaNoRegistraActividad() throws Exception {
        InventarioRepository inventario = crearInventario();
        ActividadRepository actividades = crearActividades();
        ENTRADAINVENTARIO vista = crearVista(inventario, actividades);

        SwingUtilities.invokeAndWait(() -> assertThrows(IllegalArgumentException.class, vista::guardarEntrada));
        assertFalse(actividades.obtenerArchivoActividades().exists());
    }

    @Test
    void actividadUsaUsuarioCuandoNoHayNombre() throws Exception {
        InventarioRepository inventario = crearInventario();
        ActividadRepository actividades = crearActividades();
        SesionActual.iniciar(new Usuario("operador", ""));
        ENTRADAINVENTARIO vista = crearVista(inventario, actividades);
        DefaultTableModel modelo = modelo(vista);
        SwingUtilities.invokeAndWait(() -> {
            modelo.addRow(new Object[]{"COD-1", "Producto 1", "SKU-1", "PZA", "1", "10"});
            try {
                vista.guardarEntrada();
            } catch (IOException ex) {
                throw new AssertionError(ex);
            }
        });
        try (Workbook libro = actividades.abrirLibro(actividades.obtenerArchivoActividades())) {
            assertEquals("operador", libro.getSheetAt(0).getRow(1).getCell(4).getStringCellValue());
        }
    }

    @Test
    void falloAlGuardarInventarioNoRegistraActividad() throws Exception {
        InventarioRepository original = crearInventario();
        InventarioRepository inventario = new InventarioRepository() {
            @Override
            public File obtenerArchivoInventarios() {
                return original.obtenerArchivoInventarios();
            }

            @Override
            public void guardarLibro(Workbook libro, File archivo) throws IOException {
                if (archivo.equals(obtenerArchivoInventarios())) {
                    throw new IOException("No se pudo guardar el inventario");
                }
                super.guardarLibro(libro, archivo);
            }
        };
        ActividadRepository actividades = crearActividades();
        ENTRADAINVENTARIO vista = crearVista(inventario, actividades);
        DefaultTableModel modelo = modelo(vista);
        SwingUtilities.invokeAndWait(() -> {
            modelo.addRow(new Object[]{"COD-1", "Producto 1", "SKU-1", "PZA", "1", "10"});
            assertThrows(IOException.class, vista::guardarEntrada);
        });
        assertFalse(actividades.obtenerArchivoActividades().exists());
    }

    @Test
    void falloDeActividadInformaQueLaEntradaYaFueGuardada() throws Exception {
        InventarioRepository inventario = crearInventario();
        ActividadRepository actividades = new ActividadRepository() {
            @Override
            public synchronized void registrar(String modulo, String actividad, String referencia, String usuario)
                    throws IOException {
                throw new IOException("No se pudo escribir la actividad");
            }
        };
        ENTRADAINVENTARIO vista = crearVista(inventario, actividades);
        DefaultTableModel modelo = modelo(vista);
        SwingUtilities.invokeAndWait(() -> {
            modelo.addRow(new Object[]{"COD-1", "Producto 1", "SKU-1", "PZA", "1", "10"});
            IOException error = assertThrows(IOException.class, vista::guardarEntrada);
            assertTrue(error.getMessage().contains("No se pudo escribir la actividad"));
        });
        try (Workbook libro = inventario.cargarLibro(inventario.obtenerArchivoInventarios())) {
            assertEquals("6", inventario.leerCelda(libro.getSheetAt(0).getRow(1), 4));
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

    private ENTRADAINVENTARIO crearVista(InventarioRepository inventario, ActividadRepository actividades)
            throws Exception {
        ENTRADAINVENTARIO[] vista = new ENTRADAINVENTARIO[1];
        SwingUtilities.invokeAndWait(() -> vista[0] = new ENTRADAINVENTARIO(
                () -> {}, inventario, actividades, directorio.resolve("movimientos.xlsx").toFile()));
        return vista[0];
    }

    private DefaultTableModel modelo(ENTRADAINVENTARIO vista) throws ReflectiveOperationException {
        java.lang.reflect.Field campo = ENTRADAINVENTARIO.class.getDeclaredField("modelo");
        campo.setAccessible(true);
        return (DefaultTableModel) campo.get(vista);
    }
}
