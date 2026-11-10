package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.repository.InventarioRepository;
import java.awt.Component;
import java.awt.Container;
import java.awt.event.ActionEvent;
import java.awt.event.KeyEvent;
import java.io.File;
import javax.swing.Action;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.SwingUtilities;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ENTRADAINVENTARIOTest {

    @Test
    void skuCompletoAgregaProductoYPermiteCapturasConsecutivas() throws Exception {
        ENTRADAINVENTARIO[] vista = new ENTRADAINVENTARIO[1];
        SwingUtilities.invokeAndWait(() -> vista[0] = nuevaVista());
        JTextField busqueda = campo(vista[0], "producto", JTextField.class);
        JTextField cantidad = campo(vista[0], "cantidad", JTextField.class);
        JTextField costo = campo(vista[0], "costo", JTextField.class);
        DefaultTableModel modelo = campo(vista[0], "modelo", DefaultTableModel.class);

        SwingUtilities.invokeAndWait(() -> busqueda.setText("sku-123"));
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(1, modelo.getRowCount());
            assertEquals(6, modelo.getColumnCount());
            assertEquals("Costo", modelo.getColumnName(5));
            assertEquals("COD-1", modelo.getValueAt(0, 0));
            assertEquals("Producto de prueba", modelo.getValueAt(0, 1));
            assertEquals("SKU-123", modelo.getValueAt(0, 2));
            assertEquals("PZA", modelo.getValueAt(0, 3));
            assertEquals("1", modelo.getValueAt(0, 4));
            assertEquals("42.50", modelo.getValueAt(0, 5));
            assertEquals("", busqueda.getText());
            assertEquals("1", cantidad.getText());
            assertEquals("", costo.getText());
            busqueda.setText("SKU-123");
        });
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(1, modelo.getRowCount());
            assertEquals("2", modelo.getValueAt(0, 4));
            assertEquals("42.50", modelo.getValueAt(0, 5));
            busqueda.setText("SKU-123");
        });
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(1, modelo.getRowCount());
            assertEquals("3", modelo.getValueAt(0, 4));
        });
    }

    @Test
    void productoRepetidoPorCodigoSumaUnoYConservaCostoYOtrosProductos() throws Exception {
        ENTRADAINVENTARIO[] vista = new ENTRADAINVENTARIO[1];
        SwingUtilities.invokeAndWait(() -> vista[0] = nuevaVista());
        JTextField busqueda = campo(vista[0], "producto", JTextField.class);
        JTextField cantidad = campo(vista[0], "cantidad", JTextField.class);
        DefaultTableModel modelo = campo(vista[0], "modelo", DefaultTableModel.class);

        SwingUtilities.invokeAndWait(() -> {
            modelo.addRow(new Object[]{"OTRO", "Otro producto", "OTRO-SKU", "PZA", "4", "15"});
            cantidad.setText("5");
            busqueda.setText("COD-1");
            encontrarBoton(vista[0], "Agregar").doClick();
            assertEquals(2, modelo.getRowCount());
            assertEquals("5", modelo.getValueAt(1, 4));
            modelo.setValueAt("35", 1, 5);
            cantidad.setText("9");
            busqueda.setText("cod-1");
            encontrarBoton(vista[0], "Agregar").doClick();
            assertEquals(2, modelo.getRowCount());
            assertEquals("4", modelo.getValueAt(0, 4));
            assertEquals("6", modelo.getValueAt(1, 4));
            assertEquals("35", modelo.getValueAt(1, 5));
            assertEquals("1", cantidad.getText());
        });
    }

    @Test
    void productoEliminadoPuedeAgregarseDeNuevo() throws Exception {
        ENTRADAINVENTARIO[] vista = new ENTRADAINVENTARIO[1];
        SwingUtilities.invokeAndWait(() -> vista[0] = nuevaVista());
        JTextField busqueda = campo(vista[0], "producto", JTextField.class);
        DefaultTableModel modelo = campo(vista[0], "modelo", DefaultTableModel.class);

        SwingUtilities.invokeAndWait(() -> busqueda.setText("SKU-123"));
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(1, modelo.getRowCount());
            modelo.removeRow(0);
            busqueda.setText("SKU-123");
        });
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(1, modelo.getRowCount());
            assertEquals("1", modelo.getValueAt(0, 4));
        });
    }

    @Test
    void productoRepetidoSumaUnoALaCantidadQueSeEstaEditando() throws Exception {
        ENTRADAINVENTARIO[] vista = new ENTRADAINVENTARIO[1];
        SwingUtilities.invokeAndWait(() -> vista[0] = nuevaVista());
        JTextField busqueda = campo(vista[0], "producto", JTextField.class);
        DefaultTableModel modelo = campo(vista[0], "modelo", DefaultTableModel.class);

        SwingUtilities.invokeAndWait(() -> {
            modelo.addRow(new Object[]{"COD-1", "Producto de prueba", "SKU-123", "PZA", "1", "42.50"});
            JTable tabla = encontrarTabla(vista[0]);
            assertTrue(tabla.editCellAt(0, 4));
            ((JTextField) tabla.getEditorComponent()).setText("7");
            busqueda.setText("SKU-123");
        });
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(1, modelo.getRowCount());
            assertEquals("8", modelo.getValueAt(0, 4));
        });
    }

    @Test
    void busquedaIncompletaNoAgregaYElBotonSigueFuncionando() throws Exception {
        ENTRADAINVENTARIO[] vista = new ENTRADAINVENTARIO[1];
        SwingUtilities.invokeAndWait(() -> vista[0] = nuevaVista());
        JTextField busqueda = campo(vista[0], "producto", JTextField.class);
        DefaultTableModel modelo = campo(vista[0], "modelo", DefaultTableModel.class);

        SwingUtilities.invokeAndWait(() -> busqueda.setText("SKU-12"));
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(0, modelo.getRowCount());
            busqueda.setText("COD-1");
            encontrarBoton(vista[0], "Agregar").doClick();
            assertEquals(1, modelo.getRowCount());
            assertEquals("SKU-123", modelo.getValueAt(0, 2));
        });
    }

    @Test
    void cambioDeTextoInvalidaLaSeleccionAutomaticaPendiente() throws Exception {
        ENTRADAINVENTARIO[] vista = new ENTRADAINVENTARIO[1];
        SwingUtilities.invokeAndWait(() -> vista[0] = nuevaVista());
        JTextField busqueda = campo(vista[0], "producto", JTextField.class);
        DefaultTableModel modelo = campo(vista[0], "modelo", DefaultTableModel.class);

        SwingUtilities.invokeAndWait(() -> {
            busqueda.setText("SKU-123");
            busqueda.setText("otra búsqueda");
        });
        SwingUtilities.invokeAndWait(() -> {
            assertEquals(0, modelo.getRowCount());
            assertEquals("otra búsqueda", busqueda.getText());
        });
    }

    @Test
    void suprimirEliminaSoloLasFilasSeleccionadasInclusoConOrdenamiento() throws Exception {
        ENTRADAINVENTARIO[] vista = new ENTRADAINVENTARIO[1];
        SwingUtilities.invokeAndWait(() -> vista[0] = nuevaVista());
        DefaultTableModel modelo = campo(vista[0], "modelo", DefaultTableModel.class);

        SwingUtilities.invokeAndWait(() -> {
            JTable tabla = encontrarTabla(vista[0]);
            assertNotNull(tabla);
            modelo.addRow(new Object[]{"B", "Segundo", "2", "PZA", "1", "10"});
            modelo.addRow(new Object[]{"A", "Primero", "1", "PZA", "1", "10"});
            modelo.addRow(new Object[]{"C", "Tercero", "3", "PZA", "1", "10"});
            tabla.setAutoCreateRowSorter(true);
            tabla.getRowSorter().toggleSortOrder(0);
            tabla.addRowSelectionInterval(0, 0);
            tabla.addRowSelectionInterval(2, 2);

            Object clave = tabla.getInputMap(JComponent.WHEN_FOCUSED)
                    .get(KeyStroke.getKeyStroke(KeyEvent.VK_DELETE, 0));
            assertNotNull(clave);
            Action accion = tabla.getActionMap().get(clave);
            assertNotNull(accion);
            accion.actionPerformed(new ActionEvent(tabla, ActionEvent.ACTION_PERFORMED, ""));

            assertEquals(1, modelo.getRowCount());
            assertEquals("B", modelo.getValueAt(0, 0));
            tabla.clearSelection();
            accion.actionPerformed(new ActionEvent(tabla, ActionEvent.ACTION_PERFORMED, ""));
            assertEquals(1, modelo.getRowCount());
        });
    }

    private static ENTRADAINVENTARIO nuevaVista() {
        InventarioRepository repository = new InventarioRepository() {
            @Override
            public Workbook cargarLibro(File archivo) {
                Workbook libro = new XSSFWorkbook();
                Sheet hoja = libro.createSheet("INVENTARIOS");
                Row fila = hoja.createRow(1);
                fila.createCell(0).setCellValue("COD-1");
                fila.createCell(1).setCellValue("Producto de prueba");
                fila.createCell(2).setCellValue("PZA");
                fila.createCell(3).setCellValue("42.50");
                fila.createCell(11).setCellValue("SKU-123");
                return libro;
            }
        };
        return new ENTRADAINVENTARIO(() -> {}, repository);
    }

    private static <T> T campo(ENTRADAINVENTARIO vista, String nombre, Class<T> tipo)
            throws ReflectiveOperationException {
        java.lang.reflect.Field campo = ENTRADAINVENTARIO.class.getDeclaredField(nombre);
        campo.setAccessible(true);
        return tipo.cast(campo.get(vista));
    }

    private static JButton encontrarBoton(java.awt.Container contenedor, String texto) {
        for (java.awt.Component componente : contenedor.getComponents()) {
            if (componente instanceof JButton && texto.equals(((JButton) componente).getText())) {
                return (JButton) componente;
            }
            if (componente instanceof java.awt.Container) {
                JButton encontrado = encontrarBoton((java.awt.Container) componente, texto);
                if (encontrado != null) {
                    return encontrado;
                }
            }
        }
        return null;
    }

    private static JTable encontrarTabla(Container contenedor) {
        for (Component componente : contenedor.getComponents()) {
            if (componente instanceof JTable) {
                return (JTable) componente;
            }
            if (componente instanceof Container) {
                JTable encontrada = encontrarTabla((Container) componente);
                if (encontrada != null) {
                    return encontrada;
                }
            }
        }
        return null;
    }
}
