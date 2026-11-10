package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.repository.InventarioRepository;
import com.mycompany.ocxrst.repository.ActividadRepository;
import com.mycompany.ocxrst.util.EstiloAplicacion;
import com.mycompany.ocxrst.config.AppConfig;
import com.mycompany.ocxrst.service.SesionActual;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.KeyEvent;
import java.io.IOException;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import javax.swing.AbstractAction;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.DefaultListModel;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JComponent;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.table.DefaultTableModel;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ENTRADAINVENTARIO extends JPanel {

    enum TipoMovimiento {
        ENTRADA("Entrada", "EI", "ENT"),
        SALIDA("Salida", "SI", "SAL");

        final String nombre;
        final String prefijoFolio;
        final String prefijoDocumento;

        TipoMovimiento(String nombre, String prefijoFolio, String prefijoDocumento) {
            this.nombre = nombre;
            this.prefijoFolio = prefijoFolio;
            this.prefijoDocumento = prefijoDocumento;
        }
    }

    private static final Color AZUL_OSCURO = new Color(28, 82, 132);
    private static final Color AZUL = new Color(31, 112, 184);
    private static final String FAMILIA = "Segoe UI";
    private static final Font FUENTE_BOTON = new Font(FAMILIA, Font.BOLD, 11);
    private static final Font FUENTE_ETIQUETA = new Font(FAMILIA, Font.BOLD, 10);

    private final InventarioRepository repository;
    private final ActividadRepository actividadRepository;
    private final File archivoMovimientos;
    private final TipoMovimiento tipo;
    private final Runnable volver;
    private final DefaultTableModel modelo;
    private JTable tabla;
    private final JTextField producto = new JTextField();
    private final JList<Producto> sugerencias = new JList<>();
    private final JTextField cantidad = new JTextField("1");
    private final JTextField costo = new JTextField();
    private final JTextField movimiento = new JTextField();
    private final JPopupMenu popupSugerencias = new JPopupMenu();
    private final List<Producto> productos = new ArrayList<>();
    private Producto productoSeleccionado;
    private boolean actualizandoProducto;
    private long revisionBusqueda;

    public ENTRADAINVENTARIO(Runnable volver) {
        this(volver, new InventarioRepository());
    }

    ENTRADAINVENTARIO(Runnable volver, InventarioRepository repository) {
        this(volver, repository, new ActividadRepository(),
                AppConfig.resolveBaseFile("MOVIMIENTOSINVENTARIO.xlsx"));
    }

    ENTRADAINVENTARIO(Runnable volver, InventarioRepository repository,
            ActividadRepository actividadRepository, File archivoMovimientos) {
        this(volver, repository, actividadRepository, archivoMovimientos, TipoMovimiento.ENTRADA);
    }

    ENTRADAINVENTARIO(Runnable volver, InventarioRepository repository,
            ActividadRepository actividadRepository, File archivoMovimientos, TipoMovimiento tipo) {
        this.repository = repository;
        this.actividadRepository = actividadRepository;
        this.archivoMovimientos = archivoMovimientos;
        this.tipo = tipo;
        this.volver = volver;
        setLayout(new BorderLayout(0, 10));
        setBackground(Color.WHITE);
        modelo = crearModelo();
        construirVista();
        cargarProductos();
    }

    private DefaultTableModel crearModelo() {
        return new DefaultTableModel(
                new String[]{"Código", "Descripción", "SKU", "UM", "Cantidad", "Costo"}, 0) {
            @Override
            public boolean isCellEditable(int fila, int columna) {
                return columna == 4 || columna == 5;
            }
        };
    }

    private void construirVista() {
        JPanel encabezado = new JPanel(new BorderLayout());
        encabezado.setBackground(AZUL_OSCURO);
        encabezado.setBorder(BorderFactory.createEmptyBorder(10, 14, 10, 14));
        JLabel titulo = new JLabel("▣  " + tipo.nombre + " de Inventario Manual");
        titulo.setForeground(Color.WHITE);
        titulo.setFont(new Font(FAMILIA, Font.BOLD, 18));
        encabezado.add(titulo, BorderLayout.WEST);
        add(encabezado, BorderLayout.NORTH);

        JPanel datos = new JPanel(new GridBagLayout());
        datos.setOpaque(false);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(2, 8, 2, 8);
        movimiento.setText(generarMovimiento());
        movimiento.setEditable(false);
        agregarCampo(datos, gbc, 0, "T.M.", movimiento, 0.8);
        agregarCampo(datos, gbc, 1, "Documento", new JTextField(tipo.prefijoDocumento + "-"
                + new SimpleDateFormat("yyyyMMdd-HHmm").format(new Date())), 1.2);
        agregarCampo(datos, gbc, 2, "Fecha", new JTextField(
                new SimpleDateFormat("dd/MM/yyyy").format(new Date())), 0.9);
        agregarCampo(datos, gbc, 3, "Lote", new JTextField("0.00"), 0.8);
        agregarCampo(datos, gbc, 4, "Referencia", new JTextField(), 1.8);

        JPanel captura = new JPanel(new GridBagLayout());
        captura.setOpaque(false);
        GridBagConstraints capturaGbc = new GridBagConstraints();
        capturaGbc.fill = GridBagConstraints.HORIZONTAL;
        capturaGbc.insets = new Insets(2, 8, 2, 8);
        sugerencias.setVisibleRowCount(5);
        sugerencias.setSelectionMode(javax.swing.ListSelectionModel.SINGLE_SELECTION);
        JScrollPane listaSugerencias = new JScrollPane(sugerencias);
        listaSugerencias.setPreferredSize(new java.awt.Dimension(300, 90));
        popupSugerencias.setBorder(BorderFactory.createLineBorder(new Color(150, 150, 150)));
        popupSugerencias.setFocusable(false);
        popupSugerencias.add(listaSugerencias);
        agregarCampo(captura, capturaGbc, 0, "Código, descripción o SKU",
                producto, 2.4);
        agregarCampo(captura, capturaGbc, 1, "Cantidad", cantidad, 0.9);
        agregarCampo(captura, capturaGbc, 2, "Costo", costo, 0.9);
        JButton agregar = new JButton("Agregar");
        agregar.setFont(FUENTE_BOTON);
        producto.setPreferredSize(new java.awt.Dimension(300, 30));
        producto.setEditable(true);
        producto.setFocusable(true);
        producto.setEnabled(true);
        cantidad.setPreferredSize(new java.awt.Dimension(90, 30));
        costo.setPreferredSize(new java.awt.Dimension(100, 30));
        agregar.setPreferredSize(new java.awt.Dimension(105, 30));
        producto.getDocument().addDocumentListener(new DocumentListener() {
            @Override public void insertUpdate(DocumentEvent evento) {
                actualizarBusqueda(true);
            }

            @Override public void removeUpdate(DocumentEvent evento) {
                actualizarBusqueda(false);
            }

            @Override public void changedUpdate(DocumentEvent evento) {
                actualizarBusqueda(false);
            }
        });
        producto.addFocusListener(new java.awt.event.FocusAdapter() {
            @Override public void focusGained(java.awt.event.FocusEvent evento) {
                filtrarProductos(producto.getText());
            }
        });
        sugerencias.setSelectionMode(javax.swing.ListSelectionModel.SINGLE_SELECTION);
        sugerencias.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override public void mouseClicked(java.awt.event.MouseEvent evento) {
                if (evento.getClickCount() == 1) {
                    seleccionarSugerencia();
                }
            }
        });
        sugerencias.addKeyListener(new java.awt.event.KeyAdapter() {
            @Override public void keyPressed(java.awt.event.KeyEvent evento) {
                if (evento.getKeyCode() == java.awt.event.KeyEvent.VK_ENTER) {
                    seleccionarSugerencia();
                }
            }
        });
        agregar.addActionListener(e -> agregarProducto());
        capturaGbc.gridx = 3;
        capturaGbc.weightx = 0.6;
        capturaGbc.fill = GridBagConstraints.NONE;
        capturaGbc.anchor = GridBagConstraints.SOUTH;
        captura.add(agregar, capturaGbc);

        JPanel superior = new JPanel(new BorderLayout(0, 4));
        superior.setOpaque(false);
        superior.add(datos, BorderLayout.NORTH);
        superior.add(captura, BorderLayout.CENTER);

        tabla = new JTable(modelo);
        EstiloAplicacion.aplicarEstiloTabla(tabla);
        tabla.setRowHeight(38);
        tabla.getInputMap(JComponent.WHEN_FOCUSED).put(
                KeyStroke.getKeyStroke(KeyEvent.VK_DELETE, 0), "eliminarFilas");
        tabla.getActionMap().put("eliminarFilas", new AbstractAction() {
            @Override
            public void actionPerformed(ActionEvent evento) {
                int[] filas = Arrays.stream(tabla.getSelectedRows())
                        .map(tabla::convertRowIndexToModel)
                        .sorted()
                        .toArray();
                for (int indice = filas.length - 1; indice >= 0; indice--) {
                    modelo.removeRow(filas[indice]);
                }
            }
        });
        JScrollPane scroll = new JScrollPane(tabla);
        JPanel contenido = new JPanel(new BorderLayout(0, 8));
        contenido.setOpaque(false);
        contenido.add(superior, BorderLayout.NORTH);
        contenido.add(scroll, BorderLayout.CENTER);
        JLabel ayudaEliminar = new JLabel("Selecciona un renglón y presiona Suprimir para eliminarlo.");
        ayudaEliminar.setFont(FUENTE_ETIQUETA);
        ayudaEliminar.setForeground(AZUL_OSCURO);
        contenido.add(ayudaEliminar, BorderLayout.SOUTH);
        add(contenido, BorderLayout.CENTER);
        javax.swing.SwingUtilities.invokeLater(() -> producto.requestFocusInWindow());

        JPanel botones = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 8));
        botones.setOpaque(false);
        JButton aceptar = new JButton("✓  Aceptar");
        aceptar.setBackground(AZUL);
        aceptar.setForeground(Color.WHITE);
        aceptar.setFont(FUENTE_BOTON);
        aceptar.addActionListener(e -> aceptar());
        JButton cancelar = new JButton("✕  Cancelar");
        cancelar.setFont(FUENTE_BOTON);
        cancelar.addActionListener(e -> volver.run());
        botones.add(aceptar);
        botones.add(cancelar);
        add(botones, BorderLayout.SOUTH);
    }

    private void agregarCampo(JPanel panel, GridBagConstraints gbc, int columna,
            String etiqueta, java.awt.Component campo, double peso) {
        gbc.gridx = columna;
        gbc.gridy = 0;
        gbc.weightx = peso;
        JPanel contenedor = new JPanel(new BorderLayout(0, 3));
        contenedor.setOpaque(false);
        JLabel texto = new JLabel(etiqueta);
        texto.setForeground(AZUL_OSCURO);
        texto.setFont(FUENTE_ETIQUETA);
        contenedor.add(texto, BorderLayout.NORTH);
        contenedor.add(campo, BorderLayout.CENTER);
        panel.add(contenedor, gbc);
    }

    private void cargarProductos() {
        try (Workbook libro = repository.cargarLibro(repository.obtenerArchivoInventarios())) {
            Sheet hoja = repository.obtenerHojaInventarios(libro);
            productos.clear();
            for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                Row fila = hoja.getRow(indice);
                String codigo = repository.leerCelda(fila, 0);
                if (!codigo.isBlank()) {
                    productos.add(new Producto(codigo, repository.leerCelda(fila, 1),
                            repository.leerCelda(fila, 2), repository.leerCelda(fila, 3),
                            repository.leerCelda(fila, 11)));
                }
            }
            filtrarProductos("");
        } catch (IOException ex) {
            mostrarError("No se pudo cargar el catálogo de productos.", ex);
        }
    }

    private void actualizarBusqueda(boolean buscarSku) {
        if (actualizandoProducto) {
            return;
        }
        productoSeleccionado = null;
        long revision = ++revisionBusqueda;
        String texto = producto.getText();
        if (buscarSku && seleccionarProductoPorSku(texto, revision)) {
            return;
        }
        filtrarProductos(texto);
    }

    private void filtrarProductos(String texto) {
            if (actualizandoProducto) {
                return;
            }
            String filtro = texto == null ? "" : texto.trim().toLowerCase();
            DefaultListModel<Producto> modeloSugerencias = new DefaultListModel<>();
            if (filtro.isBlank()) {
                sugerencias.setModel(modeloSugerencias);
                popupSugerencias.setVisible(false);
                return;
            }
            for (Producto item : productos) {
                String datos = (item.codigo + " " + item.descripcion + " " + item.sku).toLowerCase();
                if (datos.contains(filtro)) {
                    modeloSugerencias.addElement(item);
                }
            }
            sugerencias.setModel(modeloSugerencias);
            if (producto.isFocusOwner() && modeloSugerencias.getSize() > 0) {
                popupSugerencias.setPopupSize(producto.getWidth(), 90);
                popupSugerencias.show(producto, 0, producto.getHeight());
            } else {
                popupSugerencias.setVisible(false);
            }
        }

    private boolean seleccionarProductoPorSku(String texto, long revision) {
        String skuCapturado = texto == null ? "" : texto.trim();
        if (skuCapturado.isBlank()) {
            return false;
        }
        for (Producto item : productos) {
            if (item.sku.equalsIgnoreCase(skuCapturado)) {
                javax.swing.SwingUtilities.invokeLater(() -> {
                    if (revision != revisionBusqueda || !producto.getText().equals(texto)) {
                        return;
                    }
                    productoSeleccionado = item;
                    costo.setText(item.costo);
                    popupSugerencias.setVisible(false);
                    agregarProducto();
                });
                return true;
            }
        }
        return false;
    }

        private void seleccionarSugerencia() {
            Producto seleccionado = sugerencias.getSelectedValue();
            if (seleccionado == null && sugerencias.getModel().getSize() > 0) {
                seleccionado = sugerencias.getModel().getElementAt(0);
            }
            if (seleccionado == null) {
                return;
            }
            productoSeleccionado = seleccionado;
            actualizandoProducto = true;
            try {
                producto.setText(seleccionado.codigo + " - " + seleccionado.descripcion);
            } finally {
                actualizandoProducto = false;
            }
            costo.setText(seleccionado.costo);
            popupSugerencias.setVisible(false);
        }

    private void agregarProducto() {
        Producto seleccionado = productoSeleccionado();
        if (seleccionado == null) {
            mostrarError("Selecciona un producto por código o descripción.",
                    new IllegalArgumentException("Producto no seleccionado."));
            return;
        }
        if (tabla.isEditing() && !tabla.getCellEditor().stopCellEditing()) {
            mostrarError("No se pudo actualizar la tabla.",
                    new IllegalArgumentException("Termina de editar la celda antes de agregar el producto."));
            return;
        }
        int filaExistente = -1;
        for (int fila = 0; fila < modelo.getRowCount(); fila++) {
            if (seleccionado.codigo.equalsIgnoreCase(String.valueOf(modelo.getValueAt(fila, 0)).trim())) {
                filaExistente = fila;
                break;
            }
        }
        try {
            if (filaExistente >= 0) {
                double cantidadExistente = numero(String.valueOf(modelo.getValueAt(filaExistente, 4)));
                if (!Double.isFinite(cantidadExistente) || cantidadExistente <= 0) {
                    throw new IllegalArgumentException("La cantidad del producto en la tabla debe ser mayor que cero.");
                }
                modelo.setValueAt(formatoNumero(cantidadExistente + 1), filaExistente, 4);
            } else {
                String cantidadTexto = cantidad.getText().trim();
                double cantidadNueva = numero(cantidadTexto);
                if (!Double.isFinite(cantidadNueva) || cantidadNueva <= 0) {
                    throw new IllegalArgumentException("Captura una cantidad mayor que cero.");
                }
                modelo.addRow(new Object[]{seleccionado.codigo, seleccionado.descripcion,
                    seleccionado.sku, seleccionado.unidad, cantidadTexto,
                    costo.getText().trim().isBlank() ? seleccionado.costo : costo.getText().trim()});
            }
        } catch (IllegalArgumentException ex) {
            mostrarError("No se pudo agregar el producto.", ex);
            return;
        }
        productoSeleccionado = null;
        producto.setText("");
        cantidad.setText("1");
        costo.setText("");
    }

    private Producto productoSeleccionado() {
        if (productoSeleccionado != null) {
            return productoSeleccionado;
        }
        String texto = producto.getText().trim();
        if (texto.isBlank()) {
            return null;
        }
        for (Producto item : productos) {
            if (item.codigo.equalsIgnoreCase(texto)
                    || item.descripcion.toLowerCase().contains(texto.toLowerCase())
                    || item.sku.equalsIgnoreCase(texto)) {
                return item;
            }
        }
        return null;
    }

    private void aceptar() {
        int confirmacion = JOptionPane.showConfirmDialog(this,
                "¿Deseas realizar la " + tipo.nombre.toLowerCase(java.util.Locale.ROOT) + " de inventario?",
                "Confirmar " + tipo.nombre.toLowerCase(java.util.Locale.ROOT) + " manual",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.QUESTION_MESSAGE);
        if (confirmacion != JOptionPane.YES_OPTION) {
            return;
        }
        try {
            guardarEntrada();
            JOptionPane.showMessageDialog(this, tipo.nombre + " de inventario registrada correctamente.");
            volver.run();
        } catch (ActividadEntradaException ex) {
            mostrarError("El movimiento se guardó, pero no se pudo registrar la actividad en el tablero.", ex);
            volver.run();
        } catch (IOException | IllegalArgumentException ex) {
            mostrarError("No se pudo registrar el movimiento de inventario.", ex);
        }
    }

    void guardarEntrada() throws IOException {
        if (modelo.getRowCount() == 0) {
            throw new IllegalArgumentException("Agrega al menos un producto.");
        }
        if (tabla.isEditing() && !tabla.getCellEditor().stopCellEditing()) {
            throw new IllegalArgumentException("Termina de editar la celda antes de guardar.");
        }
        String usuario = SesionActual.nombre();
        if (usuario == null || usuario.isBlank()) {
            usuario = SesionActual.usuario();
        }
        if (usuario == null || usuario.isBlank()) {
            usuario = "Usuario no identificado";
        }
        File archivo = repository.obtenerArchivoInventarios();
        try (Workbook libro = repository.cargarLibro(archivo)) {
            Sheet hoja = repository.obtenerHojaInventarios(libro);
            String folio = movimiento.getText();
            for (int indice = 0; indice < modelo.getRowCount(); indice++) {
                double cantidadMovimiento = numero(String.valueOf(modelo.getValueAt(indice, 4)));
                if (!Double.isFinite(cantidadMovimiento) || cantidadMovimiento <= 0) {
                    throw new IllegalArgumentException("La cantidad debe ser mayor que cero.");
                }
                int filaIndice = buscarFilaPorCodigo(hoja, String.valueOf(modelo.getValueAt(indice, 0)));
                if (filaIndice < 1) {
                    throw new IllegalArgumentException("No se encontró el producto seleccionado.");
                }
                Row fila = hoja.getRow(filaIndice);
                double existencia = numero(repository.leerCelda(fila, 4));
                if (!Double.isFinite(existencia)) {
                    throw new IllegalArgumentException("La existencia del producto no es válida.");
                }
                if (tipo == TipoMovimiento.SALIDA && cantidadMovimiento > existencia) {
                    throw new IllegalArgumentException("Existencia insuficiente para "
                            + repository.leerCelda(fila, 0) + ". Disponible: " + formatoNumero(existencia) + ".");
                }
                double nuevaExistencia = tipo == TipoMovimiento.ENTRADA
                        ? existencia + cantidadMovimiento : existencia - cantidadMovimiento;
                repository.escribirCelda(fila, 4, formatoNumero(nuevaExistencia));
                if (tipo == TipoMovimiento.ENTRADA) {
                    repository.escribirCelda(fila, 3, String.valueOf(modelo.getValueAt(indice, 5)));
                    repository.escribirCelda(fila, 11, String.valueOf(modelo.getValueAt(indice, 2)));
                }
            }
            for (int indice = 0; indice < modelo.getRowCount(); indice++) {
                registrarMovimiento(folio, modelo, indice, usuario);
            }
            repository.guardarLibro(libro, archivo);
        }
        try {
            actividadRepository.registrar("Inventarios", tipo.nombre + " de inventario",
                    movimiento.getText(), usuario);
        } catch (IOException ex) {
            throw new ActividadEntradaException(ex);
        }
    }

    private int buscarFilaPorCodigo(Sheet hoja, String codigo) {
        for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
            if (codigo.equalsIgnoreCase(repository.leerCelda(hoja.getRow(indice), 0))) {
                return indice;
            }
        }
        return -1;
    }

    private double numero(String valor) {
        try {
            return Double.parseDouble(valor.replace(",", "").replace("$", "").trim());
        } catch (NumberFormatException ex) {
            throw new IllegalArgumentException("La cantidad o el costo no tiene un formato válido.");
        }
    }

    private String formatoNumero(double valor) {
        return valor == Math.rint(valor) ? String.valueOf((long) valor) : String.valueOf(valor);
    }

    private String generarMovimiento() {
        return tipo.prefijoFolio + String.format("%06d", obtenerSiguienteMovimiento());
    }

    private int obtenerSiguienteMovimiento() {
        int mayor = 0;
        File archivo = archivoMovimientos;
        if (archivo.exists()) {
            try (Workbook libro = repository.cargarLibro(archivo)) {
                Sheet hoja = repository.obtenerHojaInventarios(libro);
                for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                    String folio = repository.leerCelda(hoja.getRow(indice), 0);
                    if (folio.startsWith(tipo.prefijoFolio) && folio.length() == 8) {
                        mayor = Math.max(mayor, Integer.parseInt(folio.substring(2)));
                    }
                }
            } catch (IOException | NumberFormatException ex) {
                throw new IllegalStateException("No se pudo generar el folio de inventario.", ex);
            }
        }
        return mayor + 1;
    }

    private void registrarMovimiento(String folio, DefaultTableModel tabla, int indice, String usuario)
            throws IOException {
        File archivo = archivoMovimientos;
        try (Workbook libro = archivo.exists() ? repository.cargarLibro(archivo) : new XSSFWorkbook()) {
            Sheet hoja = repository.obtenerHojaInventarios(libro);
            if (hoja.getRow(0) == null || !"FOLIO".equalsIgnoreCase(
                    repository.leerCelda(hoja.getRow(0), 0))) {
                if (hoja.getRow(0) != null) {
                    hoja.removeRow(hoja.getRow(0));
                }
                Row encabezado = hoja.createRow(0);
                String[] columnas = {"FOLIO", "FECHA", "CODIGO", "SKU", "CANTIDAD", "COSTO", "USUARIO"};
                for (int columna = 0; columna < columnas.length; columna++) {
                    encabezado.createCell(columna).setCellValue(columnas[columna]);
                }
            }
            Row encabezado = hoja.getRow(0);
            if (encabezado.getCell(6) == null) {
                encabezado.createCell(6).setCellValue("USUARIO");
            }
            Row fila = hoja.createRow(hoja.getLastRowNum() + 1);
            repository.escribirCelda(fila, 0, folio);
            repository.escribirCelda(fila, 1, new SimpleDateFormat("dd/MM/yyyy").format(new Date()));
            repository.escribirCelda(fila, 2, String.valueOf(tabla.getValueAt(indice, 0)));
            repository.escribirCelda(fila, 3, String.valueOf(tabla.getValueAt(indice, 2)));
            repository.escribirCelda(fila, 4, String.valueOf(tabla.getValueAt(indice, 4)));
            repository.escribirCelda(fila, 5, String.valueOf(tabla.getValueAt(indice, 5)));
            repository.escribirCelda(fila, 6, usuario);
            repository.guardarLibro(libro, archivo);
        }
    }

    private void mostrarError(String mensaje, Exception ex) {
        JOptionPane.showMessageDialog(this, ex.getMessage(), mensaje, JOptionPane.ERROR_MESSAGE);
    }

    private static final class ActividadEntradaException extends IOException {
        private ActividadEntradaException(IOException causa) {
            super(causa.getMessage(), causa);
        }
    }

    private static final class Producto {
        private final String codigo;
        private final String descripcion;
        private final String unidad;
        private final String costo;
        private final String sku;

        private Producto(String codigo, String descripcion, String unidad,
                String costo, String sku) {
            this.codigo = codigo;
            this.descripcion = descripcion;
            this.unidad = unidad;
            this.costo = costo;
            this.sku = sku;
        }

        @Override
        public String toString() {
            return codigo + " - " + descripcion;
        }
    }
}
