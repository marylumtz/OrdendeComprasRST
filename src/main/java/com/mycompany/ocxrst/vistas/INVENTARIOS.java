/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.IconoVentanaUtil;
import com.mycompany.ocxrst.EstiloAplicacion;
import com.mycompany.ocxrst.repository.InventarioRepository;
import com.mycompany.ocxrst.repository.ProductoProveedorRepository;
import com.mycompany.ocxrst.repository.ProveedorRepository;
import com.mycompany.ocxrst.service.InventarioService;
import java.io.File;
import java.io.IOException;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Maria Luisa Martinez
 */
public class INVENTARIOS extends javax.swing.JFrame {

    private static final java.util.logging.Logger logger = java.util.logging.Logger.getLogger(INVENTARIOS.class.getName());
    private final InventarioRepository inventarioRepository = new InventarioRepository();
    private final ProveedorRepository proveedorRepository = new ProveedorRepository();
    private final ProductoProveedorRepository productoProveedorRepository = new ProductoProveedorRepository();
    private int filaInventarioActual = -1;
    private String rutaImagenActual = "";
    private final java.util.List<String[]> filasProveedorProducto = new java.util.ArrayList<>();

    // Caché de productos existentes (código | descripción) para el autocompletado del campo CÓDIGO.
    private final java.util.List<String[]> cacheProductos = new java.util.ArrayList<>();
    private final javax.swing.JPopupMenu popupProductos = new javax.swing.JPopupMenu();
    private boolean seleccionandoProducto = false;

    private static final String[] UNIDADES_PRESET = {"PZA", "KG", "LT", "MT", "CAJA", "PAQUETE", "SERVICIO"};
    private static final String[] CATEGORIAS_PRESET = {
        "General", "Electr\u00f3nica", "Ferreter\u00eda", "Oficina", "Limpieza", "Consumibles", "Herramientas", "Refacciones"
    };
    private static final String[] MARCAS_PRESET = {"Gen\u00e9rico", "Sin marca"};

    private static final int COL_PREDETERMINADO = 5;
    private static final int COL_EDITAR = 6;
    private static final int COL_ELIMINAR = 7;

    // ===================== Paleta de colores (misma que PROVEEDORES) =====================
    private static final java.awt.Color COLOR_AZUL_OSCURO = new java.awt.Color(28, 82, 132);
    private static final java.awt.Color COLOR_AZUL = new java.awt.Color(31, 112, 184);
    private static final java.awt.Color COLOR_FONDO = new java.awt.Color(237, 245, 253);
    private static final java.awt.Color COLOR_BORDE_CAMPO = new java.awt.Color(198, 210, 225);
    private static final java.awt.Color COLOR_VERDE = new java.awt.Color(0, 130, 60);
    private static final java.awt.Color COLOR_ROJO = new java.awt.Color(190, 0, 0);
    private static final java.awt.Color COLOR_AZUL_CLARO = new java.awt.Color(219, 234, 248);
    private static final java.awt.Color COLOR_AZUL_PALIDO = new java.awt.Color(240, 246, 253);
    private static final java.awt.Color COLOR_TEXTO_GRIS = new java.awt.Color(110, 122, 138);
    private static final java.awt.Color COLOR_VERDE_CLARO = new java.awt.Color(224, 246, 233);
    private static final java.awt.Color COLOR_ROJO_CLARO = new java.awt.Color(253, 226, 226);

    // ===================== Campos de la interfaz =====================
    private javax.swing.JPanel jPanel1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton6;
    private javax.swing.JButton botonGuardar;
    private javax.swing.JButton botonLimpiar;
    private javax.swing.JButton botonAgregarProveedor;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField2;
    private javax.swing.JTextField jTextField5;
    private javax.swing.JTextField jTextField6;
    private javax.swing.JComboBox<String> comboUnidad;
    private javax.swing.JComboBox<String> comboCategoria;
    private javax.swing.JComboBox<String> comboMarca;
    private javax.swing.JTextField campoStockMinimo;
    private javax.swing.JTextArea areaObservaciones;
    private javax.swing.JComboBox<String> comboEstado;
    private javax.swing.JLabel etiquetaId;
    private javax.swing.JLabel etiquetaImagen;
    private javax.swing.JTable tablaProveedoresProducto;
    private javax.swing.table.DefaultTableModel modeloProveedoresProducto;

    /**
     * Creates new form INVENTARIOS
     */
    public INVENTARIOS() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        initComponents();
        configurarDisenoVisual();
        IconoVentanaUtil.aplicar(this);
        setSize(1200, 820);
        setResizable(false);
        setLocationRelativeTo(null);
        asegurarArchivoInventarios();
        configurarEventos();
        configurarAutocompletadoCodigo();
        cargarCacheProductos();
    }

    private void initComponents() {
        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("INVENTARIOS");

        jPanel1 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel("INVENTARIOS");
        jButton3 = new javax.swing.JButton();
        jButton6 = new javax.swing.JButton();
        botonGuardar = new javax.swing.JButton();
        botonLimpiar = new javax.swing.JButton();
        botonAgregarProveedor = new javax.swing.JButton();
        jTextField1 = new javax.swing.JTextField();
        jTextField2 = new javax.swing.JTextField();
        jTextField5 = new javax.swing.JTextField();
        jTextField6 = new javax.swing.JTextField();
        comboUnidad = new javax.swing.JComboBox<>(UNIDADES_PRESET);
        comboCategoria = new javax.swing.JComboBox<>(CATEGORIAS_PRESET);
        comboMarca = new javax.swing.JComboBox<>(MARCAS_PRESET);
        campoStockMinimo = new javax.swing.JTextField();
        areaObservaciones = new javax.swing.JTextArea();
        comboEstado = new javax.swing.JComboBox<>(new String[]{"\u25cf Activo", "\u25cf Inactivo"});
        etiquetaId = new javax.swing.JLabel("ID: ---");
        etiquetaImagen = new javax.swing.JLabel("+ Agregar imagen");
        comboUnidad.setEditable(true);
        comboCategoria.setEditable(true);
        comboMarca.setEditable(true);

        getContentPane().setLayout(new java.awt.BorderLayout());
        getContentPane().add(jPanel1, java.awt.BorderLayout.CENTER);

        jButton6.addActionListener(this::jButton6ActionPerformed);

        pack();
    }

    private void jButton6ActionPerformed(java.awt.event.ActionEvent evt) {
        PRINCIPAL ventana = new PRINCIPAL();
        ventana.setVisible(true);
        this.dispose();
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ReflectiveOperationException | javax.swing.UnsupportedLookAndFeelException ex) {
            logger.log(java.util.logging.Level.SEVERE, null, ex);
        }

        java.awt.EventQueue.invokeLater(() -> new INVENTARIOS().setVisible(true));
    }

    private void configurarEventos() {
        jButton3.addActionListener(e -> alternarEstadoProducto());
        botonGuardar.addActionListener(e -> guardarProducto());
        botonLimpiar.addActionListener(e -> limpiarFormulario(false));
        botonAgregarProveedor.addActionListener(e -> mostrarDialogoProveedor(null));
        comboEstado.addActionListener(e -> actualizarBotonEstadoProducto());
        restablecerIndicadorEstado();
    }//configurarEventos

    // ===================== Diseño visual =====================
    private void configurarDisenoVisual() {
        jPanel1.setLayout(new java.awt.BorderLayout());
        jPanel1.setBackground(COLOR_FONDO);
        jPanel1.setBorder(javax.swing.BorderFactory.createEmptyBorder(18, 24, 18, 24));

        jPanel1.add(crearPanelEncabezado(), java.awt.BorderLayout.NORTH);
        jPanel1.add(crearPanelCuerpo(), java.awt.BorderLayout.CENTER);
        jPanel1.add(crearBarraInferior(), java.awt.BorderLayout.SOUTH);
    }//configurarDisenoVisual

    private javax.swing.JPanel crearPanelEncabezado() {
        javax.swing.JPanel panelHeader = new javax.swing.JPanel(new java.awt.BorderLayout(10, 0));
        panelHeader.setOpaque(false);
        panelHeader.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 0, 16, 0));

        javax.swing.JPanel panelIzquierda = new javax.swing.JPanel(new java.awt.BorderLayout(14, 0));
        panelIzquierda.setOpaque(false);
        panelIzquierda.add(crearInsigniaCuadrada(iconoCaja(java.awt.Color.WHITE), COLOR_AZUL, 46, 12),
                java.awt.BorderLayout.WEST);

        javax.swing.JPanel textoTitulo = new javax.swing.JPanel();
        textoTitulo.setOpaque(false);
        textoTitulo.setLayout(new javax.swing.BoxLayout(textoTitulo, javax.swing.BoxLayout.Y_AXIS));
        jLabel1.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 22));
        jLabel1.setForeground(COLOR_AZUL_OSCURO);
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        javax.swing.JLabel subtitulo = new javax.swing.JLabel("Gesti\u00f3n de productos y proveedores asociados");
        subtitulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        subtitulo.setForeground(COLOR_TEXTO_GRIS);
        textoTitulo.add(jLabel1);
        textoTitulo.add(subtitulo);
        panelIzquierda.add(textoTitulo, java.awt.BorderLayout.CENTER);
        panelHeader.add(panelIzquierda, java.awt.BorderLayout.WEST);

        javax.swing.JPanel panelDerecha = new javax.swing.JPanel();
        panelDerecha.setOpaque(false);
        panelDerecha.setLayout(new javax.swing.BoxLayout(panelDerecha, javax.swing.BoxLayout.Y_AXIS));

        javax.swing.JLabel breadcrumb = new javax.swing.JLabel(iconoInicio(COLOR_TEXTO_GRIS));
        breadcrumb.setText("Inicio > Inventarios > Producto");
        breadcrumb.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 11));
        breadcrumb.setForeground(COLOR_TEXTO_GRIS);
        breadcrumb.setIconTextGap(6);
        breadcrumb.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        breadcrumb.setAlignmentX(java.awt.Component.RIGHT_ALIGNMENT);

        panelDerecha.add(breadcrumb);
        panelHeader.add(panelDerecha, java.awt.BorderLayout.EAST);

        return panelHeader;
    }//crearPanelEncabezado

    private void actualizarColorPildoraEstado() {
        boolean activo = comboEstado.getSelectedIndex() == 0;
        comboEstado.setBackground(activo ? COLOR_VERDE_CLARO : COLOR_ROJO_CLARO);
        comboEstado.setForeground(activo ? COLOR_VERDE : COLOR_ROJO);
    }//actualizarColorPildoraEstado

    private javax.swing.JPanel crearPanelCuerpo() {
        javax.swing.JPanel panelCuerpo = new javax.swing.JPanel(new java.awt.GridBagLayout());
        panelCuerpo.setOpaque(false);
        java.awt.GridBagConstraints gc = new java.awt.GridBagConstraints();
        gc.gridy = 0;
        gc.fill = java.awt.GridBagConstraints.BOTH;
        gc.weighty = 1;

        gc.gridx = 0;
        gc.weightx = 0.52;
        gc.insets = new java.awt.Insets(0, 0, 0, 10);
        panelCuerpo.add(crearTarjetaProducto(), gc);

        gc.gridx = 1;
        gc.weightx = 0.48;
        gc.insets = new java.awt.Insets(0, 10, 0, 0);
        panelCuerpo.add(crearTarjetaProveedores(), gc);

        return panelCuerpo;
    }//crearPanelCuerpo

    private javax.swing.JPanel crearTarjetaProducto() {
        RoundedPanel tarjeta = new RoundedPanel(14, java.awt.Color.WHITE);
        tarjeta.setLayout(new java.awt.BorderLayout(0, 14));
        tarjeta.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO),
                javax.swing.BorderFactory.createEmptyBorder(18, 20, 18, 20)));

        tarjeta.add(crearTituloTarjeta("Informaci\u00f3n del Producto", iconoCaja(java.awt.Color.WHITE), COLOR_AZUL),
                java.awt.BorderLayout.NORTH);

        javax.swing.JPanel formulario = new javax.swing.JPanel(new java.awt.GridBagLayout());
        formulario.setOpaque(false);
        java.awt.GridBagConstraints gc = new java.awt.GridBagConstraints();
        gc.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gc.insets = new java.awt.Insets(6, 0, 6, 10);
        gc.weightx = 1;

        estilizarCampoTexto(jTextField1);
        estilizarCampoTexto(jTextField2);
        estilizarCampoCombo(comboUnidad);
        estilizarCampoTexto(jTextField5);
        estilizarCampoTexto(jTextField6);
        estilizarCampoCombo(comboCategoria);
        estilizarCampoCombo(comboMarca);
        estilizarCampoTexto(campoStockMinimo);

        gc.gridx = 0;
        gc.gridy = 0;
        gc.gridwidth = 1;
        formulario.add(crearCampoVertical("C\u00d3DIGO", jTextField1), gc);

        gc.gridx = 1;
        gc.gridy = 0;
        gc.gridwidth = 1;
        gc.insets = new java.awt.Insets(6, 0, 6, 0);
        formulario.add(crearRecuadroImagen(), gc);
        gc.insets = new java.awt.Insets(6, 0, 6, 10);

        gc.gridx = 0;
        gc.gridy = 1;
        gc.gridwidth = 2;
        formulario.add(crearCampoVertical("DESCRIPCI\u00d3N", jTextField2), gc);
        gc.gridwidth = 1;

        gc.gridx = 0;
        gc.gridy = 2;
        formulario.add(crearCampoVertical("CATEGOR\u00cdA", comboCategoria), gc);
        gc.gridx = 1;
        formulario.add(crearCampoVertical("MARCA", comboMarca), gc);

        gc.gridx = 0;
        gc.gridy = 3;
        formulario.add(crearCampoVertical("UNIDAD", comboUnidad), gc);
        gc.gridx = 1;
        formulario.add(crearCampoVertical("PRECIO UNITARIO", jTextField5), gc);

        gc.gridx = 0;
        gc.gridy = 4;
        formulario.add(crearCampoVertical("CANTIDAD EN INVENTARIO", jTextField6), gc);
        gc.gridx = 1;
        formulario.add(crearCampoVertical("STOCK M\u00cdNIMO", campoStockMinimo), gc);

        areaObservaciones.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        areaObservaciones.setRows(3);
        areaObservaciones.setLineWrap(true);
        areaObservaciones.setWrapStyleWord(true);
        areaObservaciones.setBorder(javax.swing.BorderFactory.createEmptyBorder(6, 8, 6, 8));
        javax.swing.JScrollPane scrollObservaciones = new javax.swing.JScrollPane(areaObservaciones);
        scrollObservaciones.setBorder(new BordeRedondeado(9, COLOR_BORDE_CAMPO, 1.2f));
        scrollObservaciones.setPreferredSize(new java.awt.Dimension(120, 70));

        gc.gridx = 0;
        gc.gridy = 5;
        gc.gridwidth = 2;
        gc.fill = java.awt.GridBagConstraints.BOTH;
        gc.weighty = 1;
        formulario.add(crearCampoVertical("OBSERVACIONES", scrollObservaciones), gc);

        tarjeta.add(formulario, java.awt.BorderLayout.CENTER);
        return tarjeta;
    }//crearTarjetaProducto

    private javax.swing.JPanel crearRecuadroImagen() {
        etiquetaImagen.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        etiquetaImagen.setVerticalAlignment(javax.swing.SwingConstants.CENTER);
        etiquetaImagen.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 11));
        etiquetaImagen.setForeground(COLOR_AZUL);
        etiquetaImagen.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        etiquetaImagen.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        etiquetaImagen.setOpaque(true);
        etiquetaImagen.setBackground(COLOR_AZUL_PALIDO);
        etiquetaImagen.setBorder(new BordeRedondeado(10, COLOR_BORDE_CAMPO, 1.2f));
        etiquetaImagen.setPreferredSize(new java.awt.Dimension(120, 88));
        etiquetaImagen.setCursor(java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR));
        etiquetaImagen.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                seleccionarImagen();
            }
        });

        javax.swing.JPanel contenedor = new javax.swing.JPanel(new java.awt.BorderLayout());
        contenedor.setOpaque(false);
        contenedor.add(etiquetaImagen, java.awt.BorderLayout.CENTER);
        return contenedor;
    }//crearRecuadroImagen

    private javax.swing.JPanel crearTarjetaProveedores() {
        RoundedPanel tarjeta = new RoundedPanel(14, java.awt.Color.WHITE);
        tarjeta.setLayout(new java.awt.BorderLayout(0, 12));
        tarjeta.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO),
                javax.swing.BorderFactory.createEmptyBorder(18, 20, 18, 20)));

        javax.swing.JPanel encabezado = new javax.swing.JPanel(new java.awt.BorderLayout());
        encabezado.setOpaque(false);
        encabezado.add(crearTituloTarjeta("Proveedores del Producto", iconoCamion(java.awt.Color.WHITE), COLOR_AZUL),
                java.awt.BorderLayout.WEST);
        prepararBoton(botonAgregarProveedor, "+ Agregar Proveedor", null, COLOR_AZUL, java.awt.Color.WHITE, null);
        javax.swing.JPanel contenedorBoton = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 0, 0));
        contenedorBoton.setOpaque(false);
        contenedorBoton.add(botonAgregarProveedor);
        encabezado.add(contenedorBoton, java.awt.BorderLayout.EAST);
        tarjeta.add(encabezado, java.awt.BorderLayout.NORTH);

        modeloProveedoresProducto = new javax.swing.table.DefaultTableModel(
                new Object[]{"PROVEEDOR", "C\u00d3DIGO PROV.", "PRECIO", "TIEMPO ENTREGA",
                    "\u00dALTIMA COMPRA", "PRED.", "", ""}, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false;
            }
        };
        tablaProveedoresProducto = new javax.swing.JTable(modeloProveedoresProducto);
        tablaProveedoresProducto.setRowHeight(30);
        tablaProveedoresProducto.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        tablaProveedoresProducto.getTableHeader().setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        tablaProveedoresProducto.getTableHeader().setForeground(COLOR_AZUL_OSCURO);
        tablaProveedoresProducto.setSelectionBackground(COLOR_AZUL_CLARO);
        tablaProveedoresProducto.setGridColor(COLOR_BORDE_CAMPO);
        tablaProveedoresProducto.setShowVerticalLines(false);

        tablaProveedoresProducto.getColumnModel().getColumn(4).setMinWidth(92);
        tablaProveedoresProducto.getColumnModel().getColumn(4).setPreferredWidth(105);
        tablaProveedoresProducto.getColumnModel().getColumn(4).setMaxWidth(120);
        tablaProveedoresProducto.getColumnModel().getColumn(COL_PREDETERMINADO).setPreferredWidth(46);
        tablaProveedoresProducto.getColumnModel().getColumn(COL_PREDETERMINADO).setMaxWidth(50);
        tablaProveedoresProducto.getColumnModel().getColumn(COL_EDITAR).setPreferredWidth(34);
        tablaProveedoresProducto.getColumnModel().getColumn(COL_EDITAR).setMaxWidth(36);
        tablaProveedoresProducto.getColumnModel().getColumn(COL_ELIMINAR).setPreferredWidth(34);
        tablaProveedoresProducto.getColumnModel().getColumn(COL_ELIMINAR).setMaxWidth(36);

        tablaProveedoresProducto.getColumnModel().getColumn(COL_PREDETERMINADO).setCellRenderer(
                (t, valor, seleccionado, foco, fila, columna) -> {
                    boolean predeterminado = Boolean.TRUE.equals(valor);
                    javax.swing.JLabel etiqueta = new javax.swing.JLabel(iconoRadio(predeterminado));
                    etiqueta.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
                    etiqueta.setOpaque(true);
                    etiqueta.setBackground(seleccionado ? COLOR_AZUL_CLARO : java.awt.Color.WHITE);
                    return etiqueta;
                });
        tablaProveedoresProducto.getColumnModel().getColumn(COL_EDITAR).setCellRenderer(
                (t, valor, seleccionado, foco, fila, columna) -> {
                    javax.swing.JLabel etiqueta = new javax.swing.JLabel(iconoEditar(COLOR_AZUL));
                    etiqueta.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
                    etiqueta.setOpaque(true);
                    etiqueta.setBackground(seleccionado ? COLOR_AZUL_CLARO : java.awt.Color.WHITE);
                    return etiqueta;
                });
        tablaProveedoresProducto.getColumnModel().getColumn(COL_ELIMINAR).setCellRenderer(
                (t, valor, seleccionado, foco, fila, columna) -> {
                    javax.swing.JLabel etiqueta = new javax.swing.JLabel(iconoBasura(COLOR_ROJO));
                    etiqueta.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
                    etiqueta.setOpaque(true);
                    etiqueta.setBackground(seleccionado ? COLOR_AZUL_CLARO : java.awt.Color.WHITE);
                    return etiqueta;
                });

        tablaProveedoresProducto.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                manejarClicTablaProveedores(evt);
            }
        });

        javax.swing.JScrollPane scroll = new javax.swing.JScrollPane(tablaProveedoresProducto);
        scroll.setBorder(javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO));
        tarjeta.add(scroll, java.awt.BorderLayout.CENTER);

        return tarjeta;
    }//crearTarjetaProveedores

    private void manejarClicTablaProveedores(java.awt.event.MouseEvent evt) {
        int fila = tablaProveedoresProducto.rowAtPoint(evt.getPoint());
        int columna = tablaProveedoresProducto.columnAtPoint(evt.getPoint());
        if (fila < 0 || columna < 0 || fila >= filasProveedorProducto.size()) {
            return;
        }
        if (columna == COL_PREDETERMINADO) {
            marcarProveedorPredeterminado(fila);
        } else if (columna == COL_EDITAR) {
            mostrarDialogoProveedor(fila);
        } else if (columna == COL_ELIMINAR) {
            eliminarProveedorFila(fila);
        }
    }//manejarClicTablaProveedores

    private javax.swing.JPanel crearBarraInferior() {
        javax.swing.JPanel barra = new javax.swing.JPanel(new java.awt.BorderLayout());
        barra.setOpaque(false);
        barra.setBorder(javax.swing.BorderFactory.createEmptyBorder(16, 0, 0, 0));

        prepararBoton(jButton6, "ATR\u00c1S", iconoChevronIzquierda(COLOR_AZUL_OSCURO), java.awt.Color.WHITE, COLOR_AZUL_OSCURO, COLOR_BORDE_CAMPO);
        javax.swing.JPanel panelIzquierda = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.LEFT, 0, 0));
        panelIzquierda.setOpaque(false);
        panelIzquierda.add(jButton6);
        barra.add(panelIzquierda, java.awt.BorderLayout.WEST);

        prepararBoton(botonLimpiar, "LIMPIAR", iconoLimpiar(COLOR_AZUL_OSCURO), java.awt.Color.WHITE, COLOR_AZUL_OSCURO, COLOR_BORDE_CAMPO);
        prepararBoton(botonGuardar, "GUARDAR", iconoGuardar(java.awt.Color.WHITE), COLOR_AZUL, java.awt.Color.WHITE, null);
        actualizarBotonEstadoProducto();

        javax.swing.JPanel panelDerecha = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 12, 0));
        panelDerecha.setOpaque(false);
        panelDerecha.add(botonLimpiar);
        panelDerecha.add(botonGuardar);
        panelDerecha.add(jButton3);
        barra.add(panelDerecha, java.awt.BorderLayout.EAST);

        return barra;
    }//crearBarraInferior

    private javax.swing.JPanel crearTituloTarjeta(String texto, javax.swing.Icon icono, java.awt.Color colorFondo) {
        javax.swing.JPanel panel = new javax.swing.JPanel(new java.awt.BorderLayout(10, 0));
        panel.setOpaque(false);
        panel.add(crearInsigniaCuadrada(icono, colorFondo, 34, 9), java.awt.BorderLayout.WEST);
        javax.swing.JLabel etiqueta = new javax.swing.JLabel(texto);
        etiqueta.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 15));
        etiqueta.setForeground(COLOR_AZUL_OSCURO);
        panel.add(etiqueta, java.awt.BorderLayout.CENTER);
        return panel;
    }//crearTituloTarjeta

    private javax.swing.JPanel crearCampoVertical(String etiqueta, javax.swing.JComponent campo) {
        javax.swing.JPanel contenedor = new javax.swing.JPanel(new java.awt.BorderLayout(0, 4));
        contenedor.setOpaque(false);
        javax.swing.JLabel label = new javax.swing.JLabel(etiqueta);
        label.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 10));
        label.setForeground(COLOR_TEXTO_GRIS);
        contenedor.add(label, java.awt.BorderLayout.NORTH);
        contenedor.add(campo, java.awt.BorderLayout.CENTER);
        return contenedor;
    }//crearCampoVertical

    private void estilizarCampoTexto(javax.swing.JTextField campo) {
        campo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        campo.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                new BordeRedondeado(9, COLOR_BORDE_CAMPO, 1.2f),
                javax.swing.BorderFactory.createEmptyBorder(6, 10, 6, 10)));
        campo.setPreferredSize(new java.awt.Dimension(120, 32));
    }//estilizarCampoTexto

    private void estilizarCampoCombo(javax.swing.JComboBox<String> combo) {
        combo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        combo.setPreferredSize(new java.awt.Dimension(120, 32));
    }//estilizarCampoCombo

    // ===================== Imagen del producto =====================
    private void seleccionarImagen() {
        javax.swing.JFileChooser selector = new javax.swing.JFileChooser();
        selector.setDialogTitle("Selecciona una imagen del producto");
        selector.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "Im\u00e1genes (PNG, JPG)", "png", "jpg", "jpeg"));
        int resultado = selector.showOpenDialog(this);
        if (resultado == javax.swing.JFileChooser.APPROVE_OPTION) {
            rutaImagenActual = selector.getSelectedFile().getAbsolutePath();
            actualizarPreviewImagen();
        }
    }//seleccionarImagen

    private void actualizarPreviewImagen() {
        if (rutaImagenActual == null || rutaImagenActual.isBlank()) {
            etiquetaImagen.setIcon(null);
            etiquetaImagen.setText("+ Agregar imagen");
            return;
        }
        File archivoImagen = new File(rutaImagenActual);
        if (!archivoImagen.exists()) {
            etiquetaImagen.setIcon(null);
            etiquetaImagen.setText("Imagen no disponible");
            return;
        }
        try {
            java.awt.image.BufferedImage imagen = javax.imageio.ImageIO.read(archivoImagen);
            if (imagen == null) {
                throw new IOException("Formato de imagen no soportado.");
            }
            java.awt.Image escalada = imagen.getScaledInstance(96, 72, java.awt.Image.SCALE_SMOOTH);
            etiquetaImagen.setIcon(new javax.swing.ImageIcon(escalada));
            etiquetaImagen.setText(null);
        } catch (IOException ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo cargar la imagen del producto", ex);
            etiquetaImagen.setIcon(null);
            etiquetaImagen.setText("Imagen no disponible");
        }
    }//actualizarPreviewImagen

    // ===================== Proveedores del producto =====================
    private void refrescarTablaProveedoresProducto(String codigoProducto) {
        modeloProveedoresProducto.setRowCount(0);
        filasProveedorProducto.clear();
        if (codigoProducto == null || codigoProducto.isBlank()) {
            return;
        }
        java.util.Map<String, String> ultimasCompras = cargarUltimasComprasPorProveedor();
        java.util.List<String[]> filas = productoProveedorRepository.obtenerProveedoresDeProducto(codigoProducto);
        for (String[] datos : filas) {
            String idProveedor = datos[1];
            String nombreProveedor = proveedorRepository.obtenerNombrePorId(idProveedor);
            boolean predeterminado = "SI".equalsIgnoreCase(datos[5]);
            modeloProveedoresProducto.addRow(new Object[]{
                nombreProveedor, datos[2], datos[3], datos[4],
                ultimasCompras.getOrDefault(idProveedor.toUpperCase(java.util.Locale.ROOT), "SIN COMPRAS"),
                predeterminado, "", ""
            });
            filasProveedorProducto.add(datos);
        }
    }//refrescarTablaProveedoresProducto

    private java.util.Map<String, String> cargarUltimasComprasPorProveedor() {
        java.util.Map<String, java.time.LocalDate> fechas = new java.util.HashMap<>();
        File archivoOrdenes = com.mycompany.ocxrst.config.AppConfig.registroCFile();
        if (!archivoOrdenes.exists() || archivoOrdenes.length() == 0) {
            return new java.util.HashMap<>();
        }
        try (java.io.FileInputStream fis = new java.io.FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row fila : hoja) {
                if (fila.getRowNum() == 0) {
                    continue;
                }
                String idProveedor = leerCelda(fila, 4).toUpperCase(java.util.Locale.ROOT);
                java.time.LocalDate fecha = leerFechaOrden(fila.getCell(1));
                if (idProveedor.isEmpty() || fecha == null) {
                    continue;
                }
                java.time.LocalDate fechaAnterior = fechas.get(idProveedor);
                if (fechaAnterior == null || fecha.isAfter(fechaAnterior)) {
                    fechas.put(idProveedor, fecha);
                }
            }
        } catch (IOException ex) {
            logger.log(java.util.logging.Level.WARNING,
                    "No se pudieron cargar las últimas compras por proveedor.", ex);
            return new java.util.HashMap<>();
        }

        java.util.Map<String, String> resultado = new java.util.HashMap<>();
        java.time.format.DateTimeFormatter formato = java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy");
        for (java.util.Map.Entry<String, java.time.LocalDate> entrada : fechas.entrySet()) {
            resultado.put(entrada.getKey(), entrada.getValue().format(formato));
        }
        return resultado;
    }//cargarUltimasComprasPorProveedor

    private java.time.LocalDate leerFechaOrden(Cell celda) {
        if (celda == null) {
            return null;
        }
        if (DateUtil.isCellDateFormatted(celda)) {
            return celda.getDateCellValue().toInstant()
                    .atZone(java.time.ZoneId.systemDefault()).toLocalDate();
        }
        String texto = new org.apache.poi.ss.usermodel.DataFormatter().formatCellValue(celda).trim();
        for (String patron : new String[]{"dd/MM/yyyy", "d/M/yyyy"}) {
            try {
                return java.time.LocalDate.parse(texto,
                        java.time.format.DateTimeFormatter.ofPattern(patron));
            } catch (java.time.format.DateTimeParseException ex) {
                // Intenta el siguiente formato admitido.
            }
        }
        return null;
    }//leerFechaOrden

    private void marcarProveedorPredeterminado(int filaModelo) {
        String[] datos = filasProveedorProducto.get(filaModelo);
        int numeroFilaExcel = Integer.parseInt(datos[0]);
        String codigoProducto = jTextField1.getText().trim();
        productoProveedorRepository.marcarProveedorPredeterminado(numeroFilaExcel, codigoProducto);
        refrescarTablaProveedoresProducto(codigoProducto);
    }//marcarProveedorPredeterminado

    private void eliminarProveedorFila(int filaModelo) {
        int confirmacion = JOptionPane.showConfirmDialog(this,
                "\u00bfEliminar este proveedor del producto?", "Confirmar",
                JOptionPane.YES_NO_OPTION);
        if (confirmacion != JOptionPane.YES_OPTION) {
            return;
        }
        String[] datos = filasProveedorProducto.get(filaModelo);
        productoProveedorRepository.eliminarProveedorDeProducto(Integer.parseInt(datos[0]));
        refrescarTablaProveedoresProducto(jTextField1.getText().trim());
    }//eliminarProveedorFila

    private void mostrarDialogoProveedor(Integer filaModeloParaEditar) {
        String codigoProducto = jTextField1.getText().trim();
        if (codigoProducto.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Captura y guarda el C\u00d3DIGO del producto antes de asociar proveedores.");
            return;
        }

        java.util.List<String[]> proveedoresActivos = proveedorRepository.obtenerProveedoresActivos();
        if (proveedoresActivos.isEmpty()) {
            JOptionPane.showMessageDialog(this, "No hay proveedores activos registrados. Registra uno en la pantalla de PROVEEDORES.");
            return;
        }

        javax.swing.JDialog dialogo = new javax.swing.JDialog(this,
                filaModeloParaEditar == null ? "Agregar Proveedor" : "Editar Proveedor", true);
        dialogo.setLayout(new java.awt.GridBagLayout());
        java.awt.GridBagConstraints gc = new java.awt.GridBagConstraints();
        gc.insets = new java.awt.Insets(8, 12, 8, 12);
        gc.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gc.gridx = 0;
        gc.gridy = 0;
        gc.gridwidth = 2;

        String[] opcionesProveedor = new String[proveedoresActivos.size()];
        for (int i = 0; i < proveedoresActivos.size(); i++) {
            opcionesProveedor[i] = proveedoresActivos.get(i)[0] + " - " + proveedoresActivos.get(i)[1];
        }
        javax.swing.JComboBox<String> comboProveedor = new javax.swing.JComboBox<>(opcionesProveedor);
        javax.swing.JTextField campoCodigoProveedor = new javax.swing.JTextField();
        javax.swing.JTextField campoPrecio = new javax.swing.JTextField();
        javax.swing.JTextField campoTiempo = new javax.swing.JTextField();
        javax.swing.JCheckBox checkPredeterminado = new javax.swing.JCheckBox("Marcar como predeterminado");

        String[] datosExistentes = filaModeloParaEditar == null ? null : filasProveedorProducto.get(filaModeloParaEditar);
        if (datosExistentes != null) {
            for (int i = 0; i < proveedoresActivos.size(); i++) {
                if (proveedoresActivos.get(i)[0].equalsIgnoreCase(datosExistentes[1])) {
                    comboProveedor.setSelectedIndex(i);
                    break;
                }
            }
            campoCodigoProveedor.setText(datosExistentes[2]);
            campoPrecio.setText(datosExistentes[3]);
            campoTiempo.setText(datosExistentes[4]);
            checkPredeterminado.setSelected("SI".equalsIgnoreCase(datosExistentes[5]));
        }

        dialogo.add(new javax.swing.JLabel("Proveedor:"), gc);
        gc.gridy++;
        dialogo.add(comboProveedor, gc);
        gc.gridy++;
        dialogo.add(new javax.swing.JLabel("C\u00f3digo del Proveedor:"), gc);
        gc.gridy++;
        dialogo.add(campoCodigoProveedor, gc);
        gc.gridy++;
        dialogo.add(new javax.swing.JLabel("Precio:"), gc);
        gc.gridy++;
        dialogo.add(campoPrecio, gc);
        gc.gridy++;
        dialogo.add(new javax.swing.JLabel("Tiempo de entrega:"), gc);
        gc.gridy++;
        dialogo.add(campoTiempo, gc);
        gc.gridy++;
        dialogo.add(checkPredeterminado, gc);
        gc.gridy++;

        javax.swing.JButton botonGuardarDialogo = new javax.swing.JButton("Guardar");
        javax.swing.JButton botonCancelarDialogo = new javax.swing.JButton("Cancelar");
        javax.swing.JPanel panelBotones = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT));
        panelBotones.add(botonCancelarDialogo);
        panelBotones.add(botonGuardarDialogo);
        dialogo.add(panelBotones, gc);

        botonCancelarDialogo.addActionListener(e -> dialogo.dispose());
        botonGuardarDialogo.addActionListener(e -> {
            if (campoCodigoProveedor.getText().trim().isEmpty() || campoPrecio.getText().trim().isEmpty()
                    || campoTiempo.getText().trim().isEmpty()) {
                JOptionPane.showMessageDialog(dialogo, "Todos los campos son obligatorios.");
                return;
            }
            String seleccion = (String) comboProveedor.getSelectedItem();
            String idProveedor = seleccion.split(" - ", 2)[0].trim();
            String codigoProv = campoCodigoProveedor.getText().trim();
            String precio = campoPrecio.getText().trim();
            String tiempo = campoTiempo.getText().trim();
            boolean predeterminado = checkPredeterminado.isSelected();

            if (filaModeloParaEditar == null) {
                productoProveedorRepository.agregarProveedorAProducto(codigoProducto, idProveedor, codigoProv, precio, tiempo, predeterminado);
            } else {
                int numeroFilaExcel = Integer.parseInt(datosExistentes[0]);
                productoProveedorRepository.actualizarProveedorDeProducto(numeroFilaExcel, codigoProducto, idProveedor, codigoProv, precio, tiempo, predeterminado);
            }
            refrescarTablaProveedoresProducto(codigoProducto);
            dialogo.dispose();
        });

        dialogo.pack();
        dialogo.setLocationRelativeTo(this);
        dialogo.setVisible(true);
    }//mostrarDialogoProveedor

    // ===================== Persistencia del producto =====================
    private void asegurarArchivoInventarios() {
        try {
            inventarioRepository.asegurarArchivoInventarios();
        } catch (IllegalStateException ex) {
            mostrarError("No se pudo preparar el archivo de inventarios.", ex);
        }
    }//asegurarArchivoInventarios

    private void guardarProducto() {
        if (!validarCamposBase()) {
            return;
        }

        String codigo = jTextField1.getText().trim();

        try {
            File archivo = obtenerArchivoInventarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaInventarios(libro);
                asegurarEncabezados(hoja);

                int filaExistente = buscarFilaPorCodigo(hoja, codigo);
                boolean esEdicion = filaExistente >= 0;
                int filaDestino;

                if (esEdicion) {
                    filaDestino = filaExistente;
                } else {
                    filaDestino = primeraFilaDisponible(hoja);
                }

                Row fila = hoja.getRow(filaDestino);
                if (fila == null) {
                    fila = hoja.createRow(filaDestino);
                }

                guardarDatosEnFila(fila);
                guardarLibro(libro, archivo);
                filaInventarioActual = filaDestino;
                etiquetaId.setText("ID: " + codigo);

                JOptionPane.showMessageDialog(this, esEdicion
                        ? "Producto actualizado correctamente."
                        : "Producto creado.\nC\u00d3DIGO: " + codigo);
            }
            refrescarTablaProveedoresProducto(codigo);
        } catch (IOException ex) {
            mostrarError("No se pudo guardar el producto.", ex);
        }
        cargarCacheProductos();
    }//guardarProducto

    private void alternarEstadoProducto() {
        String codigo = jTextField1.getText().trim();
        if (codigo.isEmpty() || filaInventarioActual < 0) {
            JOptionPane.showMessageDialog(this, "Selecciona un producto de la lista para cambiar su estado.");
            return;
        }

        String estadoActual = "1".equals(normalizarEstado(
                comboEstado.getSelectedIndex() == 0 ? "1" : "0")) ? "1" : "0";
        String estadoNuevo = "1".equals(estadoActual) ? "0" : "1";
        String accion = "1".equals(estadoNuevo) ? "activar" : "desactivar";
        int confirmacion = JOptionPane.showConfirmDialog(this,
                "\u00bfDeseas " + accion + " el producto " + codigo + "?",
                "Confirmar cambio de estado", JOptionPane.YES_NO_OPTION);
        if (confirmacion != JOptionPane.YES_OPTION) {
            return;
        }

        try {
            File archivo = obtenerArchivoInventarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaInventarios(libro);
                int filaIndex = buscarFilaPorCodigo(hoja, codigo);
                if (filaIndex < 0) {
                    JOptionPane.showMessageDialog(this, "El producto ya no existe.");
                    return;
                }
                Row fila = hoja.getRow(filaIndex);
                if (fila == null) {
                    JOptionPane.showMessageDialog(this, "El registro encontrado está vacío.");
                    return;
                }
                escribirCelda(fila, 5, estadoNuevo);
                guardarLibro(libro, archivo);
            }
            comboEstado.setSelectedIndex("1".equals(estadoNuevo) ? 0 : 1);
            actualizarBotonEstadoProducto();
            JOptionPane.showMessageDialog(this, "Producto " + ("1".equals(estadoNuevo)
                    ? "activado" : "desactivado") + " correctamente.");
            cargarCacheProductos();
        } catch (IOException ex) {
            mostrarError("No se pudo cambiar el estado del producto.", ex);
        }
    }//alternarEstadoProducto

    // ===================== Autocompletado del campo C\u00d3DIGO =====================
    private void configurarAutocompletadoCodigo() {
        jTextField1.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
            @Override
            public void insertUpdate(javax.swing.event.DocumentEvent e) {
                filtrarProductos();
            }

            @Override
            public void removeUpdate(javax.swing.event.DocumentEvent e) {
                filtrarProductos();
            }

            @Override
            public void changedUpdate(javax.swing.event.DocumentEvent e) {
                filtrarProductos();
            }
        });

        jTextField1.addFocusListener(new java.awt.event.FocusAdapter() {
            @Override
            public void focusLost(java.awt.event.FocusEvent e) {
                if (!e.isTemporary()) {
                    ocultarListaProductos();
                }
            }
        });
    }//configurarAutocompletadoCodigo

    private void cargarCacheProductos() {
        cacheProductos.clear();
        try {
            File archivo = obtenerArchivoInventarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaInventarios(libro);
                for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
                    Row row = hoja.getRow(fila);
                    if (row == null) {
                        continue;
                    }
                    String codigo = leerCelda(row, 0);
                    if (!codigo.isEmpty()) {
                        cacheProductos.add(new String[]{codigo, leerCelda(row, 1)});
                    }
                }
            }
        } catch (IOException ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo cargar el cach\u00e9 de productos", ex);
        }
    }//cargarCacheProductos

    private void filtrarProductos() {
        if (seleccionandoProducto) {
            return;
        }
        String texto = jTextField1.getText().trim().toLowerCase();
        ocultarListaProductos();
        if (texto.isEmpty()) {
            return;
        }

        int contador = 0;
        for (String[] producto : cacheProductos) {
            if (producto[0].toLowerCase().contains(texto) || producto[1].toLowerCase().contains(texto)) {
                String etiqueta = producto[0] + "  |  " + producto[1];
                javax.swing.JMenuItem item = new javax.swing.JMenuItem(etiqueta);
                final String codigoSeleccionado = producto[0];
                item.addActionListener(ev -> seleccionarProductoDesdeLista(codigoSeleccionado));
                popupProductos.add(item);
                if (++contador == 10) {
                    break; // m\u00e1ximo 10 sugerencias
                }
            }
        }
        if (contador > 0) {
            popupProductos.show(jTextField1, 0, jTextField1.getHeight());
            jTextField1.requestFocusInWindow();
        }
    }//filtrarProductos

    private void seleccionarProductoDesdeLista(String codigo) {
        seleccionandoProducto = true;
        try {
            jTextField1.setText(codigo);
        } finally {
            seleccionandoProducto = false;
            ocultarListaProductos();
        }
        cargarProductoPorCodigo(codigo);
        javax.swing.SwingUtilities.invokeLater(this::ocultarListaProductos);
    }//seleccionarProductoDesdeLista

    private void ocultarListaProductos() {
        popupProductos.setVisible(false);
        popupProductos.removeAll();
    }

    private void cargarProductoPorCodigo(String codigo) {
        try {
            File archivo = obtenerArchivoInventarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaInventarios(libro);
                int filaIndex = buscarFilaPorCodigo(hoja, codigo);
                if (filaIndex < 0) {
                    return;
                }
                Row fila = hoja.getRow(filaIndex);
                if (fila == null) {
                    return;
                }
                cargarFilaEnFormulario(fila);
                filaInventarioActual = filaIndex;
                etiquetaId.setText("ID: " + codigo);
                refrescarTablaProveedoresProducto(codigo);
            }
        } catch (IOException ex) {
            mostrarError("No se pudo cargar el producto seleccionado.", ex);
        }
    }//cargarProductoPorCodigo

    private void buscarInventario() {
        String codigo = jTextField1.getText().trim();
        String descripcion = jTextField2.getText().trim();

        if (codigo.isEmpty() && descripcion.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Ingresa C\u00d3DIGO o DESCRIPCI\u00d3N para buscar.");
            return;
        }

        try {
            File archivo = obtenerArchivoInventarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaInventarios(libro);
                asegurarEncabezados(hoja);

                int filaIndex = -1;
                if (!codigo.isEmpty()) {
                    filaIndex = buscarFilaPorCodigo(hoja, codigo);
                }
                if (filaIndex < 0 && !descripcion.isEmpty()) {
                    filaIndex = buscarFilaPorDescripcion(hoja, descripcion);
                }
                if (filaIndex < 0) {
                    filaInventarioActual = -1;
                    JOptionPane.showMessageDialog(this, "No se encontr\u00f3 un producto con los datos capturados.");
                    return;
                }

                Row fila = hoja.getRow(filaIndex);
                if (fila == null) {
                    filaInventarioActual = -1;
                    JOptionPane.showMessageDialog(this, "El registro encontrado est\u00e1 vac\u00edo.");
                    return;
                }

                cargarFilaEnFormulario(fila);
                filaInventarioActual = filaIndex;
                String codigoEncontrado = leerCelda(fila, 0);
                etiquetaId.setText("ID: " + codigoEncontrado);
                refrescarTablaProveedoresProducto(codigoEncontrado);
                JOptionPane.showMessageDialog(this, "Producto encontrado.\nC\u00d3DIGO: " + codigoEncontrado);
            }
        } catch (IOException ex) {
            mostrarError("No se pudo buscar el producto.", ex);
        }
    }//buscarInventario

    private boolean validarCamposBase() {
        if (jTextField1.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "C\u00d3DIGO es obligatorio.");
            return false;
        }
        if (jTextField2.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "DESCRIPCI\u00d3N es obligatoria.");
            return false;
        }
        return true;
    }//validarCamposBase

    private void guardarDatosEnFila(Row fila) {
        escribirCelda(fila, 0, jTextField1.getText().trim());
        escribirCelda(fila, 1, jTextField2.getText().trim());
        escribirCelda(fila, 2, obtenerTextoCombo(comboUnidad));
        escribirCelda(fila, 3, jTextField5.getText().trim());
        escribirCelda(fila, 4, jTextField6.getText().trim());
        escribirCelda(fila, 5, comboEstado.getSelectedIndex() == 0 ? "1" : "0");
        escribirCelda(fila, 6, obtenerTextoCombo(comboCategoria));
        escribirCelda(fila, 7, obtenerTextoCombo(comboMarca));
        escribirCelda(fila, 8, campoStockMinimo.getText().trim());
        escribirCelda(fila, 9, areaObservaciones.getText().trim());
        escribirCelda(fila, 10, rutaImagenActual == null ? "" : rutaImagenActual.trim());
    }//guardarDatosEnFila

    private String obtenerTextoCombo(javax.swing.JComboBox<String> combo) {
        Object seleccionado = combo.getEditor().getItem();
        return seleccionado == null ? "" : seleccionado.toString().trim();
    }//obtenerTextoCombo

    private void cargarFilaEnFormulario(Row fila) {
        jTextField1.setText(leerCelda(fila, 0));
        jTextField2.setText(leerCelda(fila, 1));
        asignarValorCombo(comboUnidad, leerCelda(fila, 2));
        jTextField5.setText(leerCelda(fila, 3));
        jTextField6.setText(leerCelda(fila, 4));
        String estado = normalizarEstado(leerCelda(fila, 5));
        comboEstado.setSelectedIndex("1".equals(estado) ? 0 : 1);
        actualizarBotonEstadoProducto();
        asignarValorCombo(comboCategoria, leerCelda(fila, 6));
        asignarValorCombo(comboMarca, leerCelda(fila, 7));
        campoStockMinimo.setText(leerCelda(fila, 8));
        areaObservaciones.setText(leerCelda(fila, 9));
        rutaImagenActual = leerCelda(fila, 10);
        actualizarPreviewImagen();
    }//cargarFilaEnFormulario

    private void asignarValorCombo(javax.swing.JComboBox<String> combo, String valor) {
        if (valor == null || valor.isBlank()) {
            combo.setSelectedItem("");
            return;
        }
        for (int i = 0; i < combo.getItemCount(); i++) {
            if (valor.equalsIgnoreCase(combo.getItemAt(i))) {
                combo.setSelectedIndex(i);
                return;
            }
        }
        combo.insertItemAt(valor, 0);
        combo.setSelectedIndex(0);
    }//asignarValorCombo

    private void limpiarFormulario(boolean conservarCodigo) {
        if (!conservarCodigo) {
            jTextField1.setText("");
        }
        jTextField2.setText("");
        comboUnidad.setSelectedItem("");
        jTextField5.setText("");
        jTextField6.setText("");
        comboCategoria.setSelectedItem("");
        comboMarca.setSelectedItem("");
        campoStockMinimo.setText("");
        areaObservaciones.setText("");
        rutaImagenActual = "";
        actualizarPreviewImagen();
        comboEstado.setSelectedIndex(0);
        filaInventarioActual = -1;
        etiquetaId.setText("ID: ---");
        modeloProveedoresProducto.setRowCount(0);
        filasProveedorProducto.clear();
        restablecerIndicadorEstado();
    }//limpiarFormulario

    private int buscarFilaPorCodigo(Sheet hoja, String codigo) {
        for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
            Row row = hoja.getRow(fila);
            if (row == null) {
                continue;
            }
            if (codigo.equalsIgnoreCase(leerCelda(row, 0))) {
                return fila;
            }
        }
        return -1;
    }//buscarFilaPorCodigo

    private int buscarFilaPorDescripcion(Sheet hoja, String descripcion) {
        for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
            Row row = hoja.getRow(fila);
            if (row == null) {
                continue;
            }
            if (descripcion.equalsIgnoreCase(leerCelda(row, 1))) {
                return fila;
            }
        }
        return -1;
    }//buscarFilaPorDescripcion

    private int primeraFilaDisponible(Sheet hoja) {
        for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
            Row row = hoja.getRow(fila);
            if (row == null || leerCelda(row, 0).isEmpty()) {
                return fila;
            }
        }
        return hoja.getLastRowNum() + 1;
    }//primeraFilaDisponible

    private boolean asegurarEncabezados(Sheet hoja) {
        return inventarioRepository.asegurarEncabezados(hoja);
    }//asegurarEncabezados

    private Sheet obtenerHojaInventarios(Workbook libro) {
        return inventarioRepository.obtenerHojaInventarios(libro);
    }//obtenerHojaInventarios

    private Workbook cargarLibro(File archivo) throws IOException {
        return inventarioRepository.cargarLibro(archivo);
    }//cargarLibro

    private void guardarLibro(Workbook libro, File archivo) throws IOException {
        inventarioRepository.guardarLibro(libro, archivo);
    }//guardarLibro

    private File obtenerArchivoInventarios() {
        return inventarioRepository.obtenerArchivoInventarios();
    }//obtenerArchivoInventarios

    private String leerCelda(Row fila, int columna) {
        return inventarioRepository.leerCelda(fila, columna);
    }//leerCelda

    private void escribirCelda(Row fila, int columna, String valor) {
        inventarioRepository.escribirCelda(fila, columna, valor);
    }//escribirCelda

    private String normalizarEstado(String valor) {
        return InventarioService.normalizarEstado(valor);
    }//normalizarEstado

    private void restablecerIndicadorEstado() {
        etiquetaId.setText("ID: ---");
        comboEstado.setSelectedIndex(0);
        actualizarBotonEstadoProducto();
    }

    private void actualizarBotonEstadoProducto() {
        boolean seleccionado = filaInventarioActual >= 0
                && jTextField1 != null
                && !jTextField1.getText().trim().isEmpty();
        boolean activo = seleccionado && comboEstado != null && comboEstado.getSelectedIndex() == 0;
        prepararBoton(jButton3, seleccionado
                ? (activo ? "DESACTIVAR" : "ACTIVAR")
                : "ACTIVAR/DESACTIVAR",
                iconoPower(java.awt.Color.WHITE),
                seleccionado ? (activo ? COLOR_ROJO : COLOR_VERDE) : new java.awt.Color(120, 120, 120),
                java.awt.Color.WHITE, null);
        jButton3.setEnabled(seleccionado);
    }//restablecerIndicadorEstado

    private void mostrarError(String mensajeUsuario, Exception ex) {
        logger.log(java.util.logging.Level.SEVERE, mensajeUsuario, ex);
        JOptionPane.showMessageDialog(this, mensajeUsuario + "\nDetalle t\u00e9cnico: " + ex.getMessage());
    }//mostrarError

    private void prepararBoton(javax.swing.JButton boton, String texto, javax.swing.Icon icono,
            java.awt.Color fondo, java.awt.Color textoColor, java.awt.Color colorBorde) {
        boton.setUI(new BotonPlanoUI());
        boton.setText(texto);
        boton.setIcon(icono);
        boton.setIconTextGap(8);
        boton.setBackground(fondo);
        boton.setForeground(textoColor);
        boton.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        boton.setFocusPainted(false);
        boton.setContentAreaFilled(false);
        boton.setOpaque(false);
        boton.setRolloverEnabled(true);
        boton.setBorderPainted(colorBorde != null);
        boton.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                new BordeRedondeado(9, colorBorde, 1.4f),
                javax.swing.BorderFactory.createEmptyBorder(9, 16, 9, 16)));
        boton.setCursor(java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR));
    }//prepararBoton

    private static java.awt.Color mezclarColor(java.awt.Color base, java.awt.Color destino, float proporcion) {
        int r = base.getRed() + Math.round((destino.getRed() - base.getRed()) * proporcion);
        int g = base.getGreen() + Math.round((destino.getGreen() - base.getGreen()) * proporcion);
        int b = base.getBlue() + Math.round((destino.getBlue() - base.getBlue()) * proporcion);
        return new java.awt.Color(
                Math.max(0, Math.min(255, r)),
                Math.max(0, Math.min(255, g)),
                Math.max(0, Math.min(255, b)));
    }//mezclarColor

    /**
     * ButtonUI plano que sustituye el bisel nativo de Windows por un relleno
     * redondeado propio (mismo patr\u00f3n que PROVEEDORES.java, para mantener la
     * consistencia visual entre pantallas).
     */
    private static final class BotonPlanoUI extends javax.swing.plaf.basic.BasicButtonUI {
        @Override
        public void paint(java.awt.Graphics g, javax.swing.JComponent c) {
            javax.swing.AbstractButton boton = (javax.swing.AbstractButton) c;
            javax.swing.ButtonModel modelo = boton.getModel();
            java.awt.Color fondo = boton.getBackground();
            if (modelo.isPressed()) {
                fondo = mezclarColor(fondo, java.awt.Color.BLACK, 0.16f);
            } else if (modelo.isRollover()) {
                fondo = mezclarColor(fondo, java.awt.Color.BLACK, 0.07f);
            }

            java.awt.Graphics2D g2 = (java.awt.Graphics2D) g.create();
            g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                    java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setColor(fondo);
            g2.fillRoundRect(0, 0, c.getWidth() - 1, c.getHeight() - 1, 9, 9);
            g2.dispose();

            super.paint(g, c);
        }
    }

    private static final class BordeRedondeado extends javax.swing.border.AbstractBorder {
        private final int radio;
        private final java.awt.Color color;
        private final float grosor;

        BordeRedondeado(int radio, java.awt.Color color, float grosor) {
            this.radio = radio;
            this.color = color;
            this.grosor = grosor;
        }

        @Override
        public void paintBorder(java.awt.Component c, java.awt.Graphics g, int x, int y, int w, int h) {
            if (color == null) {
                return;
            }
            java.awt.Graphics2D g2 = (java.awt.Graphics2D) g.create();
            g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                    java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setColor(color);
            g2.setStroke(new java.awt.BasicStroke(grosor));
            int mitad = Math.round(grosor / 2);
            int grosorEntero = Math.round(grosor);
            g2.drawRoundRect(x + mitad, y + mitad, w - grosorEntero - 1, h - grosorEntero - 1, radio, radio);
            g2.dispose();
        }

        @Override
        public java.awt.Insets getBorderInsets(java.awt.Component c) {
            return new java.awt.Insets(2, 2, 2, 2);
        }

        @Override
        public boolean isBorderOpaque() {
            return false;
        }
    }

    private static final class RoundedPanel extends javax.swing.JPanel {
        private final int radio;
        private final java.awt.Color color;

        RoundedPanel(int radio, java.awt.Color color) {
            this.radio = radio;
            this.color = color;
            setOpaque(false);
        }

        @Override
        protected void paintComponent(java.awt.Graphics g) {
            java.awt.Graphics2D g2 = (java.awt.Graphics2D) g.create();
            g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                    java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setColor(new java.awt.Color(180, 196, 213, 90));
            g2.fillRoundRect(1, 2, getWidth() - 2, getHeight() - 2, radio, radio);
            g2.setColor(color);
            g2.fillRoundRect(0, 0, getWidth() - 2, getHeight() - 3, radio, radio);
            g2.dispose();
            super.paintComponent(g);
        }
    }

    @FunctionalInterface
    private interface DibujoIcono {
        void dibujar(java.awt.Graphics2D g2, int tamano);
    }

    private static final class VectorIcon implements javax.swing.Icon {
        private final int tamano;
        private final java.awt.Color color;
        private final DibujoIcono dibujo;

        VectorIcon(int tamano, java.awt.Color color, DibujoIcono dibujo) {
            this.tamano = tamano;
            this.color = color;
            this.dibujo = dibujo;
        }

        @Override
        public void paintIcon(java.awt.Component c, java.awt.Graphics g, int x, int y) {
            java.awt.Graphics2D g2 = (java.awt.Graphics2D) g.create();
            g2.translate(x, y);
            g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                    java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setColor(color);
            g2.setStroke(new java.awt.BasicStroke(1.6f,
                    java.awt.BasicStroke.CAP_ROUND, java.awt.BasicStroke.JOIN_ROUND));
            dibujo.dibujar(g2, tamano);
            g2.dispose();
        }

        @Override
        public int getIconWidth() {
            return tamano;
        }

        @Override
        public int getIconHeight() {
            return tamano;
        }
    }

    private static javax.swing.JComponent crearInsigniaCuadrada(javax.swing.Icon icono, java.awt.Color colorFondo, int tamano, int radio) {
        javax.swing.JLabel insignia = new javax.swing.JLabel(icono) {
            @Override
            protected void paintComponent(java.awt.Graphics g) {
                java.awt.Graphics2D g2 = (java.awt.Graphics2D) g.create();
                g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING, java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
                g2.setColor(colorFondo);
                g2.fillRoundRect(0, 0, getWidth() - 1, getHeight() - 1, radio, radio);
                g2.dispose();
                super.paintComponent(g);
            }
        };
        insignia.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        insignia.setVerticalAlignment(javax.swing.SwingConstants.CENTER);
        insignia.setPreferredSize(new java.awt.Dimension(tamano, tamano));
        insignia.setOpaque(false);
        return insignia;
    }//crearInsigniaCuadrada

    // ===================== Iconos vectoriales (Java2D) =====================
    private static javax.swing.Icon iconoCaja(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawRect(2, 5, t - 4, t - 7);
            g2.drawLine(2, 5, t / 2, 1);
            g2.drawLine(t - 2, 5, t / 2, 1);
            g2.drawLine(t / 2, 1, t / 2, t - 2);
        });
    }//iconoCaja

    private static javax.swing.Icon iconoCamion(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawRect(1, 5, 9, 7);
            g2.drawRect(10, 8, 6, 4);
            g2.drawOval(3, t - 4, 3, 3);
            g2.drawOval(11, t - 4, 3, 3);
        });
    }//iconoCamion

    private static javax.swing.Icon iconoBuscar(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawOval(1, 1, t - 6, t - 6);
            g2.drawLine(t - 6, t - 6, t - 1, t - 1);
        });
    }//iconoBuscar

    private static javax.swing.Icon iconoLimpiar(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawLine(2, t - 2, t - 6, 4);
            g2.drawRect(t - 8, 2, 6, 6);
            g2.drawLine(2, t - 2, t - 2, t - 2);
        });
    }//iconoLimpiar

    private static javax.swing.Icon iconoGuardar(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawRoundRect(1, 1, t - 2, t - 2, 2, 2);
            g2.drawRect(4, 1, t - 8, 5);
            g2.drawRect(4, t - 6, t - 8, 5);
        });
    }//iconoGuardar

    private static javax.swing.Icon iconoBasura(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawLine(3, 4, t - 3, 4);
            g2.drawRect(4, 4, t - 8, t - 5);
            g2.drawLine(6, 1, t - 6, 1);
            g2.drawLine(6, 7, 6, t - 4);
            g2.drawLine(t / 2, 7, t / 2, t - 4);
            g2.drawLine(t - 6, 7, t - 6, t - 4);
        });
    }//iconoBasura

    private static javax.swing.Icon iconoPower(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawArc(2, 2, t - 4, t - 4, -45, 270);
            g2.drawLine(t / 2, 1, t / 2, t / 2);
        });
    }

    private static javax.swing.Icon iconoEditar(java.awt.Color color) {
        return new VectorIcon(15, color, (g2, t) -> {
            java.awt.geom.GeneralPath lapiz = new java.awt.geom.GeneralPath();
            lapiz.moveTo(3, t - 2);
            lapiz.lineTo(2, t - 4);
            lapiz.lineTo(t - 5, 2);
            lapiz.lineTo(t - 2, 5);
            lapiz.lineTo(5, t - 1);
            lapiz.closePath();
            g2.draw(lapiz);
        });
    }//iconoEditar

    private static javax.swing.Icon iconoChevronIzquierda(java.awt.Color color) {
        return new VectorIcon(10, color, (g2, t) -> {
            g2.drawLine(t - 3, 1, 2, t / 2);
            g2.drawLine(2, t / 2, t - 3, t - 1);
        });
    }//iconoChevronIzquierda

    private static javax.swing.Icon iconoInicio(java.awt.Color color) {
        return new VectorIcon(13, color, (g2, t) -> {
            g2.drawLine(1, t / 2, t / 2, 1);
            g2.drawLine(t / 2, 1, t - 1, t / 2);
            g2.drawRect(3, t / 2, t - 6, t / 2 - 2);
        });
    }//iconoInicio

    private static javax.swing.Icon iconoInfo(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawOval(1, 1, t - 2, t - 2);
            g2.fillOval(t / 2 - 1, 3, 2, 2);
            g2.drawLine(t / 2, 7, t / 2, t - 3);
        });
    }//iconoInfo

    private static javax.swing.Icon iconoRadio(boolean marcado) {
        java.awt.Color color = marcado ? new java.awt.Color(0, 130, 60) : new java.awt.Color(160, 170, 180);
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawOval(2, 2, t - 4, t - 4);
            if (marcado) {
                g2.fillOval(5, 5, t - 10, t - 10);
            }
        });
    }//iconoRadio
}
