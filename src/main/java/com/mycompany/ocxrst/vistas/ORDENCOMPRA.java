/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.IconoVentanaUtil;
import com.mycompany.ocxrst.EstiloAplicacion;
import com.mycompany.ocxrst.config.AppConfig;
import com.mycompany.ocxrst.repository.OrdenCompraRepository;
import com.mycompany.ocxrst.service.InventarioService;
import com.mycompany.ocxrst.service.ProveedorService;
import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.forms.PdfAcroForm;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.signatures.BouncyCastleDigest;
import com.itextpdf.signatures.DigestAlgorithms;
import com.itextpdf.signatures.IExternalDigest;
import com.itextpdf.signatures.IExternalSignature;
import com.itextpdf.signatures.PdfSigner;
import com.itextpdf.signatures.PrivateKeySignature;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import com.itextpdf.layout.properties.VerticalAlignment;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.properties.AreaBreakType;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.security.Key;
import java.security.KeyStore;
import java.security.PrivateKey;
import java.security.Security;
import java.security.cert.X509Certificate;
import java.text.SimpleDateFormat;
import java.util.Locale;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bouncycastle.jce.provider.BouncyCastleProvider;
/**
 *
 * @author Maria Luisa Martinez
 */
public class ORDENCOMPRA extends javax.swing.JFrame {

    // =====================================================================
    // Rutas, repositorios y estado de la orden
    // =====================================================================

    private static final String RUTA_INVENTARIOS = AppConfig.inventariosFile().getAbsolutePath();
    private static final String RUTA_USUARIOS = AppConfig.usuariosFile().getAbsolutePath();
    private static final String RUTA_PAGOS = AppConfig.pagosFile().getAbsolutePath();
    private static final String RUTA_FORMAPAGO = AppConfig.formaPagoFile().getAbsolutePath();
    private static final String RUTA_METODOPAGO = AppConfig.metodoPagoFile().getAbsolutePath();
    private static final String RUTA_CFDI = AppConfig.cfdiFile().getAbsolutePath();
    private static final String RUTA_TIPOMONEDA = AppConfig.tipoMonedaFile().getAbsolutePath();
    private static final OrdenCompraRepository ORDEN_COMPRA_REPOSITORY = new OrdenCompraRepository();
    private final DataFormatter dataFormatter = new DataFormatter();
    private final java.util.Map<String, String> monedaDescMap = new java.util.HashMap<>();

    private static final java.util.logging.Logger logger = java.util.logging.Logger.getLogger(ORDENCOMPRA.class.getName());
    private final java.util.Map<String, String> monedaLeyendaMap = new java.util.HashMap<>();
    private boolean ultimoProductoDesactivado;
    private static final int DESPLAZAMIENTO_X = -14;
    private static final int DESPLAZAMIENTO_Y = -25;
    private static final int ALTO_EXTRA_PRODUCTOS = 70;
    /** Espacio adicional agregado a la sección 4 (Condiciones de compra) y
     *  desplazamiento aplicado a todas las secciones siguientes para que
     *  quepan sin encimarse. */
    private static final int EXTRA_ESPACIO_SECCION4 = 10;
    /** Alto reservado para la nueva sección "Observaciones", ubicada debajo
     *  de la tabla de productos y antes de Autorizaciones/Impuestos/Totales. */
    private static final int ALTO_OBSERVACIONES = 110;

    // Caché de inventario: código -> "código | descripción"
    private final java.util.List<String[]> cacheInventario = new java.util.ArrayList<>();
    private final javax.swing.JPopupMenu popupProductos = new javax.swing.JPopupMenu();

    // Caché de proveedores: id -> "id | nombre | activo"
    private final java.util.List<String[]> cacheProveedores = new java.util.ArrayList<>();
    private final javax.swing.JPopupMenu popupProveedores = new javax.swing.JPopupMenu();
    private boolean seleccionandoProveedor = false;
    private boolean modoCrear = false;
    private String estadoOrdenCargada;
    private javax.swing.JButton jButtonDesactivar;
    private javax.swing.JPanel panelNavegacionOrdenes;
    private javax.swing.JButton botonPrimeraOrden;
    private javax.swing.JButton botonAnteriorOrden;
    private javax.swing.JButton botonSiguienteOrden;
    private javax.swing.JButton botonUltimaOrden;
    private javax.swing.JButton jButtonFirmar;
    private javax.swing.JScrollPane desplazamientoFormulario;
    private javax.swing.JPanel panelBusquedaOrdenes;
    private java.awt.Container contenedorVista;
    private File ultimoPdfGenerado;

    private static final class DatosFirma {
        private final PrivateKey clavePrivada;
        private final java.security.cert.Certificate[] cadena;
        private final String rol;
        private final String alias;

        private DatosFirma(PrivateKey clavePrivada,
                java.security.cert.Certificate[] cadena, String rol, String alias) {
            this.clavePrivada = clavePrivada;
            this.cadena = cadena;
            this.rol = rol;
            this.alias = alias;
        }
    }

    // =====================================================================
    // Construccion de la ventana y configuracion visual
    // =====================================================================

    /** Inicializa el formulario, catalogos, caches y tabla de productos. */
    public ORDENCOMPRA() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        initComponents();
        contenedorVista = getContentPane();

        // Campo libre de observaciones/notas de la orden. Se crea aquí, antes de
        // aplicarDisenoReferencia(), porque ese método necesita posicionar estos
        // componentes y por lo tanto deben existir antes de invocarlo.
        jLabelObservaciones = new javax.swing.JLabel("Observaciones:");
        jLabelObservaciones.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        jTextAreaObservaciones = new javax.swing.JTextArea();
        jTextAreaObservaciones.setLineWrap(true);
        jTextAreaObservaciones.setWrapStyleWord(true);
        jTextAreaObservaciones.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(198, 210, 225)));
        jScrollPaneObservaciones = new javax.swing.JScrollPane(jTextAreaObservaciones);
        jScrollPaneObservaciones.setVerticalScrollBarPolicy(javax.swing.JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        jPanel1.add(jLabelObservaciones);
        jPanel1.add(jScrollPaneObservaciones);
        jButtonDesactivar = new javax.swing.JButton("DESACTIVAR");
        jPanel1.add(jButtonDesactivar);
        crearNavegacionOrdenes();

        aplicarDisenoReferencia();
        configurarDesplazamiento();
        IconoVentanaUtil.aplicar(this);
        configurarTablaProductos();
        cargarCFDI();
        cargarUsuario();
        cargarMoneda();
        cargarPagos();
        cargarFormaPago();
        cargarMetodoPago();
        cargarCacheInventario();
        cargarCacheProveedores();
        actualizarTotales();
        generarNumeroOrden();
        mostrarBusquedaOrdenes();
        setLocationRelativeTo(null); 
    
    jTextField1.addMouseListener(new java.awt.event.MouseAdapter() {
        @Override
        public void mouseClicked(java.awt.event.MouseEvent e) {
            if (modoCrear) {
                jTextField1.setEditable(true);
                jTextField1.requestFocusInWindow();
            }
        }
    });

    jTextField1.addFocusListener(new java.awt.event.FocusAdapter() {
    @Override
    public void focusLost(java.awt.event.FocusEvent e) {
        if (!e.isTemporary()) {
            popupProveedores.setVisible(false);
            buscarProveedor(jTextField1.getText());
        }
    }
});

    // ── AUTOCOMPLETE DE PROVEEDOR ─────────────────────────────────────────
    jTextField1.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
        public void insertUpdate(javax.swing.event.DocumentEvent e) { filtrarProveedores(); }
        public void removeUpdate(javax.swing.event.DocumentEvent e) { filtrarProveedores(); }
        public void changedUpdate(javax.swing.event.DocumentEvent e) { filtrarProveedores(); }
    });

    // ── EDICIÓN DE jLabel5 CON DOBLE CLIC ────────────────────────────────
    jLabel5.setCursor(java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR));
    jLabel5.addMouseListener(new java.awt.event.MouseAdapter() {
        @Override
        public void mouseClicked(java.awt.event.MouseEvent e) {
            if (e.getClickCount() == 2 && modoCrear) {
                String nuevo = JOptionPane.showInputDialog(
                        ORDENCOMPRA.this, "Editar documento:", jLabel5.getText());
                if (nuevo != null) {
                    jLabel5.setText(nuevo.trim());
                }
            }
        }
    });

    jTextField3.addFocusListener(new java.awt.event.FocusAdapter() {
    @Override
    public void focusLost(java.awt.event.FocusEvent e) {
        // Solo procesar si no se perdió el foco hacia el popup de autocompletado
        if (!e.isTemporary()) {
            popupProductos.setVisible(false);
            procesarCodigoProducto();
        }
    }
});

    // ── AUTOCOMPLETE DE PRODUCTO ──────────────────────────────────────────
    jTextField3.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
        public void insertUpdate(javax.swing.event.DocumentEvent e) { filtrarProductos(); }
        public void removeUpdate(javax.swing.event.DocumentEvent e) { filtrarProductos(); }
        public void changedUpdate(javax.swing.event.DocumentEvent e) { filtrarProductos(); }
    });
    
  jComboBox1.addActionListener(e -> {
    String seleccionado = jComboBox1.getSelectedItem().toString();
    buscarArea(seleccionado);
});

    jComboBox2.addActionListener(e -> {
        Object sel = jComboBox2.getSelectedItem();
        if (sel != null) {
            String desc = monedaDescMap.getOrDefault(sel.toString(), "");
            jLabel37.setText(desc);
            boolean mostrarCambio = sel.toString().equalsIgnoreCase("USD") || sel.toString().equalsIgnoreCase("EUR") || sel.toString().equalsIgnoreCase("CAD");
            jTextField5.setVisible(mostrarCambio);
            jLabel38.setVisible(mostrarCambio);
            actualizarTotales();
        }
    });

    jButton7.addActionListener(e -> eliminarFilaSeleccionada());
    jButton2.addActionListener(e -> buscarOSalirOrden());
    jButton3.addActionListener(e -> {
        if (estadoOrdenCargada != null) {
            prepararDuplicado();
        } else {
            guardarEnRegistroc();
        }
    });

    javax.swing.event.DocumentListener estadoFormularioListener = new javax.swing.event.DocumentListener() {
        public void insertUpdate(javax.swing.event.DocumentEvent e) { actualizarEstadoBotonPDF(); }
        public void removeUpdate(javax.swing.event.DocumentEvent e) { actualizarEstadoBotonPDF(); }
        public void changedUpdate(javax.swing.event.DocumentEvent e) { actualizarEstadoBotonPDF(); }
    };
    jTextField1.getDocument().addDocumentListener(estadoFormularioListener);
    jTextField2.getDocument().addDocumentListener(estadoFormularioListener);
    jTextAreaObservaciones.getDocument().addDocumentListener(estadoFormularioListener);
    jTable1.getModel().addTableModelListener(e -> actualizarEstadoBotonPDF());
    jDateChooser2.addPropertyChangeListener("date", e -> actualizarEstadoBotonPDF());
    jDateChooserEntrega.addPropertyChangeListener("date", e -> actualizarEstadoBotonPDF());
    jDateChooserPago.addPropertyChangeListener("date", e -> actualizarEstadoBotonPDF());
    actualizarEstadoBotonPDF();

    // Visibilidad inicial de tipo de cambio según moneda seleccionada
    {
        Object sel = jComboBox2.getSelectedItem();
        boolean mostrarCambio = sel != null && (sel.toString().equalsIgnoreCase("USD") || sel.toString().equalsIgnoreCase("EUR") || sel.toString().equalsIgnoreCase("CAD"));
        jTextField5.setVisible(mostrarCambio);
        jLabel38.setVisible(mostrarCambio);
    }

    // Recalcular totales al escribir el tipo de cambio
    jTextField5.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
        public void insertUpdate(javax.swing.event.DocumentEvent e) { actualizarTotales(); }
        public void removeUpdate(javax.swing.event.DocumentEvent e) { actualizarTotales(); }
        public void changedUpdate(javax.swing.event.DocumentEvent e) { actualizarTotales(); }
    });

    // Checkbox para incluir o excluir IVA
    jCheckBoxIVA = new javax.swing.JCheckBox("Aplicar IVA");
    jCheckBoxIVA.setSelected(true);
    jCheckBoxIVA.setBackground(jPanel1.getBackground());
    jCheckBoxIVA.addActionListener(e -> actualizarTotales());
    jPanel1.add(jCheckBoxIVA);

    // Checkbox para aplicar descuento al subtotal
    jCheckBoxDescuento = new javax.swing.JCheckBox("Descuento %:");
    jCheckBoxDescuento.setSelected(false);
    jCheckBoxDescuento.setBackground(jPanel1.getBackground());
    jTextFieldDescuentoPct = new javax.swing.JTextField("0");
    jTextFieldDescuentoPct.setFont(jTextFieldDescuentoPct.getFont().deriveFont(java.awt.Font.BOLD, 15f));
    jTextFieldDescuentoPct.setHorizontalAlignment(javax.swing.JTextField.RIGHT);
    jTextFieldDescuentoPct.setVisible(false);
    jCheckBoxDescuento.addActionListener(e -> {
        jTextFieldDescuentoPct.setVisible(jCheckBoxDescuento.isSelected());
        jPanel1.repaint();
        actualizarTotales();
    });
    jPanel1.add(jCheckBoxDescuento);
    jPanel1.add(jTextFieldDescuentoPct);

    // Checkbox para aplicar ISR como una retencion adicional al descuento
    jCheckBoxISR = new javax.swing.JCheckBox("ISR %:");
    jCheckBoxISR.setSelected(false);
    jCheckBoxISR.setBackground(jPanel1.getBackground());
    jTextFieldISRPercent = new javax.swing.JTextField("0");
    jTextFieldISRPercent.setFont(jTextFieldISRPercent.getFont().deriveFont(java.awt.Font.BOLD, 15f));
    jTextFieldISRPercent.setHorizontalAlignment(javax.swing.JTextField.RIGHT);
    jTextFieldISRPercent.setVisible(false);
    jCheckBoxISR.addActionListener(e -> {
        jTextFieldISRPercent.setVisible(jCheckBoxISR.isSelected());
        jPanel1.repaint();
        actualizarTotales();
    });
    jPanel1.add(jCheckBoxISR);
    jPanel1.add(jTextFieldISRPercent);

    // Mantener las casillas dentro de la sección 8 (Impuestos y descuentos) del formulario.
    int fieldH = 22;
    int checkX = 402;
    int descuentoY = 625 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES;
    jCheckBoxDescuento.setBounds(checkX + DESPLAZAMIENTO_X, descuentoY + DESPLAZAMIENTO_Y, 125, fieldH);
    jTextFieldDescuentoPct.setBounds(checkX + 126 + DESPLAZAMIENTO_X, descuentoY + DESPLAZAMIENTO_Y, 55, fieldH);
    int isrY = 649 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES;
    jCheckBoxISR.setBounds(checkX + DESPLAZAMIENTO_X, isrY + DESPLAZAMIENTO_Y, 125, fieldH);
    jTextFieldISRPercent.setBounds(checkX + 126 + DESPLAZAMIENTO_X, isrY + DESPLAZAMIENTO_Y, 55, fieldH);
    jCheckBoxIVA.setBounds(checkX + DESPLAZAMIENTO_X, 673 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES + DESPLAZAMIENTO_Y, 150, fieldH);
    jPanel1.setComponentZOrder(jCheckBoxIVA, 0);
    jPanel1.setComponentZOrder(jCheckBoxISR, 0);
    jPanel1.setComponentZOrder(jTextFieldISRPercent, 0);
    jPanel1.setComponentZOrder(jCheckBoxDescuento, 0);
    jPanel1.setComponentZOrder(jTextFieldDescuentoPct, 0);
    jTextFieldDescuentoPct.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
        public void insertUpdate(javax.swing.event.DocumentEvent e) { actualizarTotales(); }
        public void removeUpdate(javax.swing.event.DocumentEvent e) { actualizarTotales(); }
        public void changedUpdate(javax.swing.event.DocumentEvent e) { actualizarTotales(); }
    });
    jTextFieldISRPercent.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
        public void insertUpdate(javax.swing.event.DocumentEvent e) { actualizarTotales(); }
        public void removeUpdate(javax.swing.event.DocumentEvent e) { actualizarTotales(); }
        public void changedUpdate(javax.swing.event.DocumentEvent e) { actualizarTotales(); }
    });

    // Auto-llenar jTextField4 con AVANTE + AñoMesDía (editable)
    java.time.LocalDate hoy = java.time.LocalDate.now();
    String cotizacion = String.format("AVANTE%04d%02d%02d",
            hoy.getYear(), hoy.getMonthValue(), hoy.getDayOfMonth());
    jTextField4.setText(cotizacion);

    // Fecha actual en el selector de fecha
    jDateChooser2.setDate(new java.util.Date());
    configurarPermisosFechas(false);

    // Quitar negrita de todos los JLabel y JCheckBox del panel
    for (java.awt.Component c : jPanel1.getComponents()) {
        if (c instanceof javax.swing.JLabel || c instanceof javax.swing.JCheckBox) {
            java.awt.Font f = c.getFont();
            if (f != null && f.isBold()) {
                c.setFont(f.deriveFont(java.awt.Font.PLAIN));
            }
        }
    }

    // Alinear a la derecha todos los labels de campo (terminan en ":")
    for (java.awt.Component c : jPanel1.getComponents()) {
        if (c instanceof javax.swing.JLabel lbl) {
            String txt = lbl.getText();
            if (txt != null && txt.trim().endsWith(":")) {
                lbl.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
            }
        }
    }
    // El rótulo de la sección Observaciones se mantiene alineado a la izquierda.
    jLabelObservaciones.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);

    // La edicion permanece visible y solo se habilita para administradores con una orden cargada.
    // Un único botón alterna entre "EDITAR" y "GUARDAR CAMBIOS" según el modo actual.
    jButtonEditar = new javax.swing.JButton("EDITAR");
    jButtonEditar.setBackground(new java.awt.Color(0, 90, 180));
    jButtonEditar.setForeground(java.awt.Color.WHITE);
    jButtonEditar.addActionListener(e -> onBotonEditarGuardarClick());
    jPanel1.add(jButtonEditar);
    jPanel1.setComponentZOrder(jButtonEditar, 0);
    prepararBotonCabecera(jButtonEditar, "EDITAR", new java.awt.Color(31, 112, 184), java.awt.Color.WHITE);
    jButtonFirmar = new javax.swing.JButton("FIRMAR PDF");
    jButtonFirmar.addActionListener(e -> firmarPdf());
    jPanel1.add(jButtonFirmar);
    jPanel1.setComponentZOrder(jButtonFirmar, 0);
    prepararBotonCabecera(jButtonFirmar, "FIRMAR PDF", new java.awt.Color(0, 128, 96), java.awt.Color.WHITE);
    posicionarBotonesCabecera(jPanel1.getPreferredSize().width);
    jButtonFirmar.setBounds(732, 38, 92, 30);
    actualizarEstadoBotonFirma();
    actualizarBotonesEdicion(false, false);

    jPanel1.addComponentListener(new java.awt.event.ComponentAdapter() {
        @Override public void componentResized(java.awt.event.ComponentEvent e) {
            int panelW = jPanel1.getWidth();
            if (panelW > 0 && jButtonEditar != null) {
                posicionarBotonesCabecera(panelW);
            }
        }
    });
    }

    private void posicionarBotonesCabecera(int anchoPanel) {
        int y = 4;
        jButton2.setBounds(300, y, 100, 30);
        jButton3.setBounds(408, y, 100, 30);
        jButtonEditar.setBounds(516, y, 100, 30);
        jButtonDesactivar.setBounds(624, y, 100, 30);
        jButton1.setBounds(732, y, 92, 30);
        panelNavegacionOrdenes.setBounds(832, y, 108, 30);
    }

    private void crearNavegacionOrdenes() {
        botonPrimeraOrden = botonNavegacion("|<", this::irPrimeraOrden);
        botonAnteriorOrden = botonNavegacion("<", () -> navegarOrden(-1));
        botonSiguienteOrden = botonNavegacion(">", () -> navegarOrden(1));
        botonUltimaOrden = botonNavegacion(">|", this::irUltimaOrden);
        panelNavegacionOrdenes = new javax.swing.JPanel(new java.awt.GridLayout(1, 4, 2, 0));
        panelNavegacionOrdenes.setOpaque(true);
        panelNavegacionOrdenes.setBackground(new java.awt.Color(255, 238, 218));
        panelNavegacionOrdenes.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(245, 112, 32), 2));
        panelNavegacionOrdenes.add(botonPrimeraOrden);
        panelNavegacionOrdenes.add(botonAnteriorOrden);
        panelNavegacionOrdenes.add(botonSiguienteOrden);
        panelNavegacionOrdenes.add(botonUltimaOrden);
        jPanel1.add(panelNavegacionOrdenes);
    }

    private javax.swing.JButton botonNavegacion(String texto, Runnable accion) {
        javax.swing.JButton boton = new javax.swing.JButton(texto);
        boton.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 12));
        boton.setForeground(new java.awt.Color(28, 82, 132));
        boton.setBackground(java.awt.Color.WHITE);
        boton.setFocusPainted(false);
        boton.setBorder(javax.swing.BorderFactory.createEmptyBorder());
        boton.setContentAreaFilled(false);
        boton.setOpaque(false);
        boton.setRolloverEnabled(false);
        boton.setMargin(new java.awt.Insets(0, 0, 0, 0));
        boton.addActionListener(e -> accion.run());
        return boton;
    }

    private void prepararBotonCabecera(javax.swing.JButton boton, String texto,
            java.awt.Color fondo, java.awt.Color textoColor) {
        boton.setText(texto);
        boton.setBackground(fondo);
        boton.setForeground(textoColor);
        boton.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        boton.setFocusPainted(false);
        boton.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(new java.awt.Color(180, 195, 215)),
                javax.swing.BorderFactory.createEmptyBorder(4, 8, 4, 8)));
        boton.setCursor(java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR));
    }

    private void prepararValorProveedor(javax.swing.JLabel valor) {
        valor.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        valor.setVerticalAlignment(javax.swing.SwingConstants.TOP);
        valor.setForeground(new java.awt.Color(38, 52, 88));
        valor.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 4, 1, 2));
    }

    private javax.swing.JPanel crearPanelConFondo() {
        return new javax.swing.JPanel() {
            @Override
            protected void paintComponent(java.awt.Graphics graphics) {
                super.paintComponent(graphics);
                java.awt.Graphics2D g2 = (java.awt.Graphics2D) graphics.create();
                g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                        java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
                dibujarSeccion(g2, 48, 112, 866, 80, "1. DATOS GENERALES");
                dibujarSeccion(g2, 48, 200, 485, 125, "2. PROVEEDOR");
                dibujarSeccion(g2, 544, 200, 370, 125, "3. SOLICITANTE / PROYECTO");
                dibujarSeccion(g2, 48, 333, 866, 113 + EXTRA_ESPACIO_SECCION4, "4. CONDICIONES DE COMPRA");
                dibujarSeccion(g2, 48, 453 + EXTRA_ESPACIO_SECCION4, 866, 139 + ALTO_EXTRA_PRODUCTOS, "5. PRODUCTOS");
                dibujarSeccion(g2, 48, 599 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4, 866, ALTO_OBSERVACIONES, "6. OBSERVACIONES");
                dibujarSeccion(g2, 48, 599 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 300, 101, "7. AUTORIZACIONES");
                dibujarSeccion(g2, 359, 599 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 267, 101, "8. IMPUESTOS Y DESCUENTOS");
                dibujarSeccion(g2, 637, 599 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 277, 101, "9. TOTALES");
                g2.dispose();
            }
        };
    }

    private void dibujarSeccion(java.awt.Graphics2D graphics, int x, int y, int ancho,
                                int alto, String titulo) {
        x += DESPLAZAMIENTO_X;
        y += DESPLAZAMIENTO_Y;
        graphics.setColor(new java.awt.Color(255, 255, 255, 224));
        graphics.fillRoundRect(x, y, ancho, alto, 10, 10);
        graphics.setColor(new java.awt.Color(90, 178, 220));
        graphics.fillRoundRect(x, y, ancho, 22, 10, 10);
        graphics.fillRect(x, y + 11, ancho, 11);
        graphics.setColor(new java.awt.Color(145, 204, 238));
        graphics.drawRoundRect(x, y, ancho, alto, 10, 10);
        graphics.setColor(java.awt.Color.WHITE);
        graphics.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 12));
        graphics.drawString(titulo, x + 14, y + 16);
    }

    /** Orden visual de la ventana según el formato de la orden impresa. */
    private void aplicarDisenoReferencia() {
        jPanel1.setLayout(null);
        jPanel1.setPreferredSize(new java.awt.Dimension(962, 735));
        setResizable(true);

        jButton6.setVisible(true);
        jButton6.setText("ATRÁS");
        prepararBotonCabecera(jButton6, "ATRÁS", new java.awt.Color(31, 112, 184), java.awt.Color.WHITE);
        posicionar(jButton6, 22, 28, 100, 30);
        posicionar(jLabel5, 430, 64, 100, 22);
        posicionar(jLabel1, 280, 78, 400, 38);
        posicionar(jLabel9, 62, 140, 105, 22);
        posicionar(jTextField6, 174, 140, 158, 24);
        posicionar(jLabel11, 450, 140, 90, 22);
        posicionar(jTextField4, 552, 140, 292, 24);
        posicionar(jLabel6, 62, 168, 105, 22);
        posicionar(jTextField1, 174, 168, 122, 24);
        posicionar(jLabel3, 330, 168, 62, 22);
        posicionar(jDateChooser2, 400, 168, 125, 24);
        posicionar(jLabel17, 552, 168, 78, 22);
        posicionar(jComboBox3, 638, 168, 206, 24);

        posicionar(jLabel2, 70, 222, 90, 22);
        posicionar(jLabel8, 174, 222, 330, 22);
        posicionar(jLabel29, 70, 249, 90, 22);
        posicionar(jLabel30, 174, 249, 330, 22);
        posicionar(jLabel14, 70, 276, 90, 22);
        posicionar(jLabel15, 174, 274, 330, 30);
        posicionar(jLabel31, 70, 304, 90, 22);
        posicionar(jLabel32, 174, 305, 330, 20);
        prepararValorProveedor(jLabel8);
        prepararValorProveedor(jLabel30);
        prepararValorProveedor(jLabel15);
        prepararValorProveedor(jLabel32);
        jLabel15.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 10));
        posicionar(jLabel7, 560, 227, 90, 22);
        posicionar(jComboBox1, 660, 227, 228, 24);
        posicionar(jLabel4, 560, 254, 90, 22);
        posicionar(jLabel26, 660, 254, 228, 22);
        posicionar(jLabel13, 560, 281, 90, 22);
        posicionar(jTextField2, 660, 281, 228, 24);

        posicionar(jLabel36, 62, 359, 125, 22);
        posicionar(jComboBox2, 195, 359, 125, 24);
        posicionar(jLabel37, 310, 359, 170, 22);
        posicionar(jLabel18, 62, 391, 125, 22);
        posicionar(jComboBox6, 195, 391, 170, 24);
        posicionar(jLabel19, 62, 423, 125, 22);
        posicionar(jComboBox8, 195, 423, 170, 24);
        posicionar(jLabel38, 475, 359, 125, 22);
        posicionar(jTextField5, 610, 359, 120, 24);
        posicionar(jLabel34, 475, 391, 125, 22);
        posicionar(jComboBox7, 610, 391, 258, 24);
        posicionar(jLabel12, 350, 423, 140, 22);
        posicionar(jDateChooserEntrega, 495, 423, 120, 24);
        posicionar(jLabel27, 625, 423, 115, 22);
        posicionar(jDateChooserPago, 745, 423, 120, 24);

        posicionar(jLabel16, 70, 478 + EXTRA_ESPACIO_SECCION4, 90, 22);
        posicionar(jTextField3, 164, 478 + EXTRA_ESPACIO_SECCION4, 180, 24);
        posicionar(jButton7, 352, 478 + EXTRA_ESPACIO_SECCION4, 82, 24);
        posicionar(jScrollPane1, 60, 508 + EXTRA_ESPACIO_SECCION4, 842, 76 + ALTO_EXTRA_PRODUCTOS);
        posicionar(jLabelObservaciones, 60, 625 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4, 200, 22);
        jLabelObservaciones.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        posicionar(jScrollPaneObservaciones, 60, 649 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4, 842, ALTO_OBSERVACIONES - 58);
        posicionar(jLabel39, 70, 625 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 70, 22);
        posicionar(jComboBox4, 145, 625 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 180, 24);
        posicionar(jLabel40, 70, 662 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 70, 22);
        posicionar(jComboBox5, 145, 662 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 180, 24);
        posicionar(jLabel20, 650, 620 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 105, 22);
        posicionar(jLabel23, 765, 620 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 135, 22);
        posicionar(jLabel21, 650, 646 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 105, 22);
        posicionar(jLabel24, 765, 646 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 135, 22);
        posicionar(jLabel22, 650, 672 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 105, 22);
        posicionar(jLabel25, 765, 672 + ALTO_EXTRA_PRODUCTOS + EXTRA_ESPACIO_SECCION4 + ALTO_OBSERVACIONES, 135, 22);
        jButton3.setText("NUEVO");
        jButton1.setText("IMPRIMIR");
        jButtonDesactivar.setEnabled(false);
        prepararBotonCabecera(jButton2, "BUSCAR", new java.awt.Color(31, 112, 184), java.awt.Color.WHITE);
        prepararBotonCabecera(jButton3, "NUEVO", new java.awt.Color(31, 112, 184), java.awt.Color.WHITE);
        prepararBotonCabecera(jButtonDesactivar, "DESACTIVAR", new java.awt.Color(170, 178, 188), java.awt.Color.WHITE);
        prepararBotonCabecera(jButton1, "IMPRIMIR", java.awt.Color.WHITE, new java.awt.Color(28, 82, 132));
        ajustarTamanoPanelAlContenido();
        pack();
        jPanel1.revalidate();
        jPanel1.repaint();
    }

    private void ajustarTamanoPanelAlContenido() {
        int margenInferior = 24;
        int margenDerecho = 24;
        int ancho = 962;
        int alto = 735;
        for (java.awt.Component componente : jPanel1.getComponents()) {
            java.awt.Rectangle limites = componente.getBounds();
            ancho = Math.max(ancho, limites.x + limites.width + margenDerecho);
            alto = Math.max(alto, limites.y + limites.height + margenInferior);
        }
        jPanel1.setPreferredSize(new java.awt.Dimension(ancho, alto));
        jPanel1.setMinimumSize(new java.awt.Dimension(ancho, alto));
    }

    private void configurarDesplazamiento() {
        desplazamientoFormulario = new javax.swing.JScrollPane(jPanel1);
        desplazamientoFormulario.setBorder(javax.swing.BorderFactory.createEmptyBorder());
        desplazamientoFormulario.setHorizontalScrollBarPolicy(
                javax.swing.JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        desplazamientoFormulario.setVerticalScrollBarPolicy(
                javax.swing.JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        desplazamientoFormulario.getVerticalScrollBar().setUnitIncrement(18);
        desplazamientoFormulario.getViewport().setBackground(java.awt.Color.WHITE);

        contenedorVista.removeAll();
        contenedorVista.setLayout(new java.awt.BorderLayout());
        contenedorVista.add(desplazamientoFormulario, java.awt.BorderLayout.CENTER);
        setSize(962, 735);
        setMinimumSize(new java.awt.Dimension(520, 420));
        setLocationRelativeTo(null);
    }

    private void mostrarBusquedaOrdenes() {
        if (panelBusquedaOrdenes == null) {
            panelBusquedaOrdenes = crearPanelBusquedaOrdenes();
        }
        mostrarVistaInterna(panelBusquedaOrdenes);
        setTitle("ÓRDENES DE COMPRA");
        setSize(1000, 680);
        setMinimumSize(new java.awt.Dimension(850, 540));
        cargarOrdenesEnBusqueda();
    }

    private void mostrarVistaInterna(java.awt.Component vista) {
        java.awt.Container padreActual = vista.getParent();
        if (padreActual != null && padreActual != contenedorVista) {
            padreActual.remove(vista);
            padreActual.revalidate();
            padreActual.repaint();
        }
        contenedorVista.removeAll();
        contenedorVista.setLayout(new java.awt.BorderLayout());
        contenedorVista.add(vista, java.awt.BorderLayout.CENTER);
        contenedorVista.revalidate();
        contenedorVista.repaint();
        java.awt.Container padre = contenedorVista.getParent();
        if (padre != null) {
            padre.revalidate();
            padre.repaint();
        }
    }

    private javax.swing.JPanel crearPanelBusquedaOrdenes() {
        final java.awt.Color azul = new java.awt.Color(28, 82, 132);
        javax.swing.JPanel fondo = new javax.swing.JPanel(new java.awt.BorderLayout(0, 12));
        fondo.setBackground(new java.awt.Color(240, 246, 253));
        fondo.setBorder(javax.swing.BorderFactory.createEmptyBorder(24, 24, 20, 24));

        javax.swing.JPanel encabezado = new javax.swing.JPanel(new java.awt.BorderLayout(12, 0));
        encabezado.setOpaque(false);

        javax.swing.JButton nueva = botonBusqueda("+  Nueva orden de compra", true);
        nueva.addActionListener(e -> abrirNuevaOrden());
        javax.swing.JButton actualizar = botonBusqueda("Actualizar", false);
        actualizar.addActionListener(e -> cargarOrdenesEnBusqueda());
        javax.swing.JPanel acciones = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 8, 5));
        acciones.setOpaque(false);
        acciones.add(nueva);
        acciones.add(actualizar);
        encabezado.add(acciones, java.awt.BorderLayout.EAST);

        javax.swing.JPanel resumen = new javax.swing.JPanel(new java.awt.GridLayout(1, 4, 10, 0));
        resumen.setOpaque(false);
        javax.swing.JLabel totalOrdenesResumen = new javax.swing.JLabel("0");
        javax.swing.JLabel pendientesResumen = new javax.swing.JLabel("0");
        javax.swing.JLabel procesoResumen = new javax.swing.JLabel("0");
        javax.swing.JLabel importeResumen = new javax.swing.JLabel("$ 0.00");
        resumen.add(crearTarjetaResumen("Total de órdenes", totalOrdenesResumen,
                new java.awt.Color(225, 239, 255), "documento"));
        resumen.add(crearTarjetaResumen("Pendientes de autorización", pendientesResumen,
                new java.awt.Color(255, 242, 211), "reloj"));
        resumen.add(crearTarjetaResumen("En proceso", procesoResumen,
                new java.awt.Color(225, 239, 255), "configuracion"));
        resumen.add(crearTarjetaResumen("Importe del mes", importeResumen,
                new java.awt.Color(224, 246, 231), "dinero"));
        javax.swing.JPanel zonaSuperior = new javax.swing.JPanel(new java.awt.BorderLayout(0, 16));
        zonaSuperior.setOpaque(false);
        zonaSuperior.add(encabezado, java.awt.BorderLayout.NORTH);
        zonaSuperior.add(resumen, java.awt.BorderLayout.CENTER);
        fondo.add(zonaSuperior, java.awt.BorderLayout.NORTH);

        javax.swing.JPanel tarjeta = new javax.swing.JPanel(new java.awt.BorderLayout(0, 12));
        tarjeta.setBackground(java.awt.Color.WHITE);
        tarjeta.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(new java.awt.Color(205, 218, 235)),
                javax.swing.BorderFactory.createEmptyBorder(16, 16, 16, 16)));

        javax.swing.JPanel filtros = new javax.swing.JPanel(new java.awt.GridBagLayout());
        filtros.setOpaque(false);
        javax.swing.JTextField buscar = new javax.swing.JTextField();
        com.toedter.calendar.JDateChooser fechaInicio = crearSelectorFechaBusqueda();
        com.toedter.calendar.JDateChooser fechaFin = crearSelectorFechaBusqueda();
        javax.swing.JComboBox<String> estado = new javax.swing.JComboBox<>(
                new String[]{"Todos", "Pendiente", "En proceso", "Autorizada", "Borrador"});
        buscar.setName("buscarOrden");
        fechaInicio.setName("fechaInicioOrden");
        fechaFin.setName("fechaFinOrden");
        prepararCampoBusqueda(buscar);
        agregarFiltroBusqueda(filtros, "Buscar por folio, proveedor o proyecto", buscar, 0, 1.25);
        agregarFiltroBusqueda(filtros, "Fecha inicio", fechaInicio, 1, 0.65);
        agregarFiltroBusqueda(filtros, "Fecha fin", fechaFin, 2, 0.65);
        agregarFiltroBusqueda(filtros, "Estado", estado, 3, 0.75);
        javax.swing.JButton limpiar = botonBusqueda("Limpiar filtros", false);
        limpiar.addActionListener(e -> {
            buscar.setText("");
            fechaInicio.setDate(null);
            fechaFin.setDate(null);
            estado.setSelectedIndex(0);
        });
        agregarFiltroBusqueda(filtros, "", limpiar, 4, 0.9);
        tarjeta.add(filtros, java.awt.BorderLayout.NORTH);

        javax.swing.table.DefaultTableModel modeloBusqueda = new javax.swing.table.DefaultTableModel(
                new String[]{"FOLIO", "FECHA", "PROVEEDOR", "SOLICITANTE", "PROYECTO", "TOTAL", "ESTADO", "ACCIONES"}, 0) {
            @Override public boolean isCellEditable(int fila, int columna) { return false; }
        };
        javax.swing.JTable tabla = new javax.swing.JTable(modeloBusqueda);
        tabla.setName("tablaOrdenes");
        tabla.setRowHeight(42);
        tabla.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 11));
        tabla.getTableHeader().setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        tabla.getTableHeader().setPreferredSize(new java.awt.Dimension(0, 38));
        tabla.getTableHeader().setBackground(new java.awt.Color(232, 238, 247));
        tabla.getTableHeader().setForeground(azul);
        tabla.setGridColor(new java.awt.Color(225, 232, 241));
        tabla.setSelectionBackground(new java.awt.Color(225, 238, 250));
        tabla.setSelectionForeground(azul);
        tabla.getColumnModel().getColumn(0).setPreferredWidth(70);
        tabla.getColumnModel().getColumn(1).setPreferredWidth(100);
        tabla.getColumnModel().getColumn(2).setPreferredWidth(220);
        tabla.getColumnModel().getColumn(5).setPreferredWidth(110);
        tabla.getColumnModel().getColumn(6).setPreferredWidth(105);
        tabla.getColumnModel().getColumn(7).setPreferredWidth(80);
        final int[] celdaAccionHover = {-1, -1};
        tabla.getColumnModel().getColumn(6).setCellRenderer((tablaActual, valor, seleccionado, tieneFoco, fila, columna) -> {
            javax.swing.JLabel estadoCelda = new javax.swing.JLabel(String.valueOf(valor));
            estadoCelda.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
            estadoCelda.setOpaque(true);
            estadoCelda.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 10));
            estadoCelda.setForeground(new java.awt.Color(28, 82, 132));
            estadoCelda.setBackground(new java.awt.Color(225, 239, 255));
            estadoCelda.setBorder(javax.swing.BorderFactory.createEmptyBorder(5, 8, 5, 8));
            return estadoCelda;
        });
        tabla.getColumnModel().getColumn(7).setCellRenderer((tablaActual, valor, seleccionado, tieneFoco, fila, columna) -> {
            javax.swing.JLabel enlace = new javax.swing.JLabel();
            boolean hover = fila == celdaAccionHover[0] && columna == 7;
            java.awt.Component celdaBase = tablaActual.getDefaultRenderer(Object.class)
                    .getTableCellRendererComponent(tablaActual,
                            tablaActual.getValueAt(fila, 0), seleccionado, tieneFoco, fila, 0);
            enlace.setText(hover ? "<html><u><font color='#1F70B8'>VER</font></u></html>" : "VER");
            enlace.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
            enlace.setFont(tablaActual.getFont());
            enlace.setForeground(seleccionado
                    ? tablaActual.getSelectionForeground()
                    : new java.awt.Color(28, 82, 132));
            enlace.setOpaque(true);
            enlace.setBackground(celdaBase.getBackground());
            return enlace;
        });
        tabla.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override public void mouseClicked(java.awt.event.MouseEvent e) {
                int fila = tabla.rowAtPoint(e.getPoint());
                int columna = tabla.columnAtPoint(e.getPoint());
                if (fila < 0 || columna != 7) return;
                String numero = String.valueOf(modeloBusqueda.getValueAt(fila, 0));
                abrirOrdenDesdeBusqueda(numero, false);
            }

            @Override public void mouseExited(java.awt.event.MouseEvent e) {
                celdaAccionHover[0] = -1;
                celdaAccionHover[1] = -1;
                tabla.repaint();
                tabla.setCursor(java.awt.Cursor.getDefaultCursor());
            }
        });
        tabla.addMouseMotionListener(new java.awt.event.MouseMotionAdapter() {
            @Override public void mouseMoved(java.awt.event.MouseEvent e) {
                int fila = tabla.rowAtPoint(e.getPoint());
                int columna = tabla.columnAtPoint(e.getPoint());
                int filaAnterior = celdaAccionHover[0];
                int columnaAnterior = celdaAccionHover[1];
                if (fila >= 0 && columna == 7) {
                    celdaAccionHover[0] = fila;
                    celdaAccionHover[1] = columna;
                    tabla.setCursor(java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR));
                } else {
                    celdaAccionHover[0] = -1;
                    celdaAccionHover[1] = -1;
                    tabla.setCursor(java.awt.Cursor.getDefaultCursor());
                }
                if (filaAnterior != celdaAccionHover[0] || columnaAnterior != celdaAccionHover[1]) {
                    tabla.repaint();
                }
            }
        });
        javax.swing.JLabel total = new javax.swing.JLabel();
        total.setName("totalOrdenes");
        total.setForeground(azul);
        total.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        tarjeta.add(new javax.swing.JScrollPane(tabla), java.awt.BorderLayout.CENTER);
        javax.swing.JPanel pie = new javax.swing.JPanel(new java.awt.BorderLayout());
        pie.setOpaque(false);
        pie.add(total, java.awt.BorderLayout.EAST);
        tarjeta.add(pie, java.awt.BorderLayout.SOUTH);
        fondo.add(tarjeta, java.awt.BorderLayout.CENTER);

        java.beans.PropertyChangeListener actualizarTabla = e -> cargarOrdenesEnTabla(
                buscar, fechaInicio, fechaFin, estado, modeloBusqueda, total,
                totalOrdenesResumen, pendientesResumen, procesoResumen, importeResumen);
        buscar.getDocument().addDocumentListener(documentListenerBusqueda(actualizarTabla));
        fechaInicio.addPropertyChangeListener("date", actualizarTabla);
        fechaFin.addPropertyChangeListener("date", actualizarTabla);
        estado.addActionListener(e -> actualizarTabla.propertyChange(null));
        fondo.putClientProperty("controlesBusqueda", new Object[]{buscar, fechaInicio, fechaFin, estado,
            modeloBusqueda, total, totalOrdenesResumen, pendientesResumen, procesoResumen, importeResumen});
        return fondo;
    }

    private javax.swing.JPanel crearTarjetaResumen(String etiqueta, javax.swing.JLabel valor,
            java.awt.Color fondoTarjeta, String tipoIcono) {
        javax.swing.JPanel tarjeta = new javax.swing.JPanel(new java.awt.BorderLayout(10, 0));
        tarjeta.setBackground(java.awt.Color.WHITE);
        tarjeta.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(new java.awt.Color(224, 232, 242)),
                javax.swing.BorderFactory.createEmptyBorder(12, 14, 12, 14)));
        javax.swing.JPanel icono = new javax.swing.JPanel(new java.awt.BorderLayout());
        icono.setPreferredSize(new java.awt.Dimension(42, 42));
        icono.setBackground(fondoTarjeta);
        icono.add(new javax.swing.JLabel(crearIconoResumen(tipoIcono,
                new java.awt.Color(31, 112, 184))), java.awt.BorderLayout.CENTER);
        tarjeta.add(icono, java.awt.BorderLayout.WEST);
        javax.swing.JPanel textos = new javax.swing.JPanel();
        textos.setOpaque(false);
        textos.setLayout(new javax.swing.BoxLayout(textos, javax.swing.BoxLayout.Y_AXIS));
        javax.swing.JLabel titulo = new javax.swing.JLabel(etiqueta);
        titulo.setForeground(new java.awt.Color(93, 119, 153));
        titulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 11));
        valor.setForeground(new java.awt.Color(22, 48, 88));
        valor.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 20));
        textos.add(titulo);
        textos.add(javax.swing.Box.createVerticalStrut(3));
        textos.add(valor);
        tarjeta.add(textos, java.awt.BorderLayout.CENTER);
        return tarjeta;
    }

    private javax.swing.Icon crearIconoResumen(String tipo, java.awt.Color color) {
        return new javax.swing.Icon() {
            @Override public int getIconWidth() { return 23; }
            @Override public int getIconHeight() { return 23; }

            @Override
            public void paintIcon(java.awt.Component componente, java.awt.Graphics graphics, int x, int y) {
                java.awt.Graphics2D g = (java.awt.Graphics2D) graphics.create();
                g.translate(x, y);
                g.setColor(color);
                g.setStroke(new java.awt.BasicStroke(1.8f,
                        java.awt.BasicStroke.CAP_ROUND, java.awt.BasicStroke.JOIN_ROUND));
                g.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                        java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
                if ("documento".equals(tipo)) {
                    g.drawRoundRect(4, 2, 14, 19, 2, 2);
                    g.drawLine(7, 7, 15, 7);
                    g.drawLine(7, 11, 15, 11);
                    g.drawLine(7, 15, 12, 15);
                } else if ("reloj".equals(tipo)) {
                    g.drawOval(2, 2, 19, 19);
                    g.drawLine(11, 6, 11, 12);
                    g.drawLine(11, 12, 15, 15);
                } else if ("configuracion".equals(tipo)) {
                    g.drawOval(5, 5, 12, 12);
                    for (int i = 0; i < 8; i++) {
                        double angulo = i * Math.PI / 4;
                        int x1 = 11 + (int) (8 * Math.cos(angulo));
                        int y1 = 11 + (int) (8 * Math.sin(angulo));
                        int x2 = 11 + (int) (10 * Math.cos(angulo));
                        int y2 = 11 + (int) (10 * Math.sin(angulo));
                        g.drawLine(x1, y1, x2, y2);
                    }
                } else {
                    g.drawOval(3, 2, 17, 19);
                    g.drawString("$", 8, 17);
                }
                g.dispose();
            }
        };
    }

    private javax.swing.JButton botonBusqueda(String texto, boolean principal) {
        javax.swing.JButton boton = new javax.swing.JButton(texto);
        boton.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        boton.setFocusPainted(false);
        boton.setPreferredSize(new java.awt.Dimension(principal ? 190 : 135, 44));
        boton.setBackground(principal ? new java.awt.Color(31, 112, 184) : java.awt.Color.WHITE);
        boton.setForeground(principal ? java.awt.Color.WHITE : new java.awt.Color(28, 82, 132));
        boton.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(31, 112, 184), 1));
        return boton;
    }

    private void prepararCampoBusqueda(javax.swing.JComponent campo) {
        if (campo instanceof javax.swing.JTextField) {
            campo.setToolTipText("Buscar por folio, proveedor o proyecto");
            campo.setBackground(java.awt.Color.WHITE);
            campo.setOpaque(true);
            javax.swing.border.Border bordeNativo = new javax.swing.border.BevelBorder(
                    javax.swing.border.BevelBorder.LOWERED,
                    new java.awt.Color(250, 250, 250),
                    new java.awt.Color(225, 225, 225),
                    new java.awt.Color(135, 135, 135),
                    new java.awt.Color(170, 170, 170));
            campo.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                    bordeNativo,
                    javax.swing.BorderFactory.createEmptyBorder(4, 9, 4, 9)));
            campo.setPreferredSize(new java.awt.Dimension(0, 28));
            return;
        }
        campo.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(new java.awt.Color(198, 210, 225)),
                javax.swing.BorderFactory.createEmptyBorder(7, 9, 7, 9)));
    }

    private void agregarFiltroBusqueda(javax.swing.JPanel panel, String texto, java.awt.Component campo,
            int columna, double peso) {
        javax.swing.JPanel envoltura = new javax.swing.JPanel(new java.awt.BorderLayout(0, 4));
        envoltura.setOpaque(false);
        javax.swing.JLabel etiqueta = new javax.swing.JLabel(texto);
        etiqueta.setForeground(new java.awt.Color(28, 82, 132));
        etiqueta.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        envoltura.add(etiqueta, java.awt.BorderLayout.NORTH);
        envoltura.add(campo, java.awt.BorderLayout.CENTER);
        java.awt.GridBagConstraints gbc = new java.awt.GridBagConstraints();
        gbc.gridx = columna; gbc.gridy = 0; gbc.weightx = peso; gbc.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gbc.insets = new java.awt.Insets(0, 0, 0, columna == 4 ? 0 : 10);
        panel.add(envoltura, gbc);
    }

    private javax.swing.event.DocumentListener documentListenerBusqueda(java.beans.PropertyChangeListener listener) {
        return new javax.swing.event.DocumentListener() {
            private void actualizar() { listener.propertyChange(null); }
            public void insertUpdate(javax.swing.event.DocumentEvent e) { actualizar(); }
            public void removeUpdate(javax.swing.event.DocumentEvent e) { actualizar(); }
            public void changedUpdate(javax.swing.event.DocumentEvent e) { actualizar(); }
        };
    }

    private com.toedter.calendar.JDateChooser crearSelectorFechaBusqueda() {
        com.toedter.calendar.JDateChooser selector = new com.toedter.calendar.JDateChooser();
        selector.setLocale(java.util.Locale.forLanguageTag("es-MX"));
        selector.setDateFormatString("d MMM yyyy");
        selector.setPreferredSize(new java.awt.Dimension(155, 28));
        selector.setMinimumSize(new java.awt.Dimension(155, 28));
        selector.setToolTipText("Selecciona una fecha");
        return selector;
    }

    private void cargarOrdenesEnBusqueda() {
        if (panelBusquedaOrdenes == null) return;
        Object[] controles = (Object[]) panelBusquedaOrdenes.getClientProperty("controlesBusqueda");
        cargarOrdenesEnTabla((javax.swing.JTextField) controles[0],
                (com.toedter.calendar.JDateChooser) controles[1],
                (com.toedter.calendar.JDateChooser) controles[2],
                (javax.swing.JComboBox<?>) controles[3],
                (javax.swing.table.DefaultTableModel) controles[4], (javax.swing.JLabel) controles[5],
                (javax.swing.JLabel) controles[6], (javax.swing.JLabel) controles[7],
                (javax.swing.JLabel) controles[8], (javax.swing.JLabel) controles[9]);
    }

    private void cargarOrdenesEnTabla(javax.swing.JTextField buscar,
            com.toedter.calendar.JDateChooser fechaInicio,
            com.toedter.calendar.JDateChooser fechaFin, javax.swing.JComboBox<?> estadoFiltro,
            javax.swing.table.DefaultTableModel modelo, javax.swing.JLabel total,
            javax.swing.JLabel totalResumen, javax.swing.JLabel pendientesResumen,
            javax.swing.JLabel procesoResumen, javax.swing.JLabel importeResumen) {
        modelo.setRowCount(0);
        String texto = buscar.getText().trim().toLowerCase(java.util.Locale.ROOT);
        java.util.Date inicio = fechaInicio.getDate();
        java.util.Date fin = fechaFin.getDate();
        String estadoSeleccionado = String.valueOf(estadoFiltro.getSelectedItem());
        int pendientes = 0;
        int enProceso = 0;
        double importeMes = 0;
        java.io.File archivo = new java.io.File(RUTA_REGISTROC);
        if (archivo.isFile()) {
            try (Workbook libro = ORDEN_COMPRA_REPOSITORY.abrirWorkbook(archivo)) {
                for (Row fila : libro.getSheetAt(0)) {
                    if (fila.getRowNum() == 0 || fila.getCell(0) == null) continue;
                    String numero = dataFormatter.formatCellValue(fila.getCell(0)).trim();
                    String proveedor = nombreProveedor(dataFormatter.formatCellValue(fila.getCell(4)).trim());
                    String fecha = dataFormatter.formatCellValue(fila.getCell(1)).trim();
                    String solicitante = dataFormatter.formatCellValue(fila.getCell(6)).trim();
                    String proyecto = dataFormatter.formatCellValue(fila.getCell(7)).trim();
                    String totalOrden = dataFormatter.formatCellValue(fila.getCell(21)).trim();
                    String estado = dataFormatter.formatCellValue(fila.getCell(23)).trim();
                    java.util.Date fechaOrden = leerFechaCelda(fila.getCell(1));
                    String datos = (numero + " " + proveedor + " " + solicitante + " " + proyecto).toLowerCase(java.util.Locale.ROOT);
                    if (!texto.isEmpty() && !datos.contains(texto)) continue;
                    if (inicio != null && (fechaOrden == null || fechaOrden.before(inicio))) continue;
                    if (fin != null && (fechaOrden == null || fechaOrden.after(fin))) continue;
                    if (!"Todos".equals(estadoSeleccionado)
                            && !estado.equalsIgnoreCase(estadoSeleccionado)) continue;
                    modelo.addRow(new Object[]{numero, fecha, proveedor, solicitante, proyecto,
                        totalOrden, estado.isEmpty() ? "Borrador" : estado, "VER"});
                    if ("PENDIENTE".equalsIgnoreCase(estado)) pendientes++;
                    if ("EN PROCESO".equalsIgnoreCase(estado)) enProceso++;
                    try {
                        importeMes += Double.parseDouble(totalOrden.replace("$", "").replace(",", "").trim());
                    } catch (NumberFormatException ignored) {
                        logger.fine("El total de la orden no es numérico: " + totalOrden);
                    }
                }
            } catch (IOException ex) {
                logger.log(java.util.logging.Level.WARNING, "No se pudieron cargar las órdenes para búsqueda.", ex);
                JOptionPane.showMessageDialog(this, "No se pudieron cargar las órdenes:\n" + ex.getMessage(),
                        "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
        total.setText("Total: " + modelo.getRowCount() + " registros");
        totalResumen.setText(String.valueOf(modelo.getRowCount()));
        pendientesResumen.setText(String.valueOf(pendientes));
        procesoResumen.setText(String.valueOf(enProceso));
        importeResumen.setText(String.format(java.util.Locale.US, "$ %,.2f", importeMes));
    }

    private String nombreProveedor(String id) {
        for (String[] proveedor : cacheProveedores) {
            if (proveedor[0].equalsIgnoreCase(id)) return proveedor[1];
        }
        return id;
    }

    private void abrirNuevaOrden() {
        mostrarFormulario();
        estadoOrdenCargada = null;
        salirDeOrden();
        setModoSoloLectura(false);
        contenedorVista.revalidate();
        contenedorVista.repaint();
    }

    private void mostrarFormulario() {
        mostrarVistaInterna(desplazamientoFormulario);
        setTitle("NUEVA ORDEN DE COMPRA");
        setSize(962, 735);
        contenedorVista.revalidate();
        contenedorVista.repaint();
    }

    private void abrirOrdenDesdeBusqueda(String numero, boolean editar) {
        cargarOrdenPorNumero(Integer.parseInt(numero));
        mostrarVistaInterna(desplazamientoFormulario);
        setTitle("ORDEN DE COMPRA " + numero);
        setSize(962, 735);
        if (editar) activarEdicionOrden();
        revalidate();
        repaint();
    }

    private void posicionar(java.awt.Component componente, int x, int y, int ancho, int alto) {
        componente.setBounds(x + DESPLAZAMIENTO_X, y + DESPLAZAMIENTO_Y, ancho, alto);
        if (componente instanceof javax.swing.JLabel etiqueta
                && etiqueta.getText() != null
                && (etiqueta.getText().endsWith(":") || etiqueta.getText().trim().equals("NO."))) {
            etiqueta.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        }
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
     // =====================================================================
     // Componentes y tabla de productos
     // =====================================================================

     private void initComponents() {

        jPanel1 = crearPanelConFondo();
        jLabel1 = new javax.swing.JLabel();
        jButton6 = new javax.swing.JButton();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jTextField1 = new javax.swing.JTextField();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jDateChooser2 = new com.toedter.calendar.JDateChooser();
        jLabel11 = new javax.swing.JLabel();
        jTextField4 = new javax.swing.JTextField();
        jLabel13 = new javax.swing.JLabel();
        jTextField2 = new javax.swing.JTextField();
        jComboBox1 = new javax.swing.JComboBox<>();
        jLabel14 = new javax.swing.JLabel();
        jLabel15 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        jLabel16 = new javax.swing.JLabel();
        jTextField3 = new javax.swing.JTextField();
        jButton7 = new javax.swing.JButton();
        jLabel17 = new javax.swing.JLabel();
        jComboBox3 = new javax.swing.JComboBox<>();
        jLabel18 = new javax.swing.JLabel();
        jLabel19 = new javax.swing.JLabel();
        jLabel20 = new javax.swing.JLabel();
        jLabel21 = new javax.swing.JLabel();
        jLabel22 = new javax.swing.JLabel();
        jButton1 = new javax.swing.JButton();
        jLabel23 = new javax.swing.JLabel();
        jLabel24 = new javax.swing.JLabel();
        jLabel25 = new javax.swing.JLabel();
        jLabel26 = new javax.swing.JLabel();
        jLabel29 = new javax.swing.JLabel();
        jLabel30 = new javax.swing.JLabel();
        jLabel31 = new javax.swing.JLabel();
        jLabel32 = new javax.swing.JLabel();
        jLabel12 = new javax.swing.JLabel();
        jLabel34 = new javax.swing.JLabel();
        jLabel36 = new javax.swing.JLabel();
        jComboBox2 = new javax.swing.JComboBox<>();
        jLabel37 = new javax.swing.JLabel();
        jTextField5 = new javax.swing.JTextField();
        jLabel38 = new javax.swing.JLabel();
        jLabel39 = new javax.swing.JLabel();
        jLabel40 = new javax.swing.JLabel();
        jComboBox4 = new javax.swing.JComboBox<>();
        jComboBox5 = new javax.swing.JComboBox<>();
        jButton2 = new javax.swing.JButton();
        jTextField6 = new javax.swing.JTextField();
        jComboBox6 = new javax.swing.JComboBox<>();
        jComboBox7 = new javax.swing.JComboBox<>();
        jComboBox8 = new javax.swing.JComboBox<>();
        jDateChooserEntrega = new com.toedter.calendar.JDateChooser();
        jDateChooserPago = new com.toedter.calendar.JDateChooser();
        jLabel27 = new javax.swing.JLabel();
        jButton3 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(255, 255, 255));

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(0, 95, 131));
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("ORDEN DE COMPRA");

        jButton6.setBackground(new java.awt.Color(0, 95, 131));
        jButton6.setForeground(new java.awt.Color(255, 255, 255));
        jButton6.setText("< ATRAS");
        jButton6.addActionListener(this::jButton6ActionPerformed);

        jLabel2.setForeground(new java.awt.Color(0, 95, 131));
        jLabel2.setText("PROVEEDOR:");
        jLabel2.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));

        jLabel3.setForeground(new java.awt.Color(0, 95, 131));
        jLabel3.setText("FECHA:");

        jLabel4.setForeground(new java.awt.Color(0, 95, 131));
        jLabel4.setText("AREA:");

        jLabel6.setForeground(new java.awt.Color(0, 95, 131));
        jLabel6.setText("ID_PROVEEDOR:");

        jLabel7.setForeground(new java.awt.Color(0, 95, 131));
        jLabel7.setText("SOLICITANTE:");

        jLabel9.setForeground(new java.awt.Color(0, 95, 131));
        jLabel9.setText("NO.");

        jLabel11.setForeground(new java.awt.Color(0, 95, 131));
        jLabel11.setText("COTIZACIÓN:");

        jLabel13.setForeground(new java.awt.Color(0, 95, 131));
        jLabel13.setText("PROYECTO:");

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jLabel14.setForeground(new java.awt.Color(0, 95, 131));
        jLabel14.setText("DIRECCIÓN:");

        jLabel15.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        jLabel15.setVerticalAlignment(javax.swing.SwingConstants.TOP);
        jLabel15.setHorizontalTextPosition(javax.swing.SwingConstants.RIGHT);

        jLabel5.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel5.setText("FO-SG-04-00");

        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null},
                {null, null, null, null, null, null}
            },
            new String [] {
                "Código", "Descripción", "Unidad de Medida", "Cantidad", "Precio", "Importe"
            }
        ) {
            Class<?>[] types = new Class<?>[] {
                java.lang.Double.class, java.lang.String.class, java.lang.String.class, java.lang.Integer.class, java.lang.Float.class, java.lang.Float.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, true, true, false, false
            };

            public Class<?> getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jScrollPane1.setViewportView(jTable1);

        jLabel16.setForeground(new java.awt.Color(0, 95, 131));
        jLabel16.setText("PRODUCTO:");

        jTextField3.addActionListener(this::jTextField3ActionPerformed);

        jButton7.setBackground(new java.awt.Color(0, 95, 131));
        jButton7.setForeground(new java.awt.Color(255, 255, 255));
        jButton7.setText("Eliminar");

        jLabel17.setForeground(new java.awt.Color(0, 95, 131));
        jLabel17.setText("USO CFDI:");

        jComboBox3.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jLabel18.setForeground(new java.awt.Color(0, 95, 131));
        jLabel18.setText("PAGO:");

        jLabel19.setForeground(new java.awt.Color(0, 95, 131));
        jLabel19.setText("METODO DE PAGO:");

        jLabel20.setForeground(new java.awt.Color(0, 95, 131));
        jLabel20.setText("SUB TOTAL:");

        jLabel21.setForeground(new java.awt.Color(0, 95, 131));
        jLabel21.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        jLabel21.setText("IVA 16%: ");

        jLabel22.setForeground(new java.awt.Color(0, 95, 131));
        jLabel22.setText("TOTAL:");

        jButton1.setBackground(new java.awt.Color(0, 95, 131));
        jButton1.setForeground(new java.awt.Color(255, 255, 255));
        jButton1.setText("IMPRIMIR PDF");
        jButton1.addActionListener(this::jButton1ActionPerformed);

        jLabel23.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel23.setText("jLabel23");

        jLabel24.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel24.setText("jLabel24");

        jLabel25.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel25.setText("jLabel25");

        jLabel29.setForeground(new java.awt.Color(0, 95, 131));
        jLabel29.setText("RFC:");

        jLabel31.setForeground(new java.awt.Color(0, 95, 131));
        jLabel31.setText("CONTACTO:");

        jLabel12.setForeground(new java.awt.Color(0, 95, 131));
        jLabel12.setText("FECHA DE ENTREGA:");

        jLabel34.setForeground(new java.awt.Color(0, 95, 131));
        jLabel34.setText("FORMA DE PAGO:");

        jLabel36.setForeground(new java.awt.Color(0, 95, 131));
        jLabel36.setText("TIPO DE MONEDA:");

        jComboBox2.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jLabel37.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);

        jTextField5.setHorizontalAlignment(javax.swing.JTextField.CENTER);

        jLabel38.setForeground(new java.awt.Color(0, 95, 131));
        jLabel38.setText("TIPO DE CAMBIO:");

        jLabel39.setForeground(new java.awt.Color(0, 95, 131));
        jLabel39.setText("ELABORÓ:");

        jLabel40.setForeground(new java.awt.Color(0, 95, 131));
        jLabel40.setText("AUTORIZÓ:");

        jComboBox4.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jComboBox5.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jButton2.setBackground(new java.awt.Color(0, 204, 51));
        jButton2.setForeground(new java.awt.Color(255, 255, 255));
        jButton2.setText("BUSCAR");

        jComboBox6.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jComboBox7.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jComboBox8.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jLabel27.setForeground(new java.awt.Color(0, 95, 131));
        jLabel27.setText("FECHA DE PAGO:");

        java.util.Locale localeFechas = java.util.Locale.forLanguageTag("es-MX");
        jDateChooser2.setLocale(localeFechas);
        jDateChooserEntrega.setLocale(localeFechas);
        jDateChooserPago.setLocale(localeFechas);
        jDateChooser2.setDateFormatString("d MMM yyyy");
        jDateChooserEntrega.setDateFormatString("d MMM yyyy");
        jDateChooserPago.setDateFormatString("d MMM yyyy");

        jButton3.setBackground(new java.awt.Color(255, 102, 0));
        jButton3.setForeground(new java.awt.Color(255, 255, 255));
        jButton3.setText("CREAR");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(22, 22, 22)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGap(9, 9, 9)
                                        .addComponent(jLabel19)
                                        .addGap(18, 18, 18)
                                        .addComponent(jComboBox8, javax.swing.GroupLayout.PREFERRED_SIZE, 239, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanel1Layout.createSequentialGroup()
                                        .addGap(55, 55, 55)
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                            .addComponent(jLabel13)
                                            .addComponent(jLabel18))
                                        .addGap(18, 18, 18)
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                            .addComponent(jTextField2)
                                            .addComponent(jComboBox6, 0, 238, Short.MAX_VALUE))))
                                .addGap(66, 66, 66)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addComponent(jLabel34)
                                    .addComponent(jLabel12))
                                .addGap(18, 18, 18)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addComponent(jComboBox7, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addGap(38, 38, 38))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addComponent(jDateChooserEntrega, javax.swing.GroupLayout.PREFERRED_SIZE, 153, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                        .addComponent(jLabel27)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(jDateChooserPago, javax.swing.GroupLayout.PREFERRED_SIZE, 153, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(0, 0, Short.MAX_VALUE))))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGap(38, 38, 38)
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                                .addComponent(jLabel7)
                                                .addGap(315, 315, 315))
                                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                                    .addGap(50, 50, 50)
                                                    .addComponent(jLabel26, javax.swing.GroupLayout.PREFERRED_SIZE, 297, javax.swing.GroupLayout.PREFERRED_SIZE))
                                                .addGroup(jPanel1Layout.createSequentialGroup()
                                                    .addComponent(jLabel4)
                                                    .addGap(315, 315, 315)))))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGap(22, 22, 22)
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addComponent(jLabel9)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                                .addComponent(jTextField6, javax.swing.GroupLayout.PREFERRED_SIZE, 98, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                .addGap(148, 148, 148)
                                                .addComponent(jLabel11)
                                                .addGap(18, 18, 18)
                                                .addComponent(jTextField4, javax.swing.GroupLayout.PREFERRED_SIZE, 231, javax.swing.GroupLayout.PREFERRED_SIZE))
                                            .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addComponent(jLabel6)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, 89, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                .addGap(18, 18, 18)
                                                .addComponent(jLabel3)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                                .addComponent(jDateChooser2, javax.swing.GroupLayout.PREFERRED_SIZE, 159, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                                .addComponent(jLabel17)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                                .addComponent(jComboBox3, javax.swing.GroupLayout.PREFERRED_SIZE, 239, javax.swing.GroupLayout.PREFERRED_SIZE))))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addGap(40, 40, 40)
                                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                                    .addComponent(jLabel2)
                                                    .addComponent(jLabel14)))
                                            .addComponent(jLabel31, javax.swing.GroupLayout.Alignment.TRAILING))
                                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addGap(412, 412, 412)
                                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                        .addComponent(jLabel37, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 272, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                        .addGroup(jPanel1Layout.createSequentialGroup()
                                                            .addComponent(jLabel36)
                                                            .addGap(18, 18, 18)
                                                            .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, 153, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                                        .addComponent(jLabel38)
                                                        .addGap(18, 18, 18)
                                                        .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, 151, javax.swing.GroupLayout.PREFERRED_SIZE))))
                                            .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addGap(18, 18, 18)
                                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                                        .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 320, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                        .addGap(114, 114, 114)
                                                        .addComponent(jLabel29)
                                                        .addGap(17, 17, 17)
                                                        .addComponent(jLabel30, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE))
                                                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                                        .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 297, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                        .addComponent(jLabel32, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.PREFERRED_SIZE, 297, javax.swing.GroupLayout.PREFERRED_SIZE))))
                                            .addGroup(jPanel1Layout.createSequentialGroup()
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                                .addComponent(jLabel15, javax.swing.GroupLayout.PREFERRED_SIZE, 666, javax.swing.GroupLayout.PREFERRED_SIZE))))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGap(90, 90, 90)
                                        .addComponent(jButton6)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 674, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(jPanel1Layout.createSequentialGroup()
                                        .addGap(9, 9, 9)
                                        .addComponent(jLabel16)
                                        .addGap(18, 18, 18)
                                        .addComponent(jTextField3, javax.swing.GroupLayout.PREFERRED_SIZE, 212, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(jButton7)))
                                .addGap(0, 53, Short.MAX_VALUE))))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jScrollPane1))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(28, 28, 28)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 105, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 105, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                    .addComponent(jLabel40)
                                    .addComponent(jLabel39))
                                .addGap(18, 18, 18)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addComponent(jComboBox4, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(jComboBox5, javax.swing.GroupLayout.PREFERRED_SIZE, 275, javax.swing.GroupLayout.PREFERRED_SIZE))))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 307, Short.MAX_VALUE)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel20, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel21, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel22, javax.swing.GroupLayout.Alignment.TRAILING))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jLabel24, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, 99, Short.MAX_VALUE)
                            .addComponent(jLabel23, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jLabel25, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addGap(14, 14, 14))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jLabel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(70, 70, 70)
                .addComponent(jLabel5)
                .addGap(9, 9, 9)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel1)
                    .addComponent(jButton6))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(jLabel11)
                                .addComponent(jTextField4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(jLabel9)
                                .addComponent(jTextField6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addComponent(jLabel6))
                            .addComponent(jLabel3)
                            .addComponent(jDateChooser2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(jLabel17)
                        .addComponent(jComboBox3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(13, 13, 13)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                        .addComponent(jLabel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                        .addComponent(jLabel29, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel30, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel14)
                    .addComponent(jLabel15, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel36)
                            .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel37, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(12, 12, 12)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel38))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel34)
                            .addComponent(jComboBox7, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jLabel31, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jLabel32, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel7)
                            .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel4)
                            .addComponent(jLabel26, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel13))
                        .addGap(9, 9, 9)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel18)
                            .addComponent(jComboBox6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel19)
                    .addComponent(jComboBox8, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel12)
                    .addComponent(jDateChooserEntrega, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel27)
                    .addComponent(jDateChooserPago, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel16)
                    .addComponent(jTextField3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton7))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 172, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel20)
                            .addComponent(jLabel23))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel21)
                            .addComponent(jLabel24))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel22)
                            .addComponent(jLabel25))
                        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(12, 12, 12)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel39)
                            .addComponent(jComboBox4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel40)
                            .addComponent(jComboBox5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 36, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 36, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 36, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addContainerGap(12, Short.MAX_VALUE))))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton6ActionPerformed
        volverAOrdenes();
    }//GEN-LAST:event_jButton6ActionPerformed

    /** Regresa a la misma vista general que se muestra al elegir Órdenes de compra. */
    private void volverAOrdenes() {
        mostrarBusquedaOrdenes();
    }

    // =====================================================================
    // Catalogos, autocompletado y calculo de totales
    // =====================================================================

   

    public void buscarProveedor(String idBuscado) {
        try {
            File archivo = AppConfig.proveedoresFile();
            try (FileInputStream file = new FileInputStream(archivo);
                 Workbook workbook = new XSSFWorkbook(file)) {
                Sheet sheet = workbook.getSheetAt(0);

                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }

                    org.apache.poi.ss.usermodel.Cell celdaId = row.getCell(0);
                    if (celdaId != null) {
                        String valorCelda = celdaId.toString();
                        if (valorCelda.equalsIgnoreCase(idBuscado)) {
                            if (!proveedorActivo(row)) {
                                limpiarDatosProveedor("Proveedor inactivo");
                                return;
                            }

                            String nombre = obtenerTextoCelda(row.getCell(1));
                            String RFC = obtenerTextoCelda(row.getCell(2));
                            String telefono = obtenerTextoCelda(row.getCell(3));
                            String correo = obtenerTextoCelda(row.getCell(4));
                            String direccion = obtenerTextoCelda(row.getCell(5));
                            String contacto = obtenerTextoCelda(row.getCell(6));

                            proveedorTelefono = telefono;
                            proveedorCorreo = correo;
                            jLabel8.setText(nombre);
                            jLabel15.setText("<html>" + direccion + "</html>");
                            jLabel30.setText("<html>" + RFC + "</html>");
                            jLabel32.setText("<html>" + contacto + "</html>");
                            return;
                        }
                    }
                }

                jLabel8.setText("Proveedor no encontrado");
            }
        } catch (Exception e) {
            e.printStackTrace();
            jLabel8.setText("Error");
        }
    }

    private void configurarTablaProductos() {
    jTable1.setModel(new DefaultTableModel(
        new Object[][] {},
        new String[] {"Código", "Descripción", "Unidad de Medida", "Cantidad", "Precio", "Importe"}
    ) {
        @Override
        public Class<?> getColumnClass(int columnIndex) {
            return switch (columnIndex) {
                case 0, 1, 2 -> String.class;
                case 3 -> Integer.class;
                case 4, 5 -> Double.class;
                default -> Object.class;
            };
        }

        @Override
        public boolean isCellEditable(int rowIndex, int columnIndex) {
            return columnIndex == 3;
        }

        @Override
        public void setValueAt(Object aValue, int row, int column) {
            if (column == 3) {
                int cantidad = parseCantidad(aValue);
                super.setValueAt(cantidad, row, 3);

                Object precioObj = getValueAt(row, 4);
                double precio = parseNumero(precioObj);
                super.setValueAt(precio * cantidad, row, 5);
                ORDENCOMPRA.this.actualizarTotales();
                return;
            }

            super.setValueAt(aValue, row, column);
        }
    });

    javax.swing.table.TableCellRenderer rendererTextoAjustable = new javax.swing.table.TableCellRenderer() {
        private final javax.swing.JTextArea areaTexto = new javax.swing.JTextArea();

        @Override
        public java.awt.Component getTableCellRendererComponent(
                javax.swing.JTable tabla, Object valor, boolean seleccionado, boolean tieneFoco,
                int fila, int columna) {
            areaTexto.setText(valor == null ? "" : valor.toString());
            areaTexto.setLineWrap(true);
            areaTexto.setWrapStyleWord(true);
            areaTexto.setFont(tabla.getFont());
            areaTexto.setForeground(seleccionado ? tabla.getSelectionForeground() : tabla.getForeground());
            areaTexto.setBackground(seleccionado ? tabla.getSelectionBackground() : tabla.getBackground());
            areaTexto.setBorder(javax.swing.BorderFactory.createEmptyBorder(2, 4, 2, 4));

            int ancho = tabla.getColumnModel().getColumn(columna).getWidth();
            areaTexto.setSize(Math.max(ancho, 1), Short.MAX_VALUE);
            int alturaNecesaria = areaTexto.getPreferredSize().height;
            if (tabla.getRowHeight(fila) < alturaNecesaria) {
                tabla.setRowHeight(fila, alturaNecesaria);
            }
            return areaTexto;
        }
    };
    jTable1.getColumnModel().getColumn(1).setCellRenderer(rendererTextoAjustable);
    jTable1.getColumnModel().getColumn(2).setCellRenderer(rendererTextoAjustable);
    int[] anchosColumnas = {105, 225, 135, 100, 125, 140};
    for (int columna = 0; columna < anchosColumnas.length; columna++) {
        javax.swing.table.TableColumn tableColumn = jTable1.getColumnModel().getColumn(columna);
        tableColumn.setPreferredWidth(anchosColumnas[columna]);
        tableColumn.setMinWidth(anchosColumnas[columna]);
    }
    jTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
}

    private int parseCantidad(Object valor) {
    if (valor == null) {
        return 1;
    }

    if (valor instanceof Number numero) {
        return Math.max(1, numero.intValue());
    }

    try {
        return Math.max(1, Integer.parseInt(valor.toString().trim()));
    } catch (NumberFormatException e) {
        return 1;
    }
}

    private double parseNumero(Object valor) {
    if (valor == null) {
        return 0;
    }

    if (valor instanceof Number numero) {
        return numero.doubleValue();
    }

    try {
        return Double.parseDouble(valor.toString().trim().replace(",", ""));
    } catch (NumberFormatException e) {
        return 0;
    }
}

    private void cargarCacheInventario() {
        cacheInventario.clear();
        try {
            for (String[] item : ORDEN_COMPRA_REPOSITORY.cargarCatalogo(new File(RUTA_INVENTARIOS), 0, 1)) {
                if (!item[0].isEmpty()) {
                    cacheInventario.add(item);
                }
            }
        } catch (IOException e) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo cargar caché de inventario", e);
        }
    }

    private void cargarCacheProveedores() {
        cacheProveedores.clear();
        File archivo = AppConfig.proveedoresFile();
        try (FileInputStream file = new FileInputStream(archivo);
             Workbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                String id     = obtenerTextoCelda(row.getCell(0));
                String nombre = obtenerTextoCelda(row.getCell(1));
                if (!id.isEmpty()) {
                    cacheProveedores.add(new String[]{id, nombre, proveedorActivo(row) ? "1" : "0"});
                }
            }
        } catch (Exception e) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo cargar caché de proveedores", e);
        }
    }

    private boolean proveedorActivo(Row row) {
        return "1".equals(ProveedorService.normalizarEstado(obtenerTextoCelda(row.getCell(11))));
    }

    private void limpiarDatosProveedor(String mensaje) {
        proveedorTelefono = "";
        proveedorCorreo = "";
        jLabel8.setText(mensaje);
        jLabel15.setText("");
        jLabel30.setText("");
        jLabel32.setText("");
    }

    private void filtrarProveedores() {
        if (seleccionandoProveedor || !modoCrear) return;
        String texto = jTextField1.getText().trim().toLowerCase();
        popupProveedores.setVisible(false);
        popupProveedores.removeAll();
        if (texto.isEmpty()) return;

        int count = 0;
        for (String[] prov : cacheProveedores) {
            if (prov[0].toLowerCase().contains(texto) || prov[1].toLowerCase().contains(texto)) {
                String etiqueta = prov[0] + "  |  " + prov[1];
                javax.swing.JMenuItem item = new javax.swing.JMenuItem(etiqueta);
                boolean activo = prov.length < 3 || "1".equals(prov[2]);
                if (!activo) {
                    item.setForeground(java.awt.Color.GRAY);
                    item.setEnabled(false);
                }
                final String id = prov[0];
                item.addActionListener(ev -> {
                    seleccionandoProveedor = true;
                    popupProveedores.setVisible(false);
                    jTextField1.setText(id);
                    buscarProveedor(id);
                    seleccionandoProveedor = false;
                });
                popupProveedores.add(item);
                if (++count == 10) break; // máximo 10 sugerencias
            }
        }
        if (count > 0) {
            popupProveedores.show(jTextField1, 0, jTextField1.getHeight());
            jTextField1.requestFocusInWindow();
        }
    }

    private void filtrarProductos() {
        String texto = jTextField3.getText().trim().toLowerCase();
        popupProductos.setVisible(false);
        popupProductos.removeAll();
        if (texto.isEmpty()) return;

        int count = 0;
        for (String[] prod : cacheInventario) {
            if (prod[0].toLowerCase().contains(texto) || prod[1].toLowerCase().contains(texto)) {
                String etiqueta = prod[0] + "  |  " + prod[1];
                javax.swing.JMenuItem item = new javax.swing.JMenuItem(etiqueta);
                final String codigo = prod[0];
                item.addActionListener(ev -> {
                    popupProductos.setVisible(false);
                    jTextField3.setText(codigo);
                    procesarCodigoProducto();
                });
                popupProductos.add(item);
                if (++count == 10) break; // máximo 10 sugerencias
            }
        }
        if (count > 0) {
            popupProductos.show(jTextField3, 0, jTextField3.getHeight());
            jTextField3.requestFocusInWindow();
        }
    }

    private void procesarCodigoProducto() {
    String codigoBuscado = jTextField3.getText().trim();

    if (codigoBuscado.isEmpty()) {
        return;
    }

    if (buscarProducto(codigoBuscado)) {
        jTextField3.setText("");
    } else if (!ultimoProductoDesactivado) {
        JOptionPane.showMessageDialog(this, "Producto no encontrado en INVENTARIOS.", "Inventarios", JOptionPane.WARNING_MESSAGE);
        jTextField3.requestFocus();
        jTextField3.selectAll();
    }
}
    
    public boolean buscarProducto(String codigoBuscado) {
    ultimoProductoDesactivado = false;
    try (FileInputStream file = new FileInputStream(RUTA_INVENTARIOS);
         Workbook workbook = new XSSFWorkbook(file)) {

        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }

            String codigo = obtenerTextoCelda(row.getCell(0));

            if (codigo.equalsIgnoreCase(codigoBuscado.trim())) {
                String estado = obtenerTextoCelda(row.getCell(5));
                if (!"1".equals(InventarioService.normalizarEstado(estado))) {
                    ultimoProductoDesactivado = true;
                    JOptionPane.showMessageDialog(this,
                            "El producto \"" + codigo + "\" está desactivado y no puede seleccionarse.",
                            "Producto desactivado", JOptionPane.WARNING_MESSAGE);
                    return false;
                }
                String descripcion = obtenerTextoCelda(row.getCell(1));
                String unidad = obtenerTextoCelda(row.getCell(2));
                double precio = obtenerNumeroCelda(row.getCell(3));

                agregarATabla(codigo, descripcion, unidad, precio);
                return true;
            }
        }
    } catch (Exception e) {
        e.printStackTrace();
        JOptionPane.showMessageDialog(this, "No se pudo leer el archivo INVENTARIOS.xlsx", "Error", JOptionPane.ERROR_MESSAGE);
    }

    return false;
}

    private String obtenerTextoCelda(org.apache.poi.ss.usermodel.Cell celda) {
    if (celda == null) {
        return "";
    }

    return dataFormatter.formatCellValue(celda).trim();
}

    private double obtenerNumeroCelda(org.apache.poi.ss.usermodel.Cell celda) {
    if (celda == null) {
        return 0;
    }

    if (celda.getCellType() == CellType.NUMERIC) {
        return celda.getNumericCellValue();
    }

    String valor = dataFormatter.formatCellValue(celda).trim().replace(",", "");
    return valor.isEmpty() ? 0 : Double.parseDouble(valor);
}

    public void agregarATabla(String codigo, String descripcion, String unidad, double precio) {

    DefaultTableModel modelo = (DefaultTableModel) jTable1.getModel();

    int filaExistente = buscarFilaPorCodigo(modelo, codigo);
    if (filaExistente >= 0) {
        JOptionPane.showMessageDialog(this, "El codigo ya existe en la tabla.", "Producto duplicado", JOptionPane.INFORMATION_MESSAGE);
        return;
    }

    Object[] fila = new Object[6];

    fila[0] = codigo;
    fila[1] = descripcion;
    fila[2] = unidad;
    fila[3] = 1; // cantidad
    fila[4] = precio;
    fila[5] = precio * 1;

    modelo.addRow(fila);
    modelo.fireTableDataChanged();
    actualizarTotales();
}

    private void eliminarFilaSeleccionada() {
    int filaSeleccionada = jTable1.getSelectedRow();

    if (filaSeleccionada < 0) {
        JOptionPane.showMessageDialog(this, "Selecciona un producto para eliminar.", "Eliminar", JOptionPane.INFORMATION_MESSAGE);
        return;
    }

    DefaultTableModel modelo = (DefaultTableModel) jTable1.getModel();
    modelo.removeRow(filaSeleccionada);
    actualizarTotales();
}

    private void actualizarTotales() {
    DefaultTableModel modelo = (DefaultTableModel) jTable1.getModel();
    double base = 0;

    for (int i = 0; i < modelo.getRowCount(); i++) {
        int cantidad = parseCantidad(modelo.getValueAt(i, 3));
        double precio = parseNumero(modelo.getValueAt(i, 4));
        double importe = cantidad * precio;

        modelo.setValueAt(importe, i, 5);
        base += importe;
    }

    // Aplicar tipo de cambio si la moneda seleccionada requiere conversión
    Object moneda = jComboBox2.getSelectedItem();
    if (moneda != null && (moneda.toString().equalsIgnoreCase("USD") || moneda.toString().equalsIgnoreCase("EUR") || moneda.toString().equalsIgnoreCase("CAD"))) {
        double tipoCambio = parseNumero(jTextField5.getText());
        if (tipoCambio > 0) {
            base = base / tipoCambio;
        }
    }

    subtotalOriginal = base;

    // Aplicar descuento si el checkbox está activo
    descuentoImporte = 0;
    isrImporte = 0;
    double subtotal = base;
    boolean aplicarDescuento = jCheckBoxDescuento != null && jCheckBoxDescuento.isSelected();
    if (aplicarDescuento) {
        double pct = parseNumero(jTextFieldDescuentoPct.getText());
        if (pct > 0 && pct <= 100) {
            descuentoImporte = base * pct / 100.0;
            subtotal = base - descuentoImporte;
        }
    }

    // Aplicar ISR debajo del descuento, con el mismo comportamiento de porcentaje
    boolean aplicarISR = jCheckBoxISR != null && jCheckBoxISR.isSelected();
    if (aplicarISR) {
        double pctISR = parseNumero(jTextFieldISRPercent.getText());
        if (pctISR > 0 && pctISR <= 100) {
            isrImporte = subtotal * pctISR / 100.0;
            subtotal -= isrImporte;
        }
    }

    // Actualizar texto del label SUB TOTAL con porcentaje si aplica
    if ((aplicarDescuento && descuentoImporte > 0) || (aplicarISR && isrImporte > 0)) {
        double pctDescuento = aplicarDescuento ? parseNumero(jTextFieldDescuentoPct.getText()) : 0;
        double pctISR = aplicarISR ? parseNumero(jTextFieldISRPercent.getText()) : 0;
        jLabel20.setText(String.format("SUB TOTAL (-%.0f%% - ISR %.0f%%):", pctDescuento, pctISR));
    } else {
        jLabel20.setText("SUB TOTAL:");
    }

    // IVA condicional según checkbox
    boolean aplicarIVA = jCheckBoxIVA != null && jCheckBoxIVA.isSelected();
    double iva = aplicarIVA ? subtotal * 0.16 : 0;
    double total = subtotal + iva;

    Object monedaSel = jComboBox2.getSelectedItem();
    String simbolo;
    if (monedaSel == null) {
        simbolo = "$";
    } else {
        switch (monedaSel.toString().toUpperCase()) {
            case "USD": simbolo = "USD $"; break;
            case "EUR": simbolo = "€";     break;
            default:    simbolo = "$";     break;
        }
    }

    jLabel23.setText(String.format(simbolo + " %,.2f", subtotal));
    jLabel24.setText(String.format(simbolo + " %,.2f", iva));
    jLabel25.setText(String.format(simbolo + " %,.2f", total));
}

    private int buscarFilaPorCodigo(DefaultTableModel modelo, String codigo) {
    String codigoNormalizado = codigo == null ? "" : codigo.trim();

    for (int i = 0; i < modelo.getRowCount(); i++) {
        Object codigoFila = modelo.getValueAt(i, 0);
        if (codigoFila != null && codigoFila.toString().trim().equalsIgnoreCase(codigoNormalizado)) {
            return i;
        }
    }

    return -1;
}
    
    public void cargarPagos() {
    try (FileInputStream file = new FileInputStream(RUTA_PAGOS)) {

        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        jComboBox6.removeAllItems();

        boolean primeraFila = true;

        for (Row row : sheet) {
            if (primeraFila) {
                primeraFila = false;
                continue;
            }
            org.apache.poi.ss.usermodel.Cell cell = row.getCell(0);
            if (cell != null) {
                jComboBox6.addItem(cell.toString());
            }
        }

        workbook.close();
        file.close();

    } catch (Exception e) {
        e.printStackTrace();
    }
}

    public void cargarFormaPago() {
    try (FileInputStream file = new FileInputStream(RUTA_FORMAPAGO)) {

        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        jComboBox7.removeAllItems();

        boolean primeraFila = true;

        for (Row row : sheet) {
            if (primeraFila) {
                primeraFila = false;
                continue;
            }
            org.apache.poi.ss.usermodel.Cell cell = row.getCell(0);
            if (cell != null) {
                jComboBox7.addItem(cell.toString());
            }
        }

        workbook.close();
        file.close();

    } catch (Exception e) {
        e.printStackTrace();
    }
}

    public void cargarMetodoPago() {
    try (FileInputStream file = new FileInputStream(RUTA_METODOPAGO)) {

        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        jComboBox8.removeAllItems();

        boolean primeraFila = true;

        for (Row row : sheet) {
            if (primeraFila) {
                primeraFila = false;
                continue;
            }
            org.apache.poi.ss.usermodel.Cell cell = row.getCell(0);
            if (cell != null) {
                jComboBox8.addItem(cell.toString());
            }
        }

        workbook.close();
        file.close();

    } catch (Exception e) {
        e.printStackTrace();
    }
}

    public void cargarCFDI() {
    try (FileInputStream file = new FileInputStream(RUTA_CFDI)) {

        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        jComboBox3.removeAllItems();

        boolean primeraFila = true;

        for (Row row : sheet) {

            if (primeraFila) {
                primeraFila = false; // 👉 saltar encabezado
                continue;
            }

            org.apache.poi.ss.usermodel.Cell cell = row.getCell(0); // columna A

            if (cell != null) {
                jComboBox3.addItem(cell.toString());
            }
        }

        workbook.close();
        file.close();

    } catch (Exception e) {
        e.printStackTrace();
    }
}
     
      public void cargarUsuario() {
    try (FileInputStream User = new FileInputStream(RUTA_USUARIOS)) {

        Workbook workbook = new XSSFWorkbook(User);
        Sheet sheet = workbook.getSheetAt(0);

        jComboBox1.removeAllItems();
        jComboBox4.removeAllItems();
        jComboBox5.removeAllItems();

        boolean primeraFila = true;

        for (Row row : sheet) {

            if (primeraFila) {
                primeraFila = false; // 👉 saltar encabezado
                continue;
            }

            org.apache.poi.ss.usermodel.Cell cell = row.getCell(2); // columna A

            if (cell != null) {
                String valor = cell.toString();
                jComboBox1.addItem(valor);
                jComboBox4.addItem(valor);
                jComboBox5.addItem(valor);
            }
        }

        workbook.close();
        User.close();

    } catch (Exception e) {
        e.printStackTrace();
    }
}
    
    public void cargarMoneda() {
    try (FileInputStream file = new FileInputStream(RUTA_TIPOMONEDA)) {

        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        jComboBox2.removeAllItems();
        monedaDescMap.clear();
        monedaLeyendaMap.clear();

        boolean primeraFila = true;

        for (Row row : sheet) {
            if (primeraFila) {
                primeraFila = false;
                continue;
            }
            org.apache.poi.ss.usermodel.Cell col1 = row.getCell(0);
            org.apache.poi.ss.usermodel.Cell col2 = row.getCell(1);
            org.apache.poi.ss.usermodel.Cell col3 = row.getCell(2);
            if (col1 != null) {
                String clave = col1.toString();
                String desc = col2 != null ? col2.toString() : "";
                String leyenda = col3 != null ? col3.toString() : "";
                jComboBox2.addItem(clave);
                monedaDescMap.put(clave, desc);
                monedaLeyendaMap.put(clave, leyenda);
            }
        }

        workbook.close();
        file.close();

        if (jComboBox2.getItemCount() > 0) {
            jComboBox2.setSelectedIndex(0);
            String desc = monedaDescMap.getOrDefault(jComboBox2.getSelectedItem().toString(), "");
            jLabel37.setText(desc);
        }

    } catch (Exception e) {
        e.printStackTrace();
    }
}

      public void buscarArea(String idBuscado) {
    try (FileInputStream file = new FileInputStream(RUTA_USUARIOS);
         Workbook workbook = new XSSFWorkbook(file)) {
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue;
            org.apache.poi.ss.usermodel.Cell celdaId = row.getCell(2);
            if (celdaId != null) {
                String valorCelda = celdaId.toString();
                if (valorCelda.equalsIgnoreCase(idBuscado)) {
                    String area= row.getCell(3).toString();
                    jLabel26.setText(area);
                   
                    workbook.close();
                    return;
                }
            }
        }
        jLabel26.setText("Proveedor no encontrado");
        workbook.close();
    } catch (Exception e) {
        e.printStackTrace();
        jLabel26.setText("Error");
    }
}
      
    private void jTextField3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField3ActionPerformed
        procesarCodigoProducto();
    }//GEN-LAST:event_jTextField3ActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        generarPDFYFirmar(null);
    }//GEN-LAST:event_jButton1ActionPerformed

    // ─────────────────────────────────────────────────────────────────────────
    //  PDF GENERATION
    // ─────────────────────────────────────────────────────────────────────────
    // =====================================================================
    // Validacion y generacion de documentos PDF
    // =====================================================================

    private boolean generarPDF() {
        String errorValidacion = validarCampos(false);
        if (errorValidacion != null) {
            JOptionPane.showMessageDialog(this, errorValidacion,
                    "Campos incompletos", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        String noOrden = jTextField6.getText().trim();
        String cotizacion = jTextField4.getText().trim();
        String nombreArchivo = "OrdenCompra#" + noOrden + "_" + cotizacion + ".pdf";

        javax.swing.JFileChooser fc = new javax.swing.JFileChooser();
        fc.setDialogTitle("Guardar Orden de Compra");
        fc.setSelectedFile(new File(System.getProperty("user.home") + "\\Desktop\\" + nombreArchivo));
        fc.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("PDF (*.pdf)", "pdf"));
        if (fc.showSaveDialog(this) != javax.swing.JFileChooser.APPROVE_OPTION) return false;

        String ruta = fc.getSelectedFile().getAbsolutePath();
        if (!ruta.toLowerCase().endsWith(".pdf")) ruta += ".pdf";

        File archivoPDF = new File(ruta);
        if (archivoPDF.exists()) {
            int reemplazar = JOptionPane.showConfirmDialog(this,
                    "El PDF '" + archivoPDF.getName() + "' ya existe.\n¿Deseas reemplazarlo?",
                    "Confirmar reemplazo", JOptionPane.YES_NO_OPTION,
                    JOptionPane.QUESTION_MESSAGE);
            if (reemplazar != JOptionPane.YES_OPTION) {
                return false;
            }
        }

        try {
            PdfWriter  writer = new PdfWriter(ruta);
            PdfDocument pdfDoc = new PdfDocument(writer);

            float altoPrimeraPagina = calcularAltoPrimeraPagina();
            PageSize tamanoOrden = new PageSize(PageSize.LETTER.getWidth(), altoPrimeraPagina);
            Document   doc    = new Document(pdfDoc, tamanoOrden);
            doc.setMargins(36, 36, 36, 36);

            PdfFont bold  = PdfFontFactory.createFont(StandardFonts.HELVETICA_BOLD);
            PdfFont plain = PdfFontFactory.createFont(StandardFonts.HELVETICA);
            DeviceRgb azul      = new DeviceRgb(0, 95, 131);
            DeviceRgb separadorColor = new DeviceRgb(42, 210, 201);
            DeviceRgb azulClaro = new DeviceRgb(204, 229, 239);
            DeviceRgb blanco    = new DeviceRgb(255, 255, 255);

            // ── ENCABEZADO ────────────────────────────────────────────────────
            Table header = new Table(UnitValue.createPercentArray(new float[]{25, 50, 25}))
                    .useAllAvailableWidth().setMarginBottom(3);

            Image logo = new Image(ImageDataFactory.create(
                    AppConfig.recursoBytes("IMAGENES/Logo PDF.png"))).scaleToFit(130, 55);
            header.addCell(new Cell().add(logo)
                    .setBorder(Border.NO_BORDER).setPadding(4)
                    .setVerticalAlignment(VerticalAlignment.MIDDLE));

            header.addCell(new Cell()
                    .add(new Paragraph("ORDEN DE COMPRA")
                            .setFont(bold).setFontSize(18).setFontColor(azul)
                            .setTextAlignment(TextAlignment.CENTER))
                    .add(new Paragraph("RED SINERGIA DE TELECOMUNICACIONES, S.A. DE C.V.\nAv. Ocote #10 Arboledas Guadalupe, C.P. 72260. RST140517D48\nTELS.: +52 (01 2222) 278842, (01 55) 7823 1585")
                            .setFont(plain).setFontSize(7).setFontColor(azul)
                            .setTextAlignment(TextAlignment.CENTER))
                    .setBorder(Border.NO_BORDER).setPadding(4)
                    .setVerticalAlignment(VerticalAlignment.MIDDLE));

            header.addCell(new Cell()
                    .add(new Paragraph(jLabel5.getText())
                            .setFont(bold).setFontSize(8)
                            .setTextAlignment(TextAlignment.RIGHT))
                    .add(new Paragraph("NO.  " + jTextField6.getText())
                            .setFont(bold).setFontSize(9).setFontColor(azul)
                            .setTextAlignment(TextAlignment.RIGHT))
                    .add(new Paragraph("Cotización: " + jTextField4.getText())
                            .setFont(plain).setFontSize(9)
                            .setTextAlignment(TextAlignment.RIGHT))
                    .setBorder(Border.NO_BORDER).setPadding(4)
                    .setVerticalAlignment(VerticalAlignment.MIDDLE));
            doc.add(header);

            // Línea separadora azul
            doc.add(new Paragraph("").setBorderBottom(new SolidBorder(separadorColor, 2)).setMarginBottom(3));

            // ── INFORMACIÓN GENERAL ───────────────────────────────────────────
            String fecha = "";
            if (jDateChooser2.getDate() != null) {
                fecha = formatearFecha(jDateChooser2.getDate());
            }
            Object cfdi = jComboBox3.getSelectedItem();

            Table infoGen = new Table(UnitValue.createPercentArray(new float[]{2, 4, 2, 4}))
                    .useAllAvailableWidth().setMarginBottom(2);
            addFilaPDF(infoGen, bold, plain, azul,
                    "FECHA:", fecha,
                    "PROYECTO:", jTextField2.getText());
            doc.add(infoGen);

            // ── DATOS DEL PROVEEDOR ───────────────────────────────────────────
            doc.add(seccionTitulo("DATOS DEL PROVEEDOR", bold, azul));

            Table provTbl = new Table(UnitValue.createPercentArray(new float[]{2, 4, 2, 4}))
                    .useAllAvailableWidth().setMarginBottom(2);
            addFilaPDF(provTbl, bold, plain, azul,
                    "ID PROVEEDOR:", jTextField1.getText(),
                    "PROVEEDOR:",    jLabel8.getText());
            // RFC + CONTACTO en columnas izquierdas, DIRECCIÓN en columnas derechas (una sola fila)
            provTbl.addCell(new Cell()
                    .add(new Paragraph("RFC:").setFont(bold).setFontSize(8).setFontColor(azul))
                    .add(new Paragraph("CONTACTO:").setFont(bold).setFontSize(8).setFontColor(azul).setMarginTop(4))
                    .add(new Paragraph("TELÉFONO:").setFont(bold).setFontSize(8).setFontColor(azul).setMarginTop(4))
                    .setBorder(Border.NO_BORDER).setPadding(3));
            provTbl.addCell(new Cell()
                    .add(new Paragraph(stripHtmlPDF(jLabel30.getText())).setFont(plain).setFontSize(8))
                    .add(new Paragraph(stripHtmlPDF(jLabel32.getText())).setFont(plain).setFontSize(8).setMarginTop(4))
                    .add(new Paragraph(proveedorTelefono).setFont(plain).setFontSize(8).setMarginTop(4))
                    .setBorder(Border.NO_BORDER).setPadding(3).setVerticalAlignment(VerticalAlignment.TOP));
            provTbl.addCell(new Cell()
                    .add(new Paragraph("DIRECCIÓN:").setFont(bold).setFontSize(8).setFontColor(azul))
                    .add(new Paragraph("CORREO:").setFont(bold).setFontSize(8).setFontColor(azul).setMarginTop(25))
                    .setBorder(Border.NO_BORDER).setPadding(3));
            provTbl.addCell(new Cell()
                    .add(new Paragraph(stripHtmlPDF(jLabel15.getText())).setFont(plain).setFontSize(8))
                    .add(new Paragraph(proveedorCorreo).setFont(plain).setFontSize(8).setMarginTop(4))
                    .setBorder(Border.NO_BORDER).setPadding(3).setVerticalAlignment(VerticalAlignment.TOP));
            doc.add(provTbl);

            // ── CONDICIONES COMERCIALES ───────────────────────────────────────
            doc.add(seccionTitulo("CONDICIONES COMERCIALES", bold, azul));

            Table condTbl = new Table(UnitValue.createPercentArray(new float[]{2, 4, 2, 4}))
                    .useAllAvailableWidth().setMarginBottom(2);
            addFilaPDF(condTbl, bold, plain, azul,
                    "PAGO:", jComboBox6.getSelectedItem() != null ? jComboBox6.getSelectedItem().toString() : "",
                    "FECHA DE ENTREGA:", formatearFecha(jDateChooserEntrega.getDate()));
                addFilaPDF(condTbl, bold, plain, azul,
                    "FECHA DE PAGO:", formatearFecha(jDateChooserPago.getDate()),
                    "", "");
            String tipoCambioTexto = jTextField5.getText().trim();
            if (jTextField5.isVisible() && !tipoCambioTexto.isEmpty()) {
                String monedaLabel = jComboBox2.getSelectedItem() != null ? jComboBox2.getSelectedItem().toString().toUpperCase() : "USD";
                addFilaPDF(condTbl, bold, plain, azul,
                        "TIPO DE CAMBIO:", "1 " + monedaLabel + " = " + tipoCambioTexto + " MXN",
                        "", "");
            }
            doc.add(condTbl);

            // ── DATOS DE LA SOLICITUD ─────────────────────────────────────────
            doc.add(seccionTitulo("DATOS DE LA SOLICITUD", bold, azul));

            Object solicitante = jComboBox1.getSelectedItem();
            Table solTbl = new Table(UnitValue.createPercentArray(new float[]{2, 4, 2, 4}))
                    .useAllAvailableWidth().setMarginBottom(2);
            addFilaPDF(solTbl, bold, plain, azul,
                    "SOLICITANTE:", solicitante != null ? solicitante.toString() : "",
                    "ÁREA:",        jLabel26.getText());
            doc.add(solTbl);

            // ── TABLA DE PRODUCTOS ────────────────────────────────────────────
            doc.add(seccionTitulo("PRODUCTOS / SERVICIOS", bold, azul));

            Table prodTbl = new Table(UnitValue.createPercentArray(new float[]{1.5f, 4f, 2f, 1f, 1.5f, 1.5f}))
                    .useAllAvailableWidth().setMarginBottom(2)
                    .setKeepTogether(false);

            String[] colHeaders = {"CÓDIGO", "DESCRIPCIÓN", "UNIDAD DE MEDIDA", "CANT.", "PRECIO UNIT.", "IMPORTE"};
            TextAlignment[] colAligns = {
                TextAlignment.CENTER, TextAlignment.LEFT, TextAlignment.CENTER,
                TextAlignment.CENTER, TextAlignment.RIGHT, TextAlignment.RIGHT
            };
            for (int c = 0; c < colHeaders.length; c++) {
                prodTbl.addHeaderCell(new Cell()
                        .add(new Paragraph(colHeaders[c]).setFont(bold).setFontSize(8)
                                .setFontColor(ColorConstants.WHITE))
                        .setBackgroundColor(azul)
                        .setTextAlignment(colAligns[c])
                        .setPadding(4));
            }

            DefaultTableModel modelo = (DefaultTableModel) jTable1.getModel();
            boolean filaAlterna = false;
            for (int i = 0; i < modelo.getRowCount(); i++) {
                Object cod = modelo.getValueAt(i, 0);
                if (cod == null || cod.toString().trim().isEmpty()) continue;
                DeviceRgb bgFila = filaAlterna ? azulClaro : blanco;
                prodTbl.addCell(celdaTablaPDF(cod,                    plain, TextAlignment.CENTER, bgFila));
                prodTbl.addCell(celdaTablaPDF(modelo.getValueAt(i, 1), plain, TextAlignment.LEFT,   bgFila));
                prodTbl.addCell(celdaTablaPDF(modelo.getValueAt(i, 2), plain, TextAlignment.CENTER, bgFila));
                prodTbl.addCell(celdaTablaPDF(modelo.getValueAt(i, 3), plain, TextAlignment.CENTER, bgFila));
                prodTbl.addCell(celdaMonedaPDF(modelo.getValueAt(i, 4), plain, bgFila));
                prodTbl.addCell(celdaMonedaPDF(modelo.getValueAt(i, 5), plain, bgFila));
                filaAlterna = !filaAlterna;
            }
            doc.add(prodTbl);

            // ── TOTALES ───────────────────────────────────────────────────────
            String monedaPDF = jComboBox2.getSelectedItem() != null ? jComboBox2.getSelectedItem().toString().toUpperCase() : "MXN";
            String simboloPDF;
            switch (monedaPDF) {
                case "USD": simboloPDF = "USD $"; break;
                case "EUR": simboloPDF = "€";     break;
                default:    simboloPDF = "$";     break;
            }
            boolean mostrarDescuento = jCheckBoxDescuento != null && jCheckBoxDescuento.isSelected() && descuentoImporte > 0;
            boolean mostrarISR = jCheckBoxISR != null && jCheckBoxISR.isSelected() && isrImporte > 0;
            int filasTot = 3 + (mostrarDescuento ? 1 : 0) + (mostrarISR ? 1 : 0);
            Table totTbl = new Table(UnitValue.createPercentArray(new float[]{6, 2.5f, 1.5f}))
                    .useAllAvailableWidth().setMarginTop(2)
                    .setKeepTogether(true);
            totTbl.addCell(new Cell(filasTot, 1).setBorder(Border.NO_BORDER));  // espacio izq.

            // Subtotal
            totTbl.addCell(new Cell()
                    .add(new Paragraph("SUB TOTAL:").setFont(bold).setFontSize(8))
                    .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azulClaro)
                    .setBorder(new SolidBorder(azul, 1)).setPadding(3));
            totTbl.addCell(new Cell()
                    .add(new Paragraph(mostrarDescuento ? String.format(simboloPDF + " %,.2f", subtotalOriginal) : jLabel23.getText()).setFont(bold).setFontSize(8))
                    .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azulClaro)
                    .setBorder(new SolidBorder(azul, 1)).setPadding(3));

            // Descuento (si aplica)
            if (mostrarDescuento) {
                double pctDesc = parseNumero(jTextFieldDescuentoPct.getText());
                totTbl.addCell(new Cell()
                        .add(new Paragraph(String.format("DESCUENTO (%.0f%%):", pctDesc)).setFont(bold).setFontSize(8))
                        .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azulClaro)
                        .setBorderLeft(new SolidBorder(azul, 1))
                        .setBorderRight(new SolidBorder(azul, 1))
                        .setBorderTop(Border.NO_BORDER)
                        .setBorderBottom(new SolidBorder(azul, 1)).setPadding(3));
                totTbl.addCell(new Cell()
                        .add(new Paragraph(String.format(simboloPDF + " -%,.2f", descuentoImporte)).setFont(bold).setFontSize(8))
                        .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azulClaro)
                        .setBorderLeft(Border.NO_BORDER)
                        .setBorderRight(new SolidBorder(azul, 1))
                        .setBorderTop(Border.NO_BORDER)
                        .setBorderBottom(new SolidBorder(azul, 1)).setPadding(3));
            }

                    // ISR (si aplica)
                    if (mostrarISR) {
                    totTbl.addCell(new Cell()
                        .add(new Paragraph("ISR:").setFont(bold).setFontSize(8))
                        .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azulClaro)
                        .setBorderLeft(new SolidBorder(azul, 1))
                        .setBorderRight(new SolidBorder(azul, 1))
                        .setBorderTop(Border.NO_BORDER)
                        .setBorderBottom(new SolidBorder(azul, 1)).setPadding(3));
                    totTbl.addCell(new Cell()
                        .add(new Paragraph(String.format(simboloPDF + " -%,.2f", isrImporte)).setFont(bold).setFontSize(8))
                        .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azulClaro)
                        .setBorderLeft(Border.NO_BORDER)
                        .setBorderRight(new SolidBorder(azul, 1))
                        .setBorderTop(Border.NO_BORDER)
                        .setBorderBottom(new SolidBorder(azul, 1)).setPadding(3));
                    }

            // IVA
            totTbl.addCell(new Cell()
                    .add(new Paragraph("IVA 16%:").setFont(bold).setFontSize(8))
                    .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azulClaro)
                    .setBorderLeft(new SolidBorder(azul, 1))
                    .setBorderRight(new SolidBorder(azul, 1))
                    .setBorderTop(Border.NO_BORDER)
                    .setBorderBottom(new SolidBorder(azul, 1)).setPadding(3));
            totTbl.addCell(new Cell()
                    .add(new Paragraph(jLabel24.getText()).setFont(bold).setFontSize(8))
                    .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azulClaro)
                    .setBorderLeft(Border.NO_BORDER)
                    .setBorderRight(new SolidBorder(azul, 1))
                    .setBorderTop(Border.NO_BORDER)
                    .setBorderBottom(new SolidBorder(azul, 1)).setPadding(3));

            // Total
            totTbl.addCell(new Cell()
                    .add(new Paragraph("TOTAL:").setFont(bold).setFontSize(9)
                            .setFontColor(ColorConstants.WHITE))
                    .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azul)
                    .setBorder(new SolidBorder(azul, 1)).setPadding(3));
            totTbl.addCell(new Cell()
                    .add(new Paragraph(jLabel25.getText()).setFont(bold).setFontSize(9)
                            .setFontColor(ColorConstants.WHITE))
                    .setTextAlignment(TextAlignment.RIGHT).setBackgroundColor(azul)
                    .setBorder(new SolidBorder(azul, 1)).setPadding(3));
            doc.add(totTbl);

            // Línea separadora azul (debajo de totales)
            doc.add(new Paragraph("").setBorderBottom(new SolidBorder(separadorColor, 2)).setMarginBottom(3));

            // ── OBSERVACIONES ──────────────────────────────────────────────────
            String textoObservaciones = jTextAreaObservaciones.getText().trim();
            if (!textoObservaciones.isEmpty()) {
                doc.add(new Paragraph("OBSERVACIONES:").setFont(bold).setFontSize(8).setFontColor(azul).setMarginBottom(2));
                doc.add(new Paragraph(textoObservaciones).setFont(plain).setFontSize(8).setMarginBottom(4));
                doc.add(new Paragraph("").setBorderBottom(new SolidBorder(separadorColor, 2)).setMarginBottom(3));
            }

            // ── INFORMACIÓN DE FACTURA ────────────────────────────────────────
            doc.add(new Paragraph("La generación de la factura debe incluir la siguiente información:")
                    .setFont(bold).setFontSize(8).setFontColor(azul).setMarginBottom(2));
            Table factTbl = new Table(UnitValue.createPercentArray(new float[]{2, 4}))
                    .useAllAvailableWidth().setMarginBottom(4);
            factTbl.addCell(new Cell()
                    .add(new Paragraph("USO CFDI:").setFont(bold).setFontSize(8).setFontColor(azul))
                    .setBorder(Border.NO_BORDER).setPadding(2));
            factTbl.addCell(new Cell()
                    .add(new Paragraph(cfdi != null ? cfdi.toString() : "").setFont(plain).setFontSize(8))
                    .setBorder(Border.NO_BORDER).setPadding(2));
            factTbl.addCell(new Cell()
                    .add(new Paragraph("FORMA DE PAGO:").setFont(bold).setFontSize(8).setFontColor(azul))
                    .setBorder(Border.NO_BORDER).setPadding(2));
            factTbl.addCell(new Cell()
                    .add(new Paragraph(jComboBox7.getSelectedItem() != null ? jComboBox7.getSelectedItem().toString() : "").setFont(plain).setFontSize(8))
                    .setBorder(Border.NO_BORDER).setPadding(2));
            factTbl.addCell(new Cell()
                    .add(new Paragraph("MÉTODO DE PAGO:").setFont(bold).setFontSize(8).setFontColor(azul))
                    .setBorder(Border.NO_BORDER).setPadding(2));
            factTbl.addCell(new Cell()
                    .add(new Paragraph(jComboBox8.getSelectedItem() != null ? jComboBox8.getSelectedItem().toString() : "").setFont(plain).setFontSize(8))
                    .setBorder(Border.NO_BORDER).setPadding(2));
            factTbl.addCell(new Cell()
                    .add(new Paragraph("TIPO DE MONEDA:").setFont(bold).setFontSize(8).setFontColor(azul))
                    .setBorder(Border.NO_BORDER).setPadding(2));
            String monedaSeleccionada = jComboBox2.getSelectedItem() != null ? jComboBox2.getSelectedItem().toString() : "";
            String descMoneda = monedaDescMap.getOrDefault(monedaSeleccionada, "");
            String monedaTexto = descMoneda.isEmpty() ? monedaSeleccionada : monedaSeleccionada + " - " + descMoneda;
            String leyendaMoneda = monedaLeyendaMap.getOrDefault(monedaSeleccionada, "");
            Cell celdaMonedaVal = new Cell()
                    .setBorder(Border.NO_BORDER).setPadding(2);
            celdaMonedaVal.add(new Paragraph(monedaTexto).setFont(plain).setFontSize(8));
            if (!leyendaMoneda.isEmpty()) {
                celdaMonedaVal.add(new Paragraph(leyendaMoneda).setFont(plain).setFontSize(7)
                        .setFontColor(new DeviceRgb(80, 80, 80)));
            }
            factTbl.addCell(celdaMonedaVal);
            doc.add(factTbl);

            // ── AVISO DE FACTURACIÓN ──────────────────────────────────────────
            doc.add(new Paragraph("Cualquier duda sobre la facturación comunicarse al departamento de contabilidad o al correo contabilidad@rstelecom.com.mx")
                    .setFont(plain).setFontSize(8).setMarginTop(2).setMarginBottom(1));
            doc.add(new Paragraph("El pago de facturas se realizará los días jueves y viernes, siempre y cuando sean enviadas antes de las 12:00 hrs viéndose reflejado el pago en 24 hrs después. En caso de que se envíen después del horario estipulado, el pago de facturas se realizará hasta la siguiente semana.")
                    .setFont(plain).setFontSize(8).setMarginBottom(2));

            // Línea separadora azul (debajo de leyenda)
            doc.add(new Paragraph("").setBorderBottom(new SolidBorder(separadorColor, 2)).setMarginBottom(3));

            // Leyenda de moneda
            doc.add(new Paragraph("Orden de compra realizada en " + monedaTexto)
                    .setFont(plain).setFontSize(8).setMarginBottom(6)
                    .setTextAlignment(TextAlignment.CENTER));

                // ── FIRMAS ─────────────────────────────────────────────────────
                // Las firmas siguen a los totales; si la tabla ocupó otra página,
                // permanecen en esa misma página sin superponerse al contenido.
            float pageWidth  = PageSize.LETTER.getWidth();
            float marginLR   = 36f;
            float tblWidth   = pageWidth - marginLR * 2;
            String elaboro   = jComboBox4.getSelectedItem() != null ? jComboBox4.getSelectedItem().toString() : "";
            String autorizo  = jComboBox5.getSelectedItem() != null ? jComboBox5.getSelectedItem().toString() : "";
            Table firmasTbl = new Table(UnitValue.createPercentArray(new float[]{1, 1}))
                    .useAllAvailableWidth()
                    .setWidth(tblWidth)
                    .setKeepTogether(true)
                    .setMarginTop(57);
            for (String[] par : new String[][]{{"ELABORÓ", elaboro}, {"AUTORIZÓ", autorizo}}) {
                firmasTbl.addCell(new Cell()
                        .add(new Paragraph("___________________________")
                                .setFont(bold).setFontSize(9)
                                .setTextAlignment(TextAlignment.CENTER))
                        .add(new Paragraph(par[0])
                                .setFont(bold).setFontSize(9)
                                .setTextAlignment(TextAlignment.CENTER))
                        .add(new Paragraph(par[1])
                                .setFont(plain).setFontSize(8)
                                .setTextAlignment(TextAlignment.CENTER))
                        .setBorder(Border.NO_BORDER));
            }
            doc.add(firmasTbl);

            // La orden usa una primera página de alto dinámico; los términos
            // comienzan siempre en una hoja carta independiente.
            pdfDoc.setDefaultPageSize(PageSize.LETTER);
            doc.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
            doc.add(seccionTitulo("TÉRMINOS Y CONDICIONES", bold, azul));
            String contenidoTerminos;
            try {
                byte[] bytesT = java.nio.file.Files.readAllBytes(AppConfig.terminosFile().toPath());
                contenidoTerminos = new String(bytesT, java.nio.charset.StandardCharsets.UTF_8);
            } catch (Exception exT) {
                contenidoTerminos = "(No se pudo cargar el archivo de términos y condiciones)";
            }
            doc.add(new Paragraph(contenidoTerminos)
                    .setFont(plain).setFontSize(8)
                    .setMarginTop(6)
                    .setTextAlignment(TextAlignment.JUSTIFIED));

            doc.close();

            ultimoPdfGenerado = archivoPDF;
            actualizarEstadoBotonFirma();
            JOptionPane.showMessageDialog(this,
                    "PDF generado exitosamente.\n"
                    + "Ahora se solicitará la firma digital.",
                    "Listo", JOptionPane.INFORMATION_MESSAGE);
            return true;

        } catch (Exception ex) {
            logger.log(java.util.logging.Level.SEVERE, "Error generando PDF", ex);
            JOptionPane.showMessageDialog(this, "Error al generar el PDF:\n" + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
            return false;
        }
    }

    private boolean generarPDFYFirmar(String rolForzado) {
        DatosFirma datosFirma = seleccionarDatosFirma(rolForzado);
        if (datosFirma == null) {
            return false;
        }
        if (!generarPDF()) {
            return false;
        }
        try {
            ultimoPdfGenerado = firmarArchivoPdf(ultimoPdfGenerado,
                    datosFirma.clavePrivada, datosFirma.cadena,
                    datosFirma.rol, datosFirma.alias);
            actualizarEstadoBotonFirma();
            abrirPdf(ultimoPdfGenerado);
            JOptionPane.showMessageDialog(this,
                    "PDF firmado digitalmente como " + datosFirma.rol + ".\n"
                    + "Se creó el archivo:\n" + ultimoPdfGenerado.getName(),
                    "Firma completada", JOptionPane.INFORMATION_MESSAGE);
            return true;
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING,
                    "No se pudo firmar el PDF generado.", ex);
            JOptionPane.showMessageDialog(this,
                    "No se pudo firmar el PDF:\n" + ex.getMessage(),
                    "Error de firma", JOptionPane.ERROR_MESSAGE);
            return false;
        }
    }

    private DatosFirma seleccionarDatosFirma(String rolForzado) {
        String rolFirma = rolForzado;
        if (rolFirma == null) {
            Object rol = JOptionPane.showInputDialog(this,
                    "Selecciona el rol que firmará el PDF:",
                    "Firma digital", JOptionPane.QUESTION_MESSAGE, null,
                    new String[]{"ELABORÓ", "AUTORIZÓ"}, "ELABORÓ");
            if (rol == null) {
                return null;
            }
            rolFirma = rol.toString();
        }
        if ("AUTORIZÓ".equals(rolFirma)
                && (ultimoPdfGenerado == null || !pdfTieneFirma("Firma_ELABORO"))) {
            JOptionPane.showMessageDialog(this,
                    "Primero debes firmar el PDF como ELABORÓ.",
                    "Orden de firmas", JOptionPane.WARNING_MESSAGE);
            return null;
        }

        try {
            KeyStore almacen = KeyStore.getInstance("Windows-MY");
            almacen.load(null, null);
            java.util.List<String> aliases = new java.util.ArrayList<>();
            java.util.List<String> opciones = new java.util.ArrayList<>();
            java.util.Enumeration<String> nombres = almacen.aliases();
            while (nombres.hasMoreElements()) {
                String alias = nombres.nextElement();
                if (!almacen.isKeyEntry(alias)) {
                    continue;
                }
                java.security.cert.Certificate certificado = almacen.getCertificate(alias);
                if (certificado instanceof X509Certificate x509) {
                    aliases.add(alias);
                    opciones.add(alias + " - " + x509.getSubjectX500Principal().getName());
                }
            }
            if (opciones.isEmpty()) {
                throw new java.security.KeyStoreException(
                        "No se encontraron certificados con clave privada en Windows.");
            }

            Object seleccion = JOptionPane.showInputDialog(this,
                    "Selecciona el certificado que se usará para " + rolFirma + ":",
                    "Certificado digital", JOptionPane.QUESTION_MESSAGE, null,
                    opciones.toArray(), opciones.get(0));
            if (seleccion == null) {
                return null;
            }
            int indice = opciones.indexOf(seleccion.toString());
            String alias = aliases.get(indice);
            Key clave = almacen.getKey(alias, null);
            if (!(clave instanceof PrivateKey privada)) {
                throw new java.security.KeyStoreException(
                        "El certificado seleccionado no tiene una clave privada accesible.");
            }
            return new DatosFirma(privada, almacen.getCertificateChain(alias), rolFirma, alias);
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING,
                    "No se pudo seleccionar el certificado.", ex);
            JOptionPane.showMessageDialog(this,
                    "No se pudo acceder al certificado:\n" + ex.getMessage(),
                    "Error de certificado", JOptionPane.ERROR_MESSAGE);
            return null;
        }
    }

    private void actualizarEstadoBotonFirma() {
        if (jButtonFirmar != null) {
            jButtonFirmar.setEnabled(ultimoPdfGenerado != null
                    && ultimoPdfGenerado.isFile());
        }
    }

    private boolean firmarPdf() {
        return firmarPdf(null);
    }

    private boolean firmarPdf(String rolForzado) {
        if (ultimoPdfGenerado == null || !ultimoPdfGenerado.isFile()) {
            JOptionPane.showMessageDialog(this,
                    "Primero genera el PDF con el botón IMPRIMIR.",
                    "PDF requerido", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        String rolFirma = rolForzado;
        if (rolFirma == null) {
            Object rol = JOptionPane.showInputDialog(this,
                    "Selecciona el rol que firmará el PDF:",
                    "Firma digital", JOptionPane.QUESTION_MESSAGE, null,
                    new String[]{"ELABORÓ", "AUTORIZÓ"}, "ELABORÓ");
            if (rol == null) {
                return false;
            }
            rolFirma = rol.toString();
        }
        if ("AUTORIZÓ".equals(rolFirma) && !pdfTieneFirma("Firma_ELABORO")) {
            JOptionPane.showMessageDialog(this,
                    "Primero debes firmar el PDF como ELABORÓ.",
                    "Orden de firmas", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        try {
            KeyStore almacen = KeyStore.getInstance("Windows-MY");
            almacen.load(null, null);
            java.util.List<String> aliases = new java.util.ArrayList<>();
            java.util.List<String> opciones = new java.util.ArrayList<>();
            java.util.Enumeration<String> nombres = almacen.aliases();
            while (nombres.hasMoreElements()) {
                String alias = nombres.nextElement();
                if (!almacen.isKeyEntry(alias)) {
                    continue;
                }
                java.security.cert.Certificate certificado = almacen.getCertificate(alias);
                if (certificado instanceof X509Certificate x509) {
                    aliases.add(alias);
                    opciones.add(alias + " - " + x509.getSubjectX500Principal().getName());
                }
            }
            if (opciones.isEmpty()) {
                throw new java.security.KeyStoreException(
                        "No se encontraron certificados con clave privada en Windows.");
            }

            Object seleccion = JOptionPane.showInputDialog(this,
                    "Selecciona el certificado que se usará para " + rolFirma + ":",
                    "Certificado digital", JOptionPane.QUESTION_MESSAGE, null,
                    opciones.toArray(), opciones.get(0));
            if (seleccion == null) {
                return false;
            }
            int indice = opciones.indexOf(seleccion.toString());
            String alias = aliases.get(indice);
            Key clave = almacen.getKey(alias, null);
            if (!(clave instanceof PrivateKey privada)) {
                throw new java.security.KeyStoreException(
                        "El certificado seleccionado no tiene una clave privada accesible.");
            }
            java.security.cert.Certificate[] cadena = almacen.getCertificateChain(alias);
            ultimoPdfGenerado = firmarArchivoPdf(ultimoPdfGenerado, privada, cadena, rolFirma, alias);
            actualizarEstadoBotonFirma();
            abrirPdf(ultimoPdfGenerado);
            JOptionPane.showMessageDialog(this,
                    "PDF firmado digitalmente como " + rolFirma + ".\n"
                            + "Se creó un nuevo archivo firmado:\n" + ultimoPdfGenerado.getName(),
                    "Firma completada", JOptionPane.INFORMATION_MESSAGE);
            return true;
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo firmar el PDF.", ex);
            JOptionPane.showMessageDialog(this,
                    "No se pudo firmar el PDF:\n" + ex.getMessage(),
                    "Error de firma", JOptionPane.ERROR_MESSAGE);
            return false;
        }
    }

    private void abrirPdf(File archivo) throws IOException {
        if (!java.awt.Desktop.isDesktopSupported()) {
            throw new IOException("El sistema no permite abrir archivos PDF automáticamente.");
        }
        java.awt.Desktop.getDesktop().open(archivo);
    }

    private File firmarArchivoPdf(File archivo, PrivateKey privada,
            java.security.cert.Certificate[] cadena, String rol, String alias) throws Exception {
        if (Security.getProvider("BC") == null) {
            Security.addProvider(new BouncyCastleProvider());
        }
        File temporal = File.createTempFile("orden-firmada-", ".pdf", archivo.getParentFile());
        File firmado = new File(archivo.getParentFile(),
                quitarExtension(archivo.getName()) + "_firmado_"
                + new java.text.SimpleDateFormat("yyyyMMdd_HHmmss").format(new java.util.Date()) + ".pdf");
        PdfSigner firmante = null;
        try {
            firmante = new PdfSigner(new PdfReader(archivo),
                    new PdfWriter(temporal), new com.itextpdf.kernel.pdf.StampingProperties());
            int pagina;
            try (PdfDocument documento = new PdfDocument(new PdfReader(archivo))) {
                pagina = documento.getNumberOfPages();
            }
            float y = "AUTORIZÓ".equals(rol) ? 36 : 96;
            firmante.setFieldName("Firma_" + ("ELABORÓ".equals(rol) ? "ELABORO" : "AUTORIZO"));
            firmante.getSignatureAppearance()
                    .setReason("Firma digital de orden de compra")
                    .setLocation("Sistema OCXRST")
                    .setPageNumber(pagina)
                    .setPageRect(new Rectangle(36, y, 240, 48))
                    .setLayer2Text(rol + "\nFirmado digitalmente\nCertificado: " + alias);
            IExternalDigest digest = new BouncyCastleDigest();
            IExternalSignature firma = new PrivateKeySignature(privada,
                    DigestAlgorithms.SHA256, "BC");
            firmante.signDetached(digest, firma, cadena, null, null, null, 0,
                    PdfSigner.CryptoStandard.CADES);
            firmante.getDocument().close();
            firmante = null;
            Files.move(temporal.toPath(), firmado.toPath(),
                    StandardCopyOption.REPLACE_EXISTING);
            return firmado;
        } finally {
            if (firmante != null) {
                firmante.getDocument().close();
            }
            if (temporal.isFile()) {
                Files.deleteIfExists(temporal.toPath());
            }
        }
    }

    private String quitarExtension(String nombre) {
        int punto = nombre.lastIndexOf('.');
        return punto > 0 ? nombre.substring(0, punto) : nombre;
    }

    private boolean pdfTieneFirma(String campo) {
        try (PdfDocument documento = new PdfDocument(new PdfReader(ultimoPdfGenerado))) {
            PdfAcroForm formulario = PdfAcroForm.getAcroForm(documento, false);
            return formulario != null && formulario.getFormFields().containsKey(campo);
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.FINE,
                    "No se pudo consultar el estado de firmas del PDF.", ex);
            return false;
        }
    }

    /** Fila de 2 pares etiqueta-valor para tablas de información */
    private void addFilaPDF(Table tbl, PdfFont bold, PdfFont plain, DeviceRgb azul,
                            String lbl1, String val1, String lbl2, String val2) {
        tbl.addCell(new Cell()
                .add(new Paragraph(lbl1).setFont(bold).setFontSize(8).setFontColor(azul))
                .setBorder(Border.NO_BORDER).setPadding(2));
        tbl.addCell(new Cell()
                .add(new Paragraph(val1 != null ? val1 : "").setFont(plain).setFontSize(8))
                .setBorder(Border.NO_BORDER).setPadding(2)
                .setVerticalAlignment(VerticalAlignment.TOP));
        tbl.addCell(new Cell()
                .add(new Paragraph(lbl2).setFont(bold).setFontSize(8).setFontColor(azul))
                .setBorder(Border.NO_BORDER).setPadding(2));
        tbl.addCell(new Cell()
                .add(new Paragraph(val2 != null ? val2 : "").setFont(plain).setFontSize(8))
                .setBorder(Border.NO_BORDER).setPadding(2)
                .setVerticalAlignment(VerticalAlignment.TOP));
    }

    /** Celda de encabezado de sección (fondo azul, texto blanco) */
    private Table seccionTitulo(String titulo, PdfFont bold, DeviceRgb azul) {
        Table t = new Table(UnitValue.createPercentArray(new float[]{1})).useAllAvailableWidth();
        t.addCell(new Cell()
                .add(new Paragraph(titulo).setFont(bold).setFontSize(8)
                        .setFontColor(ColorConstants.WHITE))
                .setBackgroundColor(azul).setPadding(3).setBorder(Border.NO_BORDER));
        return t;
    }

    private Cell celdaTablaPDF(Object val, PdfFont font, TextAlignment align, DeviceRgb bg) {
        String txt = val == null ? "" : val.toString();
        return new Cell()
                .add(new Paragraph(txt).setFont(font).setFontSize(8))
                .setTextAlignment(align).setBackgroundColor(bg).setPadding(2);
    }

    private Cell celdaMonedaPDF(Object val, PdfFont font, DeviceRgb bg) {
        double num = val instanceof Number n ? n.doubleValue() : 0;
        return celdaTablaPDF(String.format("$%,.2f", num), font, TextAlignment.RIGHT, bg);
    }

    private String stripHtmlPDF(String html) {
        if (html == null) return "";
        return html.replaceAll("<[^>]*>", "").trim();
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  VALIDACIÓN DE CAMPOS
    // ─────────────────────────────────────────────────────────────────────────

    /** Verifica que un combo tenga una selección real (no vacía ni placeholder). */
    private boolean esSeleccionValida(javax.swing.JComboBox<?> combo) {
        Object sel = combo.getSelectedItem();
        if (sel == null) return false;
        String txt = sel.toString().trim();
        return !txt.isEmpty()
                && !txt.equals("** Selecciona una opción **")
                && !txt.equals("--");
    }

    /**
     * Valida que todos los campos obligatorios estén completos.
     * @return null si todo es correcto, o un mensaje indicando los campos faltantes.
     */
    private String validarCampos() {
        return validarCampos(true);
    }

    private String validarCampos(boolean validarLimiteFechas) {
        java.util.List<String> faltantes = new java.util.ArrayList<>();

        // ID Proveedor
        if (jTextField1.getText().trim().isEmpty()) {
            faltantes.add("• ID PROVEEDOR");
        } else {
            String nomProv = jLabel8.getText();
            if (nomProv.isEmpty()
                    || nomProv.equals("Proveedor no encontrado")
                    || nomProv.equals("Error")) {
                faltantes.add("• PROVEEDOR (ID no encontrado, verifique el dato)");
            }
        }

        // Fecha
        if (jDateChooser2.getDate() == null) {
            faltantes.add("• FECHA");
        }

        // Cotización
        if (jTextField4.getText().trim().isEmpty()) {
            faltantes.add("• COTIZACIÓN");
        }

        // USO CFDI
        if (!esSeleccionValida(jComboBox3)) {
            faltantes.add("• USO CFDI");
        }

        // Solicitante
        if (!esSeleccionValida(jComboBox1)) {
            faltantes.add("• SOLICITANTE");
        }

        // Proyecto
        if (jTextField2.getText().trim().isEmpty()) {
            faltantes.add("• PROYECTO");
        }

        // Pago
        if (!esSeleccionValida(jComboBox6)) {
            faltantes.add("• PAGO");
        }

        // Forma de pago
        if (!esSeleccionValida(jComboBox7)) {
            faltantes.add("• FORMA DE PAGO");
        }

        // Método de pago
        if (!esSeleccionValida(jComboBox8)) {
            faltantes.add("• MÉTODO DE PAGO");
        }

        // Fechas comerciales
        if (jDateChooserEntrega.getDate() == null) {
            faltantes.add("• FECHA DE ENTREGA");
        }
        if (jDateChooserPago.getDate() == null) {
                faltantes.add("• FECHA DE PAGO");
        }
        if (validarLimiteFechas && !esAdministrador()) {
            if (esFechaAnteriorAHoy(jDateChooserEntrega.getDate())) {
                faltantes.add("• FECHA DE ENTREGA (no puede ser anterior a hoy)");
            }
            if (esFechaAnteriorAHoy(jDateChooserPago.getDate())) {
                faltantes.add("• FECHA DE PAGO (no puede ser anterior a hoy)");
            }
        }

        // Tipo de moneda
        if (!esSeleccionValida(jComboBox2)) {
            faltantes.add("• TIPO DE MONEDA");
        }

        // Tipo de cambio (solo cuando la moneda requiere conversión)
        if (jTextField5.isVisible() && jTextField5.getText().trim().isEmpty()) {
            faltantes.add("• TIPO DE CAMBIO");
        }

        // Elaboró
        if (!esSeleccionValida(jComboBox4)) {
            faltantes.add("• ELABORÓ");
        }

        // Autorizó
        if (!esSeleccionValida(jComboBox5)) {
            faltantes.add("• AUTORIZÓ");
        }

        // Tabla de productos (al menos un producto)
        DefaultTableModel modeloVal = (DefaultTableModel) jTable1.getModel();
        boolean tieneProductos = false;
        for (int i = 0; i < modeloVal.getRowCount(); i++) {
            Object cod = modeloVal.getValueAt(i, 0);
            if (cod != null && !cod.toString().trim().isEmpty()) {
                tieneProductos = true;
                break;
            }
        }
        if (!tieneProductos) {
            faltantes.add("• PRODUCTOS (la tabla debe contener al menos un producto)");
        }

        if (faltantes.isEmpty()) return null;

        return "Por favor, complete los siguientes campos antes de continuar:\n\n"
                + String.join("\n", faltantes);
    }

    private boolean esAdministrador() {
        return "admin".equalsIgnoreCase(INICIARSESION.usuarioActual);
    }

    private boolean esFechaAnteriorAHoy(java.util.Date fecha) {
        if (fecha == null) {
            return false;
        }
        java.time.LocalDate fechaSeleccionada = fecha.toInstant()
                .atZone(java.time.ZoneId.systemDefault()).toLocalDate();
        return fechaSeleccionada.isBefore(java.time.LocalDate.now());
    }

    private void configurarPermisosFechas(boolean soloLectura) {
        boolean administrador = esAdministrador();
        jDateChooser2.setEnabled(!soloLectura && administrador);
        jDateChooserEntrega.setEnabled(!soloLectura);
        jDateChooserPago.setEnabled(!soloLectura);

        java.util.Date fechaMinima = !soloLectura && !administrador
                ? java.sql.Date.valueOf(java.time.LocalDate.now())
                : null;
        jDateChooser2.setMinSelectableDate(fechaMinima);
        jDateChooserEntrega.setMinSelectableDate(fechaMinima);
        jDateChooserPago.setMinSelectableDate(fechaMinima);
    }

    private static final String RUTA_REGISTROC = AppConfig.registroCFile().getAbsolutePath();

    private static final String RUTA_INTORDENDECOMPRA = AppConfig.intOrdenCompraFile().getAbsolutePath();

    /**
     * Genera el siguiente número de orden único consultando REGISTROC.xlsx.
     * Busca el valor máximo en columna A de las filas cuya columna J = 1,
     * le suma 1 y verifica que no exista ya en columna A.
     * El resultado se muestra en jTextField6.
     */
    public void generarNumeroOrden() {
        try (FileInputStream fis = new FileInputStream(RUTA_REGISTROC);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            long maxNumero = 0;

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // saltar encabezado

                org.apache.poi.ss.usermodel.Cell celdaA = row.getCell(0); // columna A = NOORDEN
                if (celdaA == null) continue;

                long numOrden = 0;
                if (celdaA.getCellType() == CellType.NUMERIC) {
                    numOrden = (long) celdaA.getNumericCellValue();
                } else {
                    String val = dataFormatter.formatCellValue(celdaA).trim();
                    if (!val.isEmpty()) {
                        try { numOrden = Long.parseLong(val); } catch (NumberFormatException ignore) {}
                    }
                }

                if (numOrden > maxNumero) {
                    maxNumero = numOrden;
                }
            }

            jTextField6.setText(String.valueOf(maxNumero + 1));
            jTextField6.setEditable(estadoOrdenCargada == null);

        } catch (Exception e) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo generar número de orden", e);
            jTextField6.setText("1");
        }
        setModoSoloLectura(false);
    }

    // Agrega/actualiza encabezados en REGISTROC.xlsx cuando cambian las columnas del formulario.
    private void asegurarEncabezadosAbonadoEstatus(Sheet sheet) {
        Row encabezado = sheet.getRow(0);
        if (encabezado == null) {
            encabezado = sheet.createRow(0);
        }
        encabezado.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue("FECHA_ENTREGA");
        encabezado.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue("DIAS_PAGO");
        if (encabezado.getCell(22, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) == null) {
            encabezado.createCell(22).setCellValue("ABONADO");
        }
        if (encabezado.getCell(23, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) == null) {
            encabezado.createCell(23).setCellValue("ESTATUS");
        }
        if (encabezado.getCell(24, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) == null) {
            encabezado.createCell(24).setCellValue("OBSERVACIONES");
        }
    }//asegurarEncabezadosAbonadoEstatus

    // =====================================================================
    // Guardado y carga de ordenes en archivos Excel
    // =====================================================================

    private boolean guardarEnRegistroc() {
        // Verificar si el número de orden ya existe ANTES de pedir más datos
        java.io.File archivo = new java.io.File(RUTA_REGISTROC);
        if (archivo.exists()) {
            String noOrdenBuscar = jTextField6.getText().trim();
              try (Workbook wbPre = ORDEN_COMPRA_REPOSITORY.abrirWorkbook(archivo)) {
                Sheet sheetPre = wbPre.getSheetAt(0);
                for (Row r : sheetPre) {
                    if (r.getRowNum() == 0) continue;
                    org.apache.poi.ss.usermodel.Cell celdaA = r.getCell(0);
                    if (celdaA == null) continue;
                    String valA = celdaA.getCellType() == CellType.NUMERIC
                            ? String.valueOf((long) celdaA.getNumericCellValue())
                            : dataFormatter.formatCellValue(celdaA).trim();
                    if (valA.equals(noOrdenBuscar)) {
                        JOptionPane.showMessageDialog(this,
                                "La orden de compra No. " + noOrdenBuscar + " ya existe y no puede duplicarse.",
                                "Orden ya registrada", JOptionPane.WARNING_MESSAGE);
                        return false;
                    }
                }
            } catch (Exception exPre) {
                logger.log(java.util.logging.Level.WARNING, "No se pudo verificar duplicado", exPre);
            }
        }

        String errorValidacion = validarCampos();
        if (errorValidacion != null) {
            JOptionPane.showMessageDialog(this, errorValidacion,
                    "Campos incompletos", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        if (!generarPDFYFirmar("ELABORÓ")) {
            JOptionPane.showMessageDialog(this,
                    "La orden no se registró porque el PDF no fue firmado como ELABORÓ.",
                    "Firma requerida", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        Workbook wb;
        Sheet sheet;
        try {
            if (archivo.exists()) {
                wb = ORDEN_COMPRA_REPOSITORY.abrirWorkbook(archivo);
                sheet = wb.getSheetAt(0);
                asegurarEncabezadosAbonadoEstatus(sheet);
            } else {
                wb = ORDEN_COMPRA_REPOSITORY.abrirOCrearWorkbook(archivo, "REGISTROS");
                sheet = wb.getSheetAt(0);
                String[] headers = {"NOORDEN","FECHA","DOCUMENTO","COTIZACIÓN","ID_PROVEEDOR",
                    "CFDI","SOLICITUD","PROYECTO","PAGO","FORMAPAGO","METODOPAGO",
                    "FECHA_ENTREGA","DIAS_PAGO","DESCUENTO","IVA","ELABORO","AUTORIZO",
                    "MONEDA","TIPOCAMBIO","SUBTOTAL","IVA_IMPORTE","TOTAL","ABONADO","ESTATUS","OBSERVACIONES"};
                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < headers.length; i++) {
                    headerRow.createCell(i).setCellValue(headers[i]);
                }
            }

            int lastRow = sheet.getLastRowNum();
            Row row = sheet.createRow(lastRow + 1);

            // NOORDEN
            String noOrden = jTextField6.getText().trim();
            try { row.createCell(0).setCellValue(Long.parseLong(noOrden)); }
            catch (NumberFormatException ignore) { row.createCell(0).setCellValue(noOrden); }

            // FECHA
            java.util.Date fecha = jDateChooser2.getDate();
            if (fecha != null) {
                org.apache.poi.ss.usermodel.Cell celdaFecha = row.createCell(1);
                celdaFecha.setCellValue(fecha);
                org.apache.poi.ss.usermodel.CellStyle dateStyle = wb.createCellStyle();
                org.apache.poi.ss.usermodel.CreationHelper ch = wb.getCreationHelper();
                dateStyle.setDataFormat(ch.createDataFormat().getFormat("dd/MM/yyyy"));
                celdaFecha.setCellStyle(dateStyle);
            } else {
                row.createCell(1).setCellValue("");
            }

            // DOCUMENTO
            row.createCell(2).setCellValue(jLabel5.getText());
            // COTIZACIÓN
            row.createCell(3).setCellValue(jTextField4.getText().trim());
            // ID_PROVEEDOR
            row.createCell(4).setCellValue(jTextField1.getText().trim());
            // CFDI
            Object cfdi = jComboBox3.getSelectedItem();
            row.createCell(5).setCellValue(cfdi != null ? cfdi.toString() : "");
            // SOLICITUD
            Object solicitud = jComboBox1.getSelectedItem();
            row.createCell(6).setCellValue(solicitud != null ? solicitud.toString() : "");
            // PROYECTO
            row.createCell(7).setCellValue(jTextField2.getText().trim());
            // PAGO
            Object pago = jComboBox6.getSelectedItem();
            row.createCell(8).setCellValue(pago != null ? pago.toString() : "");
            // FORMAPAGO
            Object formaPago = jComboBox7.getSelectedItem();
            row.createCell(9).setCellValue(formaPago != null ? formaPago.toString() : "");
            // METODOPAGO
            Object metodoPago = jComboBox8.getSelectedItem();
            row.createCell(10).setCellValue(metodoPago != null ? metodoPago.toString() : "");
            escribirFechaCelda(wb, row, 11, jDateChooserEntrega.getDate());
            escribirFechaCelda(wb, row, 12, jDateChooserPago.getDate());
            // DESCUENTO
            String descuento = (jCheckBoxDescuento != null && jCheckBoxDescuento.isSelected())
                    ? jTextFieldDescuentoPct.getText().trim() : "0";
            row.createCell(13).setCellValue(descuento);
            // IVA
            row.createCell(14).setCellValue((jCheckBoxIVA != null && jCheckBoxIVA.isSelected()) ? "SI" : "NO");
            // ELABORO
            Object elaboro = jComboBox4.getSelectedItem();
            row.createCell(15).setCellValue(elaboro != null ? elaboro.toString() : "");
            // AUTORIZO
            Object autorizo = jComboBox5.getSelectedItem();
            row.createCell(16).setCellValue(autorizo != null ? autorizo.toString() : "");
            // MONEDA
            Object moneda = jComboBox2.getSelectedItem();
            row.createCell(17).setCellValue(moneda != null ? moneda.toString() : "");
            // TIPOCAMBIO
            row.createCell(18).setCellValue(jTextField5.isVisible() ? jTextField5.getText().trim() : "");
            // SUBTOTAL
            row.createCell(19).setCellValue(jLabel23.getText());
            // IVA
            row.createCell(20).setCellValue(jLabel24.getText());
            // TOTAL
            row.createCell(21).setCellValue(jLabel25.getText());
            // ABONADO: toda orden nueva comienza sin pagos registrados.
            row.createCell(22).setCellValue("0");
            row.createCell(23).setCellValue("PENDIENTE");
            // OBSERVACIONES
            row.createCell(24).setCellValue(jTextAreaObservaciones.getText().trim());

            ORDEN_COMPRA_REPOSITORY.guardarWorkbook(wb, archivo);
            wb.close();
            guardarTablaEnIntOrden(jTextField6.getText().trim());
            PRINCIPAL.actualizarAlarmasEnTiempoReal();
            JOptionPane.showMessageDialog(this,
                    "Orden de compra No. " + jTextField6.getText().trim() + " guardada correctamente.",
                    "Guardado exitoso", JOptionPane.INFORMATION_MESSAGE);
            salirDeOrden();
            return true;

        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo guardar en REGISTROC.xlsx", ex);
            JOptionPane.showMessageDialog(this,
                    "No se pudo guardar en REGISTROC.xlsx:\n" + ex.getMessage(),
                    "Advertencia", JOptionPane.WARNING_MESSAGE);
            return false;
        }
    }

    /** Actualiza la orden buscada en REGISTROC.xlsx (solo accesible para admin) */
    private boolean guardarCambiosOrden() {
        String noOrden = jTextField6.getText().trim();
        if (noOrden.isEmpty()) {
            JOptionPane.showMessageDialog(this, "No hay ninguna orden cargada para guardar.", "Guardar Cambios", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        String errorValidacion = validarCampos();
        if (errorValidacion != null) {
            JOptionPane.showMessageDialog(this, errorValidacion,
                    "Campos incompletos", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        java.io.File archivo = new java.io.File(RUTA_REGISTROC);
        if (!archivo.exists()) {
            JOptionPane.showMessageDialog(this, "No se encontró el archivo de registros.", "Error", JOptionPane.ERROR_MESSAGE);
            return false;
        }

        try (Workbook wb = ORDEN_COMPRA_REPOSITORY.abrirWorkbook(archivo)) {

            Sheet sheet = wb.getSheetAt(0);
            asegurarEncabezadosAbonadoEstatus(sheet);
            Row rowToUpdate = null;

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                org.apache.poi.ss.usermodel.Cell celdaA = row.getCell(0);
                if (celdaA == null) continue;
                String valA = celdaA.getCellType() == CellType.NUMERIC
                        ? String.valueOf((long) celdaA.getNumericCellValue())
                        : dataFormatter.formatCellValue(celdaA).trim();
                if (valA.equals(noOrden)) {
                    rowToUpdate = row;
                    break;
                }
            }

            if (rowToUpdate == null) {
                JOptionPane.showMessageDialog(this, "No se encontró la orden № " + noOrden + " para actualizar.", "Guardar Cambios", JOptionPane.WARNING_MESSAGE);
                return false;
            }

            // FECHA
            java.util.Date fecha = jDateChooser2.getDate();
            if (fecha != null) {
                org.apache.poi.ss.usermodel.Cell celdaFecha = rowToUpdate.getCell(1);
                if (celdaFecha == null) celdaFecha = rowToUpdate.createCell(1);
                celdaFecha.setCellValue(fecha);
                org.apache.poi.ss.usermodel.CellStyle dateStyle = wb.createCellStyle();
                dateStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));
                celdaFecha.setCellStyle(dateStyle);
            }
            // DOCUMENTO
            setCellString(rowToUpdate, 2, jLabel5.getText());
            // COTIZACIÓN
            setCellString(rowToUpdate, 3, jTextField4.getText().trim());
            // ID_PROVEEDOR
            setCellString(rowToUpdate, 4, jTextField1.getText().trim());
            // CFDI
            setCellString(rowToUpdate, 5, jComboBox3.getSelectedItem() != null ? jComboBox3.getSelectedItem().toString() : "");
            // SOLICITUD
            setCellString(rowToUpdate, 6, jComboBox1.getSelectedItem() != null ? jComboBox1.getSelectedItem().toString() : "");
            // PROYECTO
            setCellString(rowToUpdate, 7, jTextField2.getText().trim());
            // PAGO
            setCellString(rowToUpdate, 8, jComboBox6.getSelectedItem() != null ? jComboBox6.getSelectedItem().toString() : "");
            // FORMAPAGO
            setCellString(rowToUpdate, 9, jComboBox7.getSelectedItem() != null ? jComboBox7.getSelectedItem().toString() : "");
            // METODOPAGO
            setCellString(rowToUpdate, 10, jComboBox8.getSelectedItem() != null ? jComboBox8.getSelectedItem().toString() : "");
            escribirFechaCelda(wb, rowToUpdate, 11, jDateChooserEntrega.getDate());
            escribirFechaCelda(wb, rowToUpdate, 12, jDateChooserPago.getDate());
            // DESCUENTO
            String descuento = (jCheckBoxDescuento != null && jCheckBoxDescuento.isSelected())
                    ? jTextFieldDescuentoPct.getText().trim() : "0";
            setCellString(rowToUpdate, 13, descuento);
            // IVA
            setCellString(rowToUpdate, 14, (jCheckBoxIVA != null && jCheckBoxIVA.isSelected()) ? "SI" : "NO");
            // ELABORO
            setCellString(rowToUpdate, 15, jComboBox4.getSelectedItem() != null ? jComboBox4.getSelectedItem().toString() : "");
            // AUTORIZO
            setCellString(rowToUpdate, 16, jComboBox5.getSelectedItem() != null ? jComboBox5.getSelectedItem().toString() : "");
            // MONEDA
            setCellString(rowToUpdate, 17, jComboBox2.getSelectedItem() != null ? jComboBox2.getSelectedItem().toString() : "");
            // TIPOCAMBIO
            setCellString(rowToUpdate, 18, jTextField5.isVisible() ? jTextField5.getText().trim() : "");
            // SUBTOTAL, IVA_IMPORTE, TOTAL
            setCellString(rowToUpdate, 19, jLabel23.getText());
            setCellString(rowToUpdate, 20, jLabel24.getText());
            setCellString(rowToUpdate, 21, jLabel25.getText());
            // OBSERVACIONES
            setCellString(rowToUpdate, 24, jTextAreaObservaciones.getText().trim());

            ORDEN_COMPRA_REPOSITORY.guardarWorkbook(wb, archivo);

            guardarTablaEnIntOrden(noOrden);
            PRINCIPAL.actualizarAlarmasEnTiempoReal();

            JOptionPane.showMessageDialog(this,
                    "Orden № " + noOrden + " actualizada correctamente.",
                    "Guardar Cambios", JOptionPane.INFORMATION_MESSAGE);
            setModoSoloLectura(true);
                actualizarBotonesEdicion(true, false);
            estadoOrdenCargada = obtenerEstadoOrden();
            return true;

        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "Error al guardar cambios de orden", ex);
            JOptionPane.showMessageDialog(this, "Error al guardar cambios:\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            return false;
        }
    }

    /** Utilidad: asigna un valor string a la celda, creándola si no existe */
    private void setCellString(Row row, int col, String value) {
        org.apache.poi.ss.usermodel.Cell cell = row.getCell(col);
        if (cell == null) cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
    }

    private void escribirFechaCelda(Workbook workbook, Row row, int columna, java.util.Date fecha) {
        org.apache.poi.ss.usermodel.Cell celda = row.getCell(columna);
        if (celda == null) {
            celda = row.createCell(columna);
        }
        if (fecha == null) {
            celda.setBlank();
            return;
        }
        celda.setCellValue(fecha);
        org.apache.poi.ss.usermodel.CellStyle estiloFecha = workbook.createCellStyle();
        estiloFecha.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));
        celda.setCellStyle(estiloFecha);
    }

    private java.util.Date leerFechaCelda(org.apache.poi.ss.usermodel.Cell celda) {
        if (celda == null) {
            return null;
        }
        if (celda.getCellType() == CellType.NUMERIC && org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(celda)) {
            return celda.getDateCellValue();
        }
        String texto = dataFormatter.formatCellValue(celda).trim();
        if (texto.isEmpty()) {
            return null;
        }
        for (String formato : new String[]{"dd/MM/yyyy", "yyyy-MM-dd"}) {
            try {
                SimpleDateFormat parser = new SimpleDateFormat(formato);
                parser.setLenient(false);
                return parser.parse(texto);
            } catch (java.text.ParseException ignored) {
            }
        }
        return null;
    }

    private float calcularAltoPrimeraPagina() {
        DefaultTableModel modelo = (DefaultTableModel) jTable1.getModel();
        int filas = 0;
        int lineasDescripcion = 0;
        for (int i = 0; i < modelo.getRowCount(); i++) {
            Object codigo = modelo.getValueAt(i, 0);
            if (codigo == null || codigo.toString().trim().isEmpty()) {
                continue;
            }
            filas++;
            String descripcion = String.valueOf(modelo.getValueAt(i, 1));
            // Las descripciones se ajustan a la columna del PDF.
            lineasDescripcion += Math.max(1, (descripcion.length() + 39) / 40);
        }

        int lineasObservaciones = jTextAreaObservaciones == null
                ? 0
                : Math.max(0, (jTextAreaObservaciones.getText().trim().length() + 95) / 96);

        // Reserva el contenido fijo, las filas de productos y el bloque de firmas.
        float altoEstimado = 750f
                + (filas * 11f)
                + (lineasDescripcion * 8f)
            + (lineasObservaciones * 9f)
            + 45f;
        return Math.max(PageSize.LETTER.getHeight(), altoEstimado);
    }

    private String formatearFecha(java.util.Date fecha) {
        return fecha == null ? "" : new SimpleDateFormat("d MMM yyyy", Locale.forLanguageTag("es-MX")).format(fecha);
    }

    // =====================================================================
    // Busqueda, edicion y navegacion de ordenes
    // =====================================================================

    private void buscarOSalirOrden() {
        if (estadoOrdenCargada != null) {
            salirDeOrden();
            return;
        }
        buscarOrden();
    }

    private void navegarOrden(int direccion) {
        java.util.List<Integer> ordenes = obtenerNumerosOrdenes();
        if (ordenes.isEmpty()) {
            JOptionPane.showMessageDialog(this, "No hay órdenes registradas para navegar.",
                    "Navegación", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        int actual;
        try {
            actual = Integer.parseInt(jTextField6.getText().trim());
        } catch (NumberFormatException ex) {
            actual = direccion > 0 ? ordenes.get(0) : ordenes.get(ordenes.size() - 1);
        }
        int indice = ordenes.indexOf(actual);
        if (indice < 0) {
            indice = direccion > 0 ? 0 : ordenes.size() - 1;
        } else {
            indice += direccion;
            if (indice < 0 || indice >= ordenes.size()) {
                return;
            }
        }
        cargarOrdenPorNumero(ordenes.get(indice));
    }

    private void irPrimeraOrden() {
        java.util.List<Integer> ordenes = obtenerNumerosOrdenes();
        if (!ordenes.isEmpty()) cargarOrdenPorNumero(ordenes.get(0));
    }

    private void irUltimaOrden() {
        java.util.List<Integer> ordenes = obtenerNumerosOrdenes();
        if (!ordenes.isEmpty()) cargarOrdenPorNumero(ordenes.get(ordenes.size() - 1));
    }

    private java.util.List<Integer> obtenerNumerosOrdenes() {
        java.util.List<Integer> ordenes = new java.util.ArrayList<>();
        java.io.File archivo = new java.io.File(RUTA_REGISTROC);
        if (!archivo.isFile()) return ordenes;
        try (Workbook libro = ORDEN_COMPRA_REPOSITORY.abrirWorkbook(archivo)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row fila : hoja) {
                if (fila.getRowNum() == 0 || fila.getCell(0) == null) continue;
                try {
                    ordenes.add(Integer.parseInt(dataFormatter.formatCellValue(fila.getCell(0)).trim()));
                } catch (NumberFormatException ex) {
                    logger.log(java.util.logging.Level.FINE, "Se ignoró un número de orden inválido.", ex);
                }
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudieron leer las órdenes para navegar.", ex);
        }
        java.util.Collections.sort(ordenes);
        return ordenes;
    }

    private void cargarOrdenPorNumero(int numero) {
        jTextField6.setText(String.valueOf(numero));
        buscarOrden();
    }

    private void actualizarBotonBuscar(boolean ordenCargada) {
        jButton2.setText(ordenCargada ? "SALIR" : "BUSCAR");
        jButton2.setBackground(ordenCargada
                ? new java.awt.Color(139, 0, 0)
                : new java.awt.Color(0, 204, 51));
    }

    private void prepararDuplicado() {
        // Se conservan los datos cargados, pero el duplicado inicia con
        // el siguiente consecutivo y la fecha actual.
        estadoOrdenCargada = null;
        ultimoPdfGenerado = null;
        generarNumeroOrden();
        jDateChooser2.setDate(new java.util.Date());
        jDateChooserEntrega.setDate(null);
        jDateChooserPago.setDate(null);
        setModoSoloLectura(false);
        jButton3.setText("CREAR");
        actualizarBotonBuscar(false);
        actualizarBotonesEdicion(false, false);
        actualizarEstadoBotonPDF();
    }

    private void actualizarEstadoBotonPDF() {
        if (jButton1 != null) {
            jButton1.setEnabled(estadoOrdenCargada != null);
        }
        actualizarEstadoBotonFirma();
    }

    private void buscarOrden() {
        String buscar = jTextField6.getText().trim();
        if (buscar.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Ingresa un número de orden para buscar.", "Buscar", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try (FileInputStream fis = new FileInputStream(RUTA_REGISTROC);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                org.apache.poi.ss.usermodel.Cell celdaA = row.getCell(0);
                if (celdaA == null) continue;

                String valA = celdaA.getCellType() == CellType.NUMERIC
                        ? String.valueOf((long) celdaA.getNumericCellValue())
                        : dataFormatter.formatCellValue(celdaA).trim();

                if (!valA.equals(buscar)) continue;

                // Col 1 - FECHA
                org.apache.poi.ss.usermodel.Cell cFecha = row.getCell(1);
                if (cFecha != null && cFecha.getCellType() == CellType.NUMERIC
                        && org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cFecha)) {
                    jDateChooser2.setDate(cFecha.getDateCellValue());
                }

                // Col 2 - DOCUMENTO (jLabel5)
                org.apache.poi.ss.usermodel.Cell cDoc = row.getCell(2);
                if (cDoc != null) jLabel5.setText(dataFormatter.formatCellValue(cDoc).trim());

                // Col 3 - COTIZACIÓN
                org.apache.poi.ss.usermodel.Cell cCot = row.getCell(3);
                if (cCot != null) jTextField4.setText(dataFormatter.formatCellValue(cCot).trim());

                // Col 4 - ID_PROVEEDOR
                org.apache.poi.ss.usermodel.Cell cProv = row.getCell(4);
                if (cProv != null) {
                    String idProv = dataFormatter.formatCellValue(cProv).trim();
                    seleccionandoProveedor = true;
                    jTextField1.setText(idProv);
                    seleccionandoProveedor = false;
                    buscarProveedor(idProv);
                }

                // Col 5 - CFDI
                seleccionarCombo(jComboBox3, row.getCell(5));

                // Col 6 - SOLICITUD (SOLICITANTE -> jComboBox1)
                seleccionarCombo(jComboBox1, row.getCell(6));

                // Col 7 - PROYECTO
                org.apache.poi.ss.usermodel.Cell cProy = row.getCell(7);
                if (cProy != null) jTextField2.setText(dataFormatter.formatCellValue(cProy).trim());

                // Col 8 - PAGO
                seleccionarCombo(jComboBox6, row.getCell(8));

                // Col 9 - FORMAPAGO
                seleccionarCombo(jComboBox7, row.getCell(9));

                // Col 10 - METODOPAGO
                seleccionarCombo(jComboBox8, row.getCell(10));

                jDateChooserEntrega.setDate(leerFechaCelda(row.getCell(11)));
                jDateChooserPago.setDate(leerFechaCelda(row.getCell(12)));

                // Col 13 - DESCUENTO
                org.apache.poi.ss.usermodel.Cell cDesc = row.getCell(13);
                if (cDesc != null) {
                    String dVal = dataFormatter.formatCellValue(cDesc).trim();
                    boolean tieneDesc = !dVal.isEmpty() && !dVal.equals("0");
                    jCheckBoxDescuento.setSelected(tieneDesc);
                    jTextFieldDescuentoPct.setText(tieneDesc ? dVal : "0");
                    jTextFieldDescuentoPct.setVisible(tieneDesc);
                }

                // Col 14 - IVA
                org.apache.poi.ss.usermodel.Cell cIva = row.getCell(14);
                if (cIva != null) {
                    jCheckBoxIVA.setSelected("SI".equalsIgnoreCase(dataFormatter.formatCellValue(cIva).trim()));
                }

                // Col 15 - ELABORO
                seleccionarCombo(jComboBox4, row.getCell(15));

                // Col 16 - AUTORIZO
                seleccionarCombo(jComboBox5, row.getCell(16));

                // Col 17 - MONEDA
                seleccionarCombo(jComboBox2, row.getCell(17));

                // Col 18 - TIPOCAMBIO
                org.apache.poi.ss.usermodel.Cell cTc = row.getCell(18);
                if (cTc != null) jTextField5.setText(dataFormatter.formatCellValue(cTc).trim());

                actualizarTotales();

                // Col 19 - SUBTOTAL
                org.apache.poi.ss.usermodel.Cell cSub = row.getCell(19);
                if (cSub != null) { String v = dataFormatter.formatCellValue(cSub).trim(); if (!v.isEmpty()) jLabel23.setText(v); }
                // Col 20 - IVA
                org.apache.poi.ss.usermodel.Cell cIvaVal = row.getCell(20);
                if (cIvaVal != null) { String v = dataFormatter.formatCellValue(cIvaVal).trim(); if (!v.isEmpty()) jLabel24.setText(v); }
                // Col 21 - TOTAL
                org.apache.poi.ss.usermodel.Cell cTot = row.getCell(21);
                if (cTot != null) { String v = dataFormatter.formatCellValue(cTot).trim(); if (!v.isEmpty()) jLabel25.setText(v); }

                // Col 24 - OBSERVACIONES
                org.apache.poi.ss.usermodel.Cell cObs = row.getCell(24);
                jTextAreaObservaciones.setText(cObs != null ? dataFormatter.formatCellValue(cObs).trim() : "");

                cargarTablaDesdeIntOrden(buscar);
                setModoSoloLectura(true);
                estadoOrdenCargada = obtenerEstadoOrden();
                actualizarBotonBuscar(true);
                actualizarEstadoBotonPDF();
                jButton3.setText("DUPLICAR");
                actualizarBotonesEdicion(true, false);
                return;
            }

            JOptionPane.showMessageDialog(this, "No se encontró la orden № " + buscar + ".", "Buscar", JOptionPane.INFORMATION_MESSAGE);

        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "Error al buscar orden", ex);
            JOptionPane.showMessageDialog(this, "Error al buscar:\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    /** Guarda las filas de jTable1 en INTORDENDECOMPRA.xlsx asociadas al número de orden dado */
    private void guardarTablaEnIntOrden(String noOrden) {
        java.io.File archivo = new java.io.File(RUTA_INTORDENDECOMPRA);
        Workbook wb;
        Sheet sheet;
        try {
            if (archivo.exists()) {
                wb = ORDEN_COMPRA_REPOSITORY.abrirWorkbook(archivo);
                sheet = wb.getSheetAt(0);
                // Eliminar filas existentes de esta orden para evitar duplicados
                for (int i = sheet.getLastRowNum(); i >= 1; i--) {
                    Row r = sheet.getRow(i);
                    if (r == null) continue;
                    org.apache.poi.ss.usermodel.Cell c0 = r.getCell(0);
                    if (c0 == null) continue;
                    String val = c0.getCellType() == CellType.NUMERIC
                            ? String.valueOf((long) c0.getNumericCellValue())
                            : dataFormatter.formatCellValue(c0).trim();
                    if (val.equals(noOrden)) {
                        sheet.removeRow(r);
                    }
                }
            } else {
                wb = ORDEN_COMPRA_REPOSITORY.abrirOCrearWorkbook(archivo, "INTORDEN");
                sheet = wb.getSheetAt(0);
                String[] headers = {"NOORDEN", "CODIGO", "DESCRIPCION", "UNIDAD", "CANTIDAD", "PRECIO", "IMPORTE"};
                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < headers.length; i++) {
                    headerRow.createCell(i).setCellValue(headers[i]);
                }
            }

            DefaultTableModel modelo = (DefaultTableModel) jTable1.getModel();
            for (int i = 0; i < modelo.getRowCount(); i++) {
                Object codigo = modelo.getValueAt(i, 0);
                if (codigo == null || codigo.toString().trim().isEmpty()) continue;
                int lastRow = sheet.getLastRowNum();
                Row fila = sheet.createRow(lastRow + 1);
                // Col 0: NOORDEN
                try { fila.createCell(0).setCellValue(Long.parseLong(noOrden)); }
                catch (NumberFormatException ignore) { fila.createCell(0).setCellValue(noOrden); }
                // Col 1: CODIGO
                fila.createCell(1).setCellValue(codigo.toString().trim());
                // Col 2: DESCRIPCION
                Object desc = modelo.getValueAt(i, 1);
                fila.createCell(2).setCellValue(desc != null ? desc.toString() : "");
                // Col 3: UNIDAD
                Object unidad = modelo.getValueAt(i, 2);
                fila.createCell(3).setCellValue(unidad != null ? unidad.toString() : "");
                // Col 4: CANTIDAD
                Object cant = modelo.getValueAt(i, 3);
                fila.createCell(4).setCellValue(cant instanceof Number n ? n.doubleValue() : parseNumero(cant));
                // Col 5: PRECIO
                Object precio = modelo.getValueAt(i, 4);
                fila.createCell(5).setCellValue(precio instanceof Number n ? n.doubleValue() : parseNumero(precio));
                // Col 6: IMPORTE
                Object importe = modelo.getValueAt(i, 5);
                fila.createCell(6).setCellValue(importe instanceof Number n ? n.doubleValue() : parseNumero(importe));
            }

            ORDEN_COMPRA_REPOSITORY.guardarWorkbook(wb, archivo);
            wb.close();
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo guardar en INTORDENDECOMPRA.xlsx", ex);
        }
    }

    /** Carga las filas de INTORDENDECOMPRA.xlsx que corresponden al número de orden dado en jTable1 */
    private void cargarTablaDesdeIntOrden(String noOrden) {
        java.io.File archivo = new java.io.File(RUTA_INTORDENDECOMPRA);
        if (!archivo.exists()) return;
        try (FileInputStream fis = new FileInputStream(archivo);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            DefaultTableModel modelo = (DefaultTableModel) jTable1.getModel();
            modelo.setRowCount(0); // limpiar tabla

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                org.apache.poi.ss.usermodel.Cell c0 = row.getCell(0);
                if (c0 == null) continue;
                String val = c0.getCellType() == CellType.NUMERIC
                        ? String.valueOf((long) c0.getNumericCellValue())
                        : dataFormatter.formatCellValue(c0).trim();
                if (!val.equals(noOrden)) continue;

                String codigo  = dataFormatter.formatCellValue(row.getCell(1));
                String desc    = dataFormatter.formatCellValue(row.getCell(2));
                String unidad  = dataFormatter.formatCellValue(row.getCell(3));
                org.apache.poi.ss.usermodel.Cell cCant   = row.getCell(4);
                org.apache.poi.ss.usermodel.Cell cPrecio = row.getCell(5);
                org.apache.poi.ss.usermodel.Cell cImp    = row.getCell(6);
                int cantidad = cCant   != null && cCant.getCellType()   == CellType.NUMERIC ? (int) cCant.getNumericCellValue()   : 0;
                double precio  = cPrecio != null && cPrecio.getCellType() == CellType.NUMERIC ? cPrecio.getNumericCellValue() : 0;
                double importe = cImp    != null && cImp.getCellType()    == CellType.NUMERIC ? cImp.getNumericCellValue()    : 0;
                modelo.addRow(new Object[]{codigo, desc, unidad, cantidad, precio, importe});
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo cargar tabla desde INTORDENDECOMPRA.xlsx", ex);
        }
    }

    /** Selecciona en un JComboBox el valor que coincide con el texto de la celda */
    private void seleccionarCombo(javax.swing.JComboBox<String> combo, org.apache.poi.ss.usermodel.Cell celda) {
        if (celda == null) return;
        String val = dataFormatter.formatCellValue(celda).trim();
        javax.swing.ComboBoxModel<String> model = combo.getModel();
        for (int i = 0; i < model.getSize(); i++) {
            if (model.getElementAt(i).equalsIgnoreCase(val)) {
                combo.setSelectedIndex(i);
                return;
            }
        }
    }

    private String obtenerEstadoOrden() {
        StringBuilder estado = new StringBuilder();
        estado.append(jTextField6.getText()).append('|')
                .append(jTextField4.getText()).append('|')
                .append(jTextField1.getText()).append('|')
                .append(jTextField2.getText()).append('|')
                .append(jTextField3.getText()).append('|')
                .append(jTextField5.getText()).append('|')
                .append(jDateChooser2.getDate() == null ? "" : jDateChooser2.getDate().getTime()).append('|')
                .append(jDateChooserEntrega.getDate() == null ? "" : jDateChooserEntrega.getDate().getTime()).append('|')
                .append(jDateChooserPago.getDate() == null ? "" : jDateChooserPago.getDate().getTime()).append('|');
        javax.swing.JComboBox<?>[] combos = {
            jComboBox1, jComboBox2, jComboBox3, jComboBox4, jComboBox5,
            jComboBox6, jComboBox7, jComboBox8
        };
        for (javax.swing.JComboBox<?> combo : combos) {
            estado.append(combo.getSelectedItem()).append('|');
        }
        estado.append(jCheckBoxIVA != null && jCheckBoxIVA.isSelected()).append('|')
                .append(jCheckBoxDescuento != null && jCheckBoxDescuento.isSelected()).append('|')
                .append(jTextFieldDescuentoPct == null ? "" : jTextFieldDescuentoPct.getText()).append('|')
                .append(jCheckBoxISR != null && jCheckBoxISR.isSelected()).append('|')
                .append(jTextFieldISRPercent == null ? "" : jTextFieldISRPercent.getText()).append('|');
        DefaultTableModel modelo = (DefaultTableModel) jTable1.getModel();
        for (int fila = 0; fila < modelo.getRowCount(); fila++) {
            for (int columna = 0; columna < modelo.getColumnCount(); columna++) {
                estado.append(modelo.getValueAt(fila, columna)).append('|');
            }
        }
        return estado.toString();
    }

    /** Limpia el formulario y habilita la creación de una nueva orden */
    private void salirDeOrden() {
        if (estadoOrdenCargada != null && !estadoOrdenCargada.equals(obtenerEstadoOrden())) {
            int respuesta = JOptionPane.showConfirmDialog(this,
                    "Hay cambios sin guardar. ¿Deseas guardarlos antes de salir?",
                    "Cambios sin guardar",
                    JOptionPane.YES_NO_CANCEL_OPTION,
                    JOptionPane.WARNING_MESSAGE);
            if (respuesta == JOptionPane.CANCEL_OPTION || respuesta == JOptionPane.CLOSED_OPTION) {
                return;
            }
            if (respuesta == JOptionPane.YES_OPTION && !guardarCambiosOrden()) {
                return;
            }
        }

        // ── Campos de texto ───────────────────────────────────────────────
        seleccionandoProveedor = true;
        jTextField1.setText("");
        seleccionandoProveedor = false;
        jTextField2.setText("");
        jTextField3.setText("");
        jTextField5.setText("");
        jTextAreaObservaciones.setText("");

        // Restaurar cotización con fecha de hoy
        java.time.LocalDate hoy = java.time.LocalDate.now();
        jTextField4.setText(String.format("AVANTE%04d%02d%02d",
                hoy.getYear(), hoy.getMonthValue(), hoy.getDayOfMonth()));

        // Restaurar fecha a hoy
        jDateChooser2.setDate(new java.util.Date());
        jDateChooserEntrega.setDate(null);
        jDateChooserPago.setDate(null);

        // ── Etiquetas de proveedor / dirección / área ─────────────────────
        jLabel8.setText("");
        jLabel15.setText("");
        jLabel26.setText("");
        jLabel30.setText("");
        jLabel32.setText("");
        proveedorTelefono = "";
        proveedorCorreo   = "";

        // ── Combos: seleccionar primer elemento disponible ────────────────
        if (jComboBox1.getItemCount() > 0) jComboBox1.setSelectedIndex(0);
        if (jComboBox3.getItemCount() > 0) jComboBox3.setSelectedIndex(0);
        if (jComboBox4.getItemCount() > 0) jComboBox4.setSelectedIndex(0);
        if (jComboBox5.getItemCount() > 0) jComboBox5.setSelectedIndex(0);
        if (jComboBox6.getItemCount() > 0) jComboBox6.setSelectedIndex(0);
        if (jComboBox7.getItemCount() > 0) jComboBox7.setSelectedIndex(0);
        if (jComboBox8.getItemCount() > 0) jComboBox8.setSelectedIndex(0);

        // Moneda: restaurar al primer elemento y actualizar descripción / visibilidad tipo cambio
        if (jComboBox2.getItemCount() > 0) {
            jComboBox2.setSelectedIndex(0);
            String clave = jComboBox2.getSelectedItem().toString();
            jLabel37.setText(monedaDescMap.getOrDefault(clave, ""));
            boolean mostrarCambio = clave.equalsIgnoreCase("USD") || clave.equalsIgnoreCase("EUR") || clave.equalsIgnoreCase("CAD");
            jTextField5.setVisible(mostrarCambio);
            jLabel38.setVisible(mostrarCambio);
        }

        // ── Checkboxes ─────────────────────────────────────────────────────
        if (jCheckBoxIVA != null) jCheckBoxIVA.setSelected(true);
        if (jCheckBoxDescuento != null) {
            jCheckBoxDescuento.setSelected(false);
            if (jTextFieldDescuentoPct != null) {
                jTextFieldDescuentoPct.setText("0");
                jTextFieldDescuentoPct.setVisible(false);
            }
        }
        if (jCheckBoxISR != null) {
            jCheckBoxISR.setSelected(false);
            if (jTextFieldISRPercent != null) {
                jTextFieldISRPercent.setText("0");
                jTextFieldISRPercent.setVisible(false);
            }
        }

        // ── Tabla de productos y totales ──────────────────────────────────
        ((DefaultTableModel) jTable1.getModel()).setRowCount(0);
        jLabel20.setText("SUB TOTAL:");
        jLabel23.setText("$0.00");
        jLabel24.setText("$0.00");
        jLabel25.setText("$0.00");

        // ── Restablecer número de documento ───────────────────────────────
        jLabel5.setText("FO-SG-04-00");

        // ── Generar nuevo número de orden y re-habilitar edición ──────────
        generarNumeroOrden();
        estadoOrdenCargada = null;
        actualizarBotonBuscar(false);
        jButton3.setText("CREAR");
        actualizarEstadoBotonPDF();
        actualizarBotonesEdicion(false, false);
    }

    private void activarEdicionOrden() {
        setModoSoloLectura(false);
        actualizarBotonesEdicion(true, true);
    }

    /** Maneja el clic del botón único: en modo lectura activa la edición, en modo edición guarda los cambios. */
    private void onBotonEditarGuardarClick() {
        if (jButtonEditar == null) {
            return;
        }
        if (modoCrear) {
            guardarCambiosOrden();
        } else {
            activarEdicionOrden();
        }
    }

    private void actualizarBotonesEdicion(boolean ordenCargada, boolean editando) {
        boolean habilitar = esAdministrador() && ordenCargada;
        if (jButtonEditar != null) {
            jButtonEditar.setVisible(true);
            jButtonEditar.setEnabled(habilitar);
            if (habilitar && editando) {
                jButtonEditar.setText("GUARDAR CAMBIOS");
                jButtonEditar.setBackground(new java.awt.Color(0, 128, 0));
            } else {
                jButtonEditar.setText("EDITAR");
                jButtonEditar.setBackground(new java.awt.Color(0, 90, 180));
            }
        }
    }

    /** Habilita o deshabilita la edición de todos los campos del formulario */
    private void setModoSoloLectura(boolean soloLectura) {
        modoCrear = !soloLectura;
        // El número identifica permanentemente la orden. Solo se permite
        // escribirlo cuando se está buscando una orden o creando una nueva.
        jTextField6.setEditable(!soloLectura && estadoOrdenCargada == null);
        jTextField1.setEditable(false); // se activa solo al hacer clic (modo crear)
        jTextField2.setEditable(!soloLectura);
        jTextField3.setEditable(!soloLectura);
        jTextField4.setEditable(!soloLectura);
        jTextField5.setEditable(!soloLectura);
        configurarPermisosFechas(soloLectura);
        jComboBox1.setEnabled(!soloLectura);
        jComboBox2.setEnabled(!soloLectura);
        jComboBox3.setEnabled(!soloLectura);
        jComboBox4.setEnabled(!soloLectura);
        jComboBox5.setEnabled(!soloLectura);
        jComboBox6.setEnabled(!soloLectura);
        jComboBox7.setEnabled(!soloLectura);
        jComboBox8.setEnabled(!soloLectura);
        jLabel5.setCursor(soloLectura
                ? java.awt.Cursor.getDefaultCursor()
                : java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR));
        if (jCheckBoxIVA != null) jCheckBoxIVA.setEnabled(!soloLectura);
        if (jCheckBoxDescuento != null) jCheckBoxDescuento.setEnabled(!soloLectura);
        if (jTextFieldDescuentoPct != null) jTextFieldDescuentoPct.setEditable(!soloLectura);
        if (jCheckBoxISR != null) jCheckBoxISR.setEnabled(!soloLectura);
        if (jTextFieldISRPercent != null) jTextFieldISRPercent.setEditable(!soloLectura);
        jTable1.setEnabled(!soloLectura);
        jButton7.setEnabled(!soloLectura);
        jButton3.setEnabled(true);
        jTextAreaObservaciones.setEditable(!soloLectura);
    }

    public static void main(String args[]) {

        java.awt.EventQueue.invokeLater(() -> new ORDENCOMPRA().setVisible(true));
    }

    private javax.swing.JCheckBox jCheckBoxIVA;
    private javax.swing.JCheckBox jCheckBoxISR;
    private String proveedorTelefono = "";
    private String proveedorCorreo = "";
    private javax.swing.JCheckBox jCheckBoxDescuento;
    private javax.swing.JTextField jTextFieldDescuentoPct;
    private javax.swing.JTextField jTextFieldISRPercent;
    private javax.swing.JButton jButtonEditar;
    private javax.swing.JLabel jLabelObservaciones;
    private javax.swing.JTextArea jTextAreaObservaciones;
    private javax.swing.JScrollPane jScrollPaneObservaciones;
    private double subtotalOriginal = 0;
    private double descuentoImporte = 0;
    private double isrImporte = 0;
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton6;
    private javax.swing.JButton jButton7;
    private javax.swing.JComboBox<String> jComboBox1;
    private javax.swing.JComboBox<String> jComboBox2;
    private javax.swing.JComboBox<String> jComboBox3;
    private javax.swing.JComboBox<String> jComboBox4;
    private javax.swing.JComboBox<String> jComboBox5;
    private javax.swing.JComboBox<String> jComboBox6;
    private javax.swing.JComboBox<String> jComboBox7;
    private javax.swing.JComboBox<String> jComboBox8;
    private com.toedter.calendar.JDateChooser jDateChooser2;
    private com.toedter.calendar.JDateChooser jDateChooserEntrega;
    private com.toedter.calendar.JDateChooser jDateChooserPago;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel15;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel17;
    private javax.swing.JLabel jLabel18;
    private javax.swing.JLabel jLabel19;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel20;
    private javax.swing.JLabel jLabel21;
    private javax.swing.JLabel jLabel22;
    private javax.swing.JLabel jLabel23;
    private javax.swing.JLabel jLabel24;
    private javax.swing.JLabel jLabel25;
    private javax.swing.JLabel jLabel26;
    private javax.swing.JLabel jLabel27;
    private javax.swing.JLabel jLabel29;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel30;
    private javax.swing.JLabel jLabel31;
    private javax.swing.JLabel jLabel32;
    private javax.swing.JLabel jLabel34;
    private javax.swing.JLabel jLabel36;
    private javax.swing.JLabel jLabel37;
    private javax.swing.JLabel jLabel38;
    private javax.swing.JLabel jLabel39;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel40;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField2;
    private javax.swing.JTextField jTextField3;
    private javax.swing.JTextField jTextField4;
    private javax.swing.JTextField jTextField5;
    private javax.swing.JTextField jTextField6;
    // End of variables declaration//GEN-END:variables
}