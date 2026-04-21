/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.ocxrst;

import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.kernel.pdf.xobject.PdfFormXObject;
import com.itextpdf.kernel.events.PdfDocumentEvent;
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
import java.text.SimpleDateFormat;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author Maria Luisa Martinez
 */
public class ORDENCOMPRA extends javax.swing.JFrame {

    private static final String RUTA_INVENTARIOS = "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\INVENTARIOS.xlsx";
    private final DataFormatter dataFormatter = new DataFormatter();
    private final java.util.Map<String, String> monedaDescMap = new java.util.HashMap<>();

    private static final java.util.logging.Logger logger = java.util.logging.Logger.getLogger(ORDENCOMPRA.class.getName());
    private final java.util.Map<String, String> monedaLeyendaMap = new java.util.HashMap<>();

    // Caché de inventario: código -> "código | descripción"
    private final java.util.List<String[]> cacheInventario = new java.util.ArrayList<>();
    private final javax.swing.JPopupMenu popupProductos = new javax.swing.JPopupMenu();

    // Caché de proveedores: id -> "id | nombre"
    private final java.util.List<String[]> cacheProveedores = new java.util.ArrayList<>();
    private final javax.swing.JPopupMenu popupProveedores = new javax.swing.JPopupMenu();
    private boolean seleccionandoProveedor = false;
    private boolean modoCrear = false;

    public ORDENCOMPRA() {
        initComponents();
        IconoVentanaUtil.aplicar(this);
        configurarTablaProductos();
        cargarCFDI();
        cargarUsuario();
        cargarMoneda();
        cargarPagos();
        cargarFormaPago();
        cargarMetodoPago();
        cargarDias();
        cargarCacheInventario();
        cargarCacheProveedores();
        actualizarTotales();
        generarNumeroOrden();
    setLocationRelativeTo(null); 
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
            boolean mostrarCambio = sel.toString().equalsIgnoreCase("USD") || sel.toString().equalsIgnoreCase("EUR");
            jTextField5.setVisible(mostrarCambio);
            jLabel38.setVisible(mostrarCambio);
            actualizarTotales();
        }
    });

    jButton7.addActionListener(e -> eliminarFilaSeleccionada());
    jButton2.addActionListener(e -> buscarOrden());
    jButton3.addActionListener(e -> guardarEnRegistroc());

    // Visibilidad inicial de tipo de cambio según moneda seleccionada
    {
        Object sel = jComboBox2.getSelectedItem();
        boolean mostrarCambio = sel != null && (sel.toString().equalsIgnoreCase("USD") || sel.toString().equalsIgnoreCase("EUR"));
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

    // Usar el mismo x de referencia (jLabel20) para alinear ambos checkboxes
    java.awt.Rectangle rSub = jLabel20.getBounds();
    java.awt.Rectangle rIVA = jLabel21.getBounds();
    int checkX = Math.min(rSub.x, rIVA.x) - 220;
    int fieldH = Math.max(rSub.height + 4, 26);
    jCheckBoxDescuento.setBounds(checkX, rSub.y - 1, 106, fieldH);
    jTextFieldDescuentoPct.setBounds(checkX + 108, rSub.y - 1, 65, fieldH);
    jCheckBoxIVA.setBounds(checkX, rIVA.y - 1, 140, fieldH);
    jPanel1.setComponentZOrder(jCheckBoxIVA, 0);
    jPanel1.setComponentZOrder(jCheckBoxDescuento, 0);
    jPanel1.setComponentZOrder(jTextFieldDescuentoPct, 0);
    jTextFieldDescuentoPct.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
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

    // Botón SALIR DE ORDEN — a la derecha de IMPRIMIR PDF
    jButtonSalirOrden = new javax.swing.JButton("SALIR DE ORDEN");
    jButtonSalirOrden.setBackground(new java.awt.Color(139, 0, 0));
    jButtonSalirOrden.setForeground(java.awt.Color.WHITE);
    java.awt.Rectangle rBtn1 = jButton1.getBounds();
    jButtonSalirOrden.setBounds(rBtn1.x + rBtn1.width + 10, rBtn1.y, 140, rBtn1.height);
    jButtonSalirOrden.addActionListener(e -> salirDeOrden());
    jButtonSalirOrden.setEnabled(false);
    jPanel1.add(jButtonSalirOrden);
    jPanel1.setComponentZOrder(jButtonSalirOrden, 0);

    // Botón GUARDAR CAMBIOS — esquina superior derecha (solo visible para admin)
    jButtonGuardarCambios = new javax.swing.JButton("Guardar Cambios");
    jButtonGuardarCambios.setBackground(new java.awt.Color(0, 128, 0));
    jButtonGuardarCambios.setForeground(java.awt.Color.WHITE);
    java.awt.Rectangle rBtn6 = jButton6.getBounds();
    // Posicionar en la esquina superior derecha alineado verticalmente con < ATRÁS
    jButtonGuardarCambios.setBounds(jPanel1.getPreferredSize().width - 170, rBtn6.y, 155, rBtn6.height);
    jButtonGuardarCambios.addComponentListener(new java.awt.event.ComponentAdapter() {
        @Override public void componentResized(java.awt.event.ComponentEvent e) {
            // reposicionar si el panel cambia de tamaño
            int panelW = jPanel1.getWidth();
            if (panelW > 0) {
                java.awt.Rectangle r = jButtonGuardarCambios.getBounds();
                jButtonGuardarCambios.setLocation(panelW - 170, r.y);
            }
        }
    });
    jPanel1.addComponentListener(new java.awt.event.ComponentAdapter() {
        @Override public void componentResized(java.awt.event.ComponentEvent e) {
            int panelW = jPanel1.getWidth();
            if (panelW > 0 && jButtonGuardarCambios != null) {
                java.awt.Rectangle r = jButtonGuardarCambios.getBounds();
                jButtonGuardarCambios.setLocation(panelW - 170, r.y);
            }
        }
    });
    jButtonGuardarCambios.addActionListener(e -> guardarCambiosOrden());
    jButtonGuardarCambios.setVisible(false);
    jPanel1.add(jButtonGuardarCambios);
    jPanel1.setComponentZOrder(jButtonGuardarCambios, 0);
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
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
        jComboBox9 = new javax.swing.JComboBox<>();
        jComboBox10 = new javax.swing.JComboBox<>();
        jLabel10 = new javax.swing.JLabel();
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

        jTextField1.addActionListener(this::jTextField1ActionPerformed);

        jLabel9.setForeground(new java.awt.Color(0, 95, 131));
        jLabel9.setText("NO.");

        jLabel11.setForeground(new java.awt.Color(0, 95, 131));
        jLabel11.setText("COTIZACIÓN:");

        jLabel13.setForeground(new java.awt.Color(0, 95, 131));
        jLabel13.setText("PROYECTO:");

        jTextField2.addActionListener(this::jTextField2ActionPerformed);

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jLabel14.setForeground(new java.awt.Color(0, 95, 131));
        jLabel14.setText("DIRECCIÓN:");

        jLabel15.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        jLabel15.setVerticalAlignment(javax.swing.SwingConstants.TOP);
        jLabel15.setHorizontalTextPosition(javax.swing.SwingConstants.RIGHT);
        jLabel15.addComponentListener(new java.awt.event.ComponentAdapter() {
            public void componentShown(java.awt.event.ComponentEvent evt) {
                jLabel15ComponentShown(evt);
            }
        });

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
            Class[] types = new Class [] {
                java.lang.Double.class, java.lang.String.class, java.lang.String.class, java.lang.Integer.class, java.lang.Float.class, java.lang.Float.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, true, true, false, false
            };

            public Class getColumnClass(int columnIndex) {
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
        jComboBox3.addActionListener(this::jComboBox3ActionPerformed);

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
        jLabel12.setText("TIEMPO DE ENTREGA:");

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

        jComboBox9.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jComboBox10.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jLabel10.setText("A");

        jLabel27.setText("DIAS");

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
                                        .addComponent(jComboBox9, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(jLabel10)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(jComboBox10, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(jLabel27)
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
                .addContainerGap()
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
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel19)
                    .addComponent(jComboBox8, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel12)
                    .addComponent(jComboBox9, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel10)
                    .addComponent(jComboBox10, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel27))
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
        PRINCIPAL ventana = new PRINCIPAL();
        ventana.setVisible(true); // 👉 abre la nueva ventana

        this.dispose(); // 👉 cierra el login        
    }//GEN-LAST:event_jButton6ActionPerformed

   

    public void buscarProveedor(String idBuscado) {
    try {
        FileInputStream file = new FileInputStream("C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\proveedores.xlsx");
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {

            if (row.getRowNum() == 0) continue;

            org.apache.poi.ss.usermodel.Cell celdaId = row.getCell(0);

            if (celdaId != null) {

                String valorCelda = celdaId.toString();

                if (valorCelda.equalsIgnoreCase(idBuscado)) {
                    String nombre       = obtenerTextoCelda(row.getCell(1));
                    String RFC          = obtenerTextoCelda(row.getCell(2));
                    String telefono     = obtenerTextoCelda(row.getCell(3));
                    String correo       = obtenerTextoCelda(row.getCell(4));
                    String direccion    = obtenerTextoCelda(row.getCell(5));
                    String contacto     = obtenerTextoCelda(row.getCell(6));
                    workbook.close();

                    proveedorTelefono = telefono;
                    proveedorCorreo   = correo;
                    jLabel8.setText(nombre);
                    jLabel15.setText("<html>" + direccion + "</html>");
                    jLabel30.setText("<html>" + RFC + "</html>");
                    jLabel32.setText("<html>" + contacto + "</html>");

                    return;
                }
            }
        }

        jLabel8.setText("Proveedor no encontrado");

        workbook.close();

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
        try (FileInputStream file = new FileInputStream(RUTA_INVENTARIOS);
             Workbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                String codigo = obtenerTextoCelda(row.getCell(0));
                String desc   = obtenerTextoCelda(row.getCell(1));
                if (!codigo.isEmpty()) {
                    cacheInventario.add(new String[]{codigo, desc});
                }
            }
        } catch (Exception e) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo cargar caché de inventario", e);
        }
    }

    private void cargarCacheProveedores() {
        cacheProveedores.clear();
        try (FileInputStream file = new FileInputStream("C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\proveedores.xlsx");
             Workbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                String id     = obtenerTextoCelda(row.getCell(0));
                String nombre = obtenerTextoCelda(row.getCell(1));
                if (!id.isEmpty()) {
                    cacheProveedores.add(new String[]{id, nombre});
                }
            }
        } catch (Exception e) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo cargar caché de proveedores", e);
        }
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
    } else {
        JOptionPane.showMessageDialog(this, "Producto no encontrado en INVENTARIOS.", "Inventarios", JOptionPane.WARNING_MESSAGE);
        jTextField3.requestFocus();
        jTextField3.selectAll();
    }
}
    
    public boolean buscarProducto(String codigoBuscado) {
    try (FileInputStream file = new FileInputStream(RUTA_INVENTARIOS);
         Workbook workbook = new XSSFWorkbook(file)) {

        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }

            String codigo = obtenerTextoCelda(row.getCell(0));

            if (codigo.equalsIgnoreCase(codigoBuscado.trim())) {
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

    // Aplicar tipo de cambio si la moneda seleccionada es USD o EUR
    Object moneda = jComboBox2.getSelectedItem();
    if (moneda != null && (moneda.toString().equalsIgnoreCase("USD") || moneda.toString().equalsIgnoreCase("EUR"))) {
        double tipoCambio = parseNumero(jTextField5.getText());
        if (tipoCambio > 0) {
            base = base / tipoCambio;
        }
    }

    subtotalOriginal = base;

    // Aplicar descuento si el checkbox está activo
    descuentoImporte = 0;
    double subtotal = base;
    boolean aplicarDescuento = jCheckBoxDescuento != null && jCheckBoxDescuento.isSelected();
    if (aplicarDescuento) {
        double pct = parseNumero(jTextFieldDescuentoPct.getText());
        if (pct > 0 && pct <= 100) {
            descuentoImporte = base * pct / 100.0;
            subtotal = base - descuentoImporte;
        }
    }

    // Actualizar texto del label SUB TOTAL con porcentaje si aplica
    if (aplicarDescuento && descuentoImporte > 0) {
        double pct = parseNumero(jTextFieldDescuentoPct.getText());
        jLabel20.setText(String.format("SUB TOTAL (-%.0f%%):", pct));
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
    try {
        FileInputStream file = new FileInputStream(
            "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\PAGOS.xlsx"
        );

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
    try {
        FileInputStream file = new FileInputStream(
            "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\FORMAPAGO.xlsx"
        );

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
    try {
        FileInputStream file = new FileInputStream(
            "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\METODOPAGO.xlsx"
        );

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

    public void cargarDias() {
    try {
        FileInputStream file = new FileInputStream(
            "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\DIAS.xlsx"
        );

        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        jComboBox9.removeAllItems();
        jComboBox10.removeAllItems();

        boolean primeraFila = true;

        for (Row row : sheet) {
            if (primeraFila) {
                primeraFila = false;
                continue;
            }
            org.apache.poi.ss.usermodel.Cell cell = row.getCell(0);
            if (cell != null) {
                String valor = cell.toString().trim();
                // Quitar decimales: "5.0" -> "5"
                if (valor.endsWith(".0")) {
                    valor = valor.substring(0, valor.length() - 2);
                } else {
                    try {
                        valor = String.valueOf((int) Double.parseDouble(valor));
                    } catch (NumberFormatException ignored) {}
                }
                jComboBox9.addItem(valor);
                jComboBox10.addItem(valor);
            }
        }

        workbook.close();
        file.close();

    } catch (Exception e) {
        e.printStackTrace();
    }
}

    public void cargarCFDI() {
    try {
        FileInputStream file = new FileInputStream(
            "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\CFDI.xlsx"
        );

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
    try {
        FileInputStream User = new FileInputStream(
            "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\USUARIOS.xlsx"
        );

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
    try {
        FileInputStream file = new FileInputStream(
            "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\TIPOMONEDA.xlsx"
        );

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
    try {
        FileInputStream file = new FileInputStream("C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\USUARIOS.xlsx");
        Workbook workbook = new XSSFWorkbook(file);
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
        generarPDF();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jTextField1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField1ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField1ActionPerformed

    private void jLabel15ComponentShown(java.awt.event.ComponentEvent evt) {//GEN-FIRST:event_jLabel15ComponentShown
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel15ComponentShown

    private void jComboBox3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox3ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jComboBox3ActionPerformed

    private void jTextField2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField2ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField2ActionPerformed

    // ─────────────────────────────────────────────────────────────────────────
    //  PDF GENERATION
    // ─────────────────────────────────────────────────────────────────────────
    private static final String LOGO_PDF_PATH =
        "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\IMAGENES\\Logo PDF.png";

    private void generarPDF() {
        String errorValidacion = validarCampos();
        if (errorValidacion != null) {
            JOptionPane.showMessageDialog(this, errorValidacion,
                    "Campos incompletos", JOptionPane.WARNING_MESSAGE);
            return;
        }

        String noOrden = jTextField6.getText().trim();
        String cotizacion = jTextField4.getText().trim();
        String nombreArchivo = "OrdenCompra#" + noOrden + "_" + cotizacion + ".pdf";

        javax.swing.JFileChooser fc = new javax.swing.JFileChooser();
        fc.setDialogTitle("Guardar Orden de Compra");
        fc.setSelectedFile(new File(System.getProperty("user.home") + "\\Desktop\\" + nombreArchivo));
        fc.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("PDF (*.pdf)", "pdf"));
        if (fc.showSaveDialog(this) != javax.swing.JFileChooser.APPROVE_OPTION) return;

        String ruta = fc.getSelectedFile().getAbsolutePath();
        if (!ruta.toLowerCase().endsWith(".pdf")) ruta += ".pdf";

        File archivoPDF = new File(ruta);
        if (archivoPDF.exists()) {
            JOptionPane.showMessageDialog(this,
                "El PDF '" + archivoPDF.getName() + "' ya existe y no se generará de nuevo.",
                "PDF ya generado", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        try {
            PdfWriter  writer = new PdfWriter(ruta);
            PdfDocument pdfDoc = new PdfDocument(writer);

            // Aplicar plantilla de fondo en todas las páginas
            String rutaPlantilla = "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\IMAGENES\\PLANTILLAOCRST.pdf";
            try {
                PdfDocument templateDoc = new PdfDocument(new PdfReader(rutaPlantilla));
                PdfFormXObject plantillaXObj = templateDoc.getFirstPage().copyAsFormXObject(pdfDoc);
                templateDoc.close();
                pdfDoc.addEventHandler(PdfDocumentEvent.START_PAGE, event -> {
                    PdfDocumentEvent docEvent = (PdfDocumentEvent) event;
                    PdfPage page = docEvent.getPage();
                    PdfCanvas canvas = new PdfCanvas(page.newContentStreamBefore(), page.getResources(), pdfDoc);
                    com.itextpdf.kernel.geom.Rectangle pageRect = page.getPageSize();
                    canvas.addXObjectFittedIntoRectangle(plantillaXObj, pageRect);
                    canvas.release();
                });
            } catch (Exception ePlantilla) {
                logger.log(java.util.logging.Level.WARNING, "No se pudo cargar la plantilla PDF", ePlantilla);
            }

            Document   doc    = new Document(pdfDoc, PageSize.LETTER);
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

            Image logo = new Image(ImageDataFactory.create(LOGO_PDF_PATH)).scaleToFit(130, 55);
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
                fecha = new SimpleDateFormat("dd/MM/yyyy").format(jDateChooser2.getDate());
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
            String tiempoDesde = jComboBox9.getSelectedItem() != null ? jComboBox9.getSelectedItem().toString() : "";
            String tiempoHasta = jComboBox10.getSelectedItem() != null ? jComboBox10.getSelectedItem().toString() : "";
            String tiempoEntrega = tiempoDesde + " A " + tiempoHasta + " DIAS";
            addFilaPDF(condTbl, bold, plain, azul,
                    "PAGO:", jComboBox6.getSelectedItem() != null ? jComboBox6.getSelectedItem().toString() : "",
                    "TIEMPO DE ENTREGA:", tiempoEntrega);
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
                    .useAllAvailableWidth().setMarginBottom(2);

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
            int filasTot = mostrarDescuento ? 4 : 3;
            Table totTbl = new Table(UnitValue.createPercentArray(new float[]{6, 2.5f, 1.5f}))
                    .useAllAvailableWidth().setMarginTop(2);
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

            // ── FIRMAS (pie de primera página) ─────────────────────────────
            // Posición fija: 36pt del margen izquierdo, 50pt del margen inferior, ancho de página menos márgenes
            float pageWidth  = PageSize.LETTER.getWidth();
            float marginLR   = 36f;
            float tblWidth   = pageWidth - marginLR * 2;
            String elaboro   = jComboBox4.getSelectedItem() != null ? jComboBox4.getSelectedItem().toString() : "";
            String autorizo  = jComboBox5.getSelectedItem() != null ? jComboBox5.getSelectedItem().toString() : "";
            Table firmasTbl = new Table(UnitValue.createPercentArray(new float[]{1, 1}))
                    .setWidth(tblWidth)
                    .setFixedPosition(1, marginLR, 50f, tblWidth);
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

            // ── SEGUNDA HOJA: TÉRMINOS Y CONDICIONES ─────────────────────────
            doc.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
            doc.add(seccionTitulo("TÉRMINOS Y CONDICIONES", bold, azul));
            String rutaTermTxt = "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\TERMINOS\\Términos y Condiciones.txt";
            String contenidoTerminos;
            try {
                byte[] bytesT = java.nio.file.Files.readAllBytes(java.nio.file.Paths.get(rutaTermTxt));
                contenidoTerminos = new String(bytesT, java.nio.charset.StandardCharsets.UTF_8);
            } catch (Exception exT) {
                contenidoTerminos = "(No se pudo cargar el archivo de términos y condiciones)";
            }
            doc.add(new Paragraph(contenidoTerminos)
                    .setFont(plain).setFontSize(8)
                    .setMarginTop(6)
                    .setTextAlignment(TextAlignment.JUSTIFIED));

            doc.close();

            JOptionPane.showMessageDialog(this, "PDF generado exitosamente.", "Listo", JOptionPane.INFORMATION_MESSAGE);
            try { java.awt.Desktop.getDesktop().open(new File(ruta)); } catch (Exception ignored) {}

        } catch (Exception ex) {
            logger.log(java.util.logging.Level.SEVERE, "Error generando PDF", ex);
            JOptionPane.showMessageDialog(this, "Error al generar el PDF:\n" + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
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

        // Tiempo de entrega
        if (!esSeleccionValida(jComboBox9)) {
            faltantes.add("• TIEMPO DE ENTREGA (inicio)");
        }
        if (!esSeleccionValida(jComboBox10)) {
            faltantes.add("• TIEMPO DE ENTREGA (fin)");
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

    private static final String RUTA_REGISTROC =
        "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\REGISTROC.xlsx";

    private static final String RUTA_INTORDENDECOMPRA =
        "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\INTORDENDECOMPRA.xlsx";

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
            jTextField6.setEditable(true);

        } catch (Exception e) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo generar número de orden", e);
            jTextField6.setText("1");
        }
        setModoSoloLectura(false);
    }

    private boolean guardarEnRegistroc() {
        // Verificar si el número de orden ya existe ANTES de pedir más datos
        java.io.File archivo = new java.io.File(RUTA_REGISTROC);
        if (archivo.exists()) {
            String noOrdenBuscar = jTextField6.getText().trim();
            try (FileInputStream fisPre = new FileInputStream(archivo);
                 Workbook wbPre = new XSSFWorkbook(fisPre)) {
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

        Workbook wb;
        Sheet sheet;
        try {
            if (archivo.exists()) {
                try (FileInputStream fis = new FileInputStream(archivo)) {
                    wb = new XSSFWorkbook(fis);
                }
                sheet = wb.getSheetAt(0);
            } else {
                wb = new XSSFWorkbook();
                sheet = wb.createSheet("REGISTROS");
                String[] headers = {"NOORDEN","FECHA","DOCUMENTO","COTIZACIÓN","ID_PROVEEDOR",
                    "CFDI","SOLICITUD","PROYECTO","PAGO","FORMAPAGO","METODOPAGO",
                    "ENTREGAINICIO","ENTREGAFINAL","DESCUENTO","IVA","ELABORO","AUTORIZO",
                    "MONEDA","TIPOCAMBIO","SUBTOTAL","IVA_IMPORTE","TOTAL"};
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
            // ENTREGAINICIO
            Object entregaInicio = jComboBox9.getSelectedItem();
            row.createCell(11).setCellValue(entregaInicio != null ? entregaInicio.toString() : "");
            // ENTREGAFINAL
            Object entregaFinal = jComboBox10.getSelectedItem();
            row.createCell(12).setCellValue(entregaFinal != null ? entregaFinal.toString() : "");
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

            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(archivo)) {
                wb.write(fos);
            }
            wb.close();
            guardarTablaEnIntOrden(jTextField6.getText().trim());
            JOptionPane.showMessageDialog(this,
                    "Orden de compra No. " + jTextField6.getText().trim() + " guardada correctamente.",
                    "Guardado exitoso", JOptionPane.INFORMATION_MESSAGE);
            generarPDF();
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
    private void guardarCambiosOrden() {
        String noOrden = jTextField6.getText().trim();
        if (noOrden.isEmpty()) {
            JOptionPane.showMessageDialog(this, "No hay ninguna orden cargada para guardar.", "Guardar Cambios", JOptionPane.WARNING_MESSAGE);
            return;
        }

        java.io.File archivo = new java.io.File(RUTA_REGISTROC);
        if (!archivo.exists()) {
            JOptionPane.showMessageDialog(this, "No se encontró el archivo de registros.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        try (FileInputStream fis = new FileInputStream(archivo);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
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
                return;
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
            // ENTREGAINICIO
            setCellString(rowToUpdate, 11, jComboBox9.getSelectedItem() != null ? jComboBox9.getSelectedItem().toString() : "");
            // ENTREGAFINAL
            setCellString(rowToUpdate, 12, jComboBox10.getSelectedItem() != null ? jComboBox10.getSelectedItem().toString() : "");
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

            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(archivo)) {
                wb.write(fos);
            }

            guardarTablaEnIntOrden(noOrden);

            JOptionPane.showMessageDialog(this,
                    "Orden № " + noOrden + " actualizada correctamente.",
                    "Guardar Cambios", JOptionPane.INFORMATION_MESSAGE);
            setModoSoloLectura(true);

        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "Error al guardar cambios de orden", ex);
            JOptionPane.showMessageDialog(this, "Error al guardar cambios:\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    /** Utilidad: asigna un valor string a la celda, creándola si no existe */
    private void setCellString(Row row, int col, String value) {
        org.apache.poi.ss.usermodel.Cell cell = row.getCell(col);
        if (cell == null) cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
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

                // Col 11 - ENTREGAINICIO
                seleccionarCombo(jComboBox9, row.getCell(11));

                // Col 12 - ENTREGAFINAL
                seleccionarCombo(jComboBox10, row.getCell(12));

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

                cargarTablaDesdeIntOrden(buscar);
                boolean esAdmin = "admin".equalsIgnoreCase(INICIARSESION.usuarioActual);
                setModoSoloLectura(!esAdmin);
                jButtonSalirOrden.setEnabled(true);
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
                try (FileInputStream fis = new FileInputStream(archivo)) {
                    wb = new XSSFWorkbook(fis);
                }
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
                wb = new XSSFWorkbook();
                sheet = wb.createSheet("INTORDEN");
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

            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(archivo)) {
                wb.write(fos);
            }
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

    /** Limpia el formulario y habilita la creación de una nueva orden */
    private void salirDeOrden() {
        // ── Campos de texto ───────────────────────────────────────────────
        seleccionandoProveedor = true;
        jTextField1.setText("");
        seleccionandoProveedor = false;
        jTextField2.setText("");
        jTextField3.setText("");
        jTextField5.setText("");

        // Restaurar cotización con fecha de hoy
        java.time.LocalDate hoy = java.time.LocalDate.now();
        jTextField4.setText(String.format("AVANTE%04d%02d%02d",
                hoy.getYear(), hoy.getMonthValue(), hoy.getDayOfMonth()));

        // Restaurar fecha a hoy
        jDateChooser2.setDate(new java.util.Date());

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
        if (jComboBox9.getItemCount() > 0) jComboBox9.setSelectedIndex(0);
        if (jComboBox10.getItemCount() > 0) jComboBox10.setSelectedIndex(0);

        // Moneda: restaurar al primer elemento y actualizar descripción / visibilidad tipo cambio
        if (jComboBox2.getItemCount() > 0) {
            jComboBox2.setSelectedIndex(0);
            String clave = jComboBox2.getSelectedItem().toString();
            jLabel37.setText(monedaDescMap.getOrDefault(clave, ""));
            boolean mostrarCambio = clave.equalsIgnoreCase("USD") || clave.equalsIgnoreCase("EUR");
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
        jButtonSalirOrden.setEnabled(false);
    }

    /** Habilita o deshabilita la edición de todos los campos del formulario */
    private void setModoSoloLectura(boolean soloLectura) {
        modoCrear = !soloLectura;
        jTextField1.setEditable(false); // se activa solo al hacer clic (modo crear)
        jTextField2.setEditable(!soloLectura);
        jTextField3.setEditable(!soloLectura);
        jTextField4.setEditable(!soloLectura);
        jTextField5.setEditable(!soloLectura);
        jDateChooser2.setEnabled(!soloLectura);
        jComboBox1.setEnabled(!soloLectura);
        jComboBox2.setEnabled(!soloLectura);
        jComboBox3.setEnabled(!soloLectura);
        jComboBox4.setEnabled(!soloLectura);
        jComboBox5.setEnabled(!soloLectura);
        // jLabel5, jComboBox6-10 editables solo en modo crear
        jComboBox6.setEnabled(!soloLectura);
        jComboBox7.setEnabled(!soloLectura);
        jComboBox8.setEnabled(!soloLectura);
        jComboBox9.setEnabled(!soloLectura);
        jComboBox10.setEnabled(!soloLectura);
        jLabel5.setCursor(soloLectura
                ? java.awt.Cursor.getDefaultCursor()
                : java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR));
        if (jCheckBoxIVA != null) jCheckBoxIVA.setEnabled(!soloLectura);
        if (jCheckBoxDescuento != null) jCheckBoxDescuento.setEnabled(!soloLectura);
        if (jTextFieldDescuentoPct != null) jTextFieldDescuentoPct.setEditable(!soloLectura);
        jTable1.setEnabled(!soloLectura);
        jButton7.setEnabled(!soloLectura);
        jButton3.setEnabled(!soloLectura);
        // Botón Guardar Cambios: visible solo cuando el admin está editando una orden buscada
        boolean esAdmin = "admin".equalsIgnoreCase(INICIARSESION.usuarioActual);
        if (jButtonGuardarCambios != null) {
            jButtonGuardarCambios.setVisible(esAdmin && !soloLectura);
        }
    }

    public static void main(String args[]) {

        java.awt.EventQueue.invokeLater(() -> new ORDENCOMPRA().setVisible(true));
    }

    private javax.swing.JCheckBox jCheckBoxIVA;
    private String proveedorTelefono = "";
    private String proveedorCorreo = "";
    private javax.swing.JCheckBox jCheckBoxDescuento;
    private javax.swing.JTextField jTextFieldDescuentoPct;
    private javax.swing.JButton jButtonSalirOrden;
    private javax.swing.JButton jButtonGuardarCambios;
    private double subtotalOriginal = 0;
    private double descuentoImporte = 0;
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton6;
    private javax.swing.JButton jButton7;
    private javax.swing.JComboBox<String> jComboBox1;
    private javax.swing.JComboBox<String> jComboBox10;
    private javax.swing.JComboBox<String> jComboBox2;
    private javax.swing.JComboBox<String> jComboBox3;
    private javax.swing.JComboBox<String> jComboBox4;
    private javax.swing.JComboBox<String> jComboBox5;
    private javax.swing.JComboBox<String> jComboBox6;
    private javax.swing.JComboBox<String> jComboBox7;
    private javax.swing.JComboBox<String> jComboBox8;
    private javax.swing.JComboBox<String> jComboBox9;
    private com.toedter.calendar.JDateChooser jDateChooser2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
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
