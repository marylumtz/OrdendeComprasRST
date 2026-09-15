/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.IconoVentanaUtil;
import com.mycompany.ocxrst.config.AppConfig;
import com.mycompany.ocxrst.repository.InventarioRepository;
import com.mycompany.ocxrst.repository.ProveedorRepository;
import com.mycompany.ocxrst.service.ProveedorService;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Maria Luisa Martinez
 */
public class PROVEEDORES extends javax.swing.JFrame {
    
    private static final java.util.logging.Logger logger = java.util.logging.Logger.getLogger(PROVEEDORES.class.getName());
    private static final String ORDENES_FILE_ABSOLUTE = AppConfig.registroCFile().getAbsolutePath();
    private static final String ORDENES_DETALLE_FILE_ABSOLUTE = AppConfig.intOrdenCompraFile().getAbsolutePath();
    private static final String[] ENCABEZADOS_REGISTRO_RECEPCION = {
        "ORDENCOMPRA", "PROVEEDOR", "FECHA_RECIBIDO", "MODELO", "CANTIDAD"
    };
    private static final ProveedorRepository PROVEEDOR_REPOSITORY = new ProveedorRepository();
    private static final InventarioRepository INVENTARIO_REPOSITORY = new InventarioRepository();
    private final ProveedorRepository proveedorRepository = new ProveedorRepository();
    private int filaProveedorActual = -1;

    // Caché de proveedores para autocompletar: id -> "id | nombre"
    private final java.util.List<String[]> cacheProveedores = new java.util.ArrayList<>();
    private final javax.swing.JPopupMenu popupProveedores = new javax.swing.JPopupMenu();
    private boolean seleccionandoProveedor = false;

    public static void mostrarAlertaEntregasPendientes(java.awt.Component parent) {
        java.util.List<String> alertasEntrega = obtenerAlertasEntregasPendientes();
        java.util.List<String> alertasPago = obtenerAlertasPagosPendientes();
        if (alertasEntrega.isEmpty() && alertasPago.isEmpty()) {
            return;
        }
        java.util.List<String> secciones = new java.util.ArrayList<>();
        if (!alertasEntrega.isEmpty()) {
            secciones.add("ENTREGAS:\n" + String.join("\n", alertasEntrega));
        }
        if (!alertasPago.isEmpty()) {
            secciones.add("PAGOS:\n" + String.join("\n", alertasPago));
        }
        Object[] opciones = {"HISTORIAL DE COMPRAS", "CANCELAR"};
        JOptionPane alerta = new JOptionPane(
                "Vencimientos o límites dentro de 7 días:\n\n" + String.join("\n\n", secciones),
                JOptionPane.WARNING_MESSAGE,
                JOptionPane.DEFAULT_OPTION,
                null,
                opciones,
                opciones[0]);
        aplicarFondoBlanco(alerta);
        javax.swing.JDialog dialogo = alerta.createDialog(parent, "Alerta de entregas y pagos");
        dialogo.getContentPane().setBackground(java.awt.Color.WHITE);
        dialogo.setVisible(true);
        Object valorSeleccionado = alerta.getValue();
        int seleccion = opciones[0].equals(valorSeleccionado) ? 0 : JOptionPane.CLOSED_OPTION;
        if (seleccion == 0) {
            PROVEEDORES ventana = new PROVEEDORES();
            ventana.setLocationRelativeTo(parent);
            ventana.setVisible(true);
            if (parent instanceof java.awt.Window ventanaAnterior) {
                ventanaAnterior.dispose();
            }
            java.awt.EventQueue.invokeLater(ventana::mostrarHistorialCompras);
        }
    }//mostrarAlertaEntregasPendientes

    private static void aplicarFondoBlanco(java.awt.Component componente) {
        if (componente instanceof javax.swing.JPanel || componente instanceof JOptionPane) {
            componente.setBackground(java.awt.Color.WHITE);
        }
        if (componente instanceof java.awt.Container contenedor) {
            for (java.awt.Component hijo : contenedor.getComponents()) {
                aplicarFondoBlanco(hijo);
            }
        }
    }

    public static int contarAlertasEntregasPendientes() {
        return obtenerAlertasEntregasPendientes().size() + obtenerAlertasPagosPendientes().size();
    }//contarAlertasEntregasPendientes

    public static int contarOrdenesRegistradas() {
        File archivoOrdenes = AppConfig.registroCFile();
        if (!archivoOrdenes.exists()) {
            return 0;
        }
        int total = 0;
        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row row : hoja) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                if (!leerCelda(row, 0).isBlank()) {
                    total++;
                }
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo contar las ordenes registradas", ex);
        }
        return total;
    }//contarOrdenesRegistradas

    public static int contarEntregasPorRevisar() {
        File archivoOrdenes = AppConfig.registroCFile();
        if (!archivoOrdenes.exists()) {
            return 0;
        }
        int total = 0;
        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row row : hoja) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                if (leerCelda(row, 0).isBlank()) {
                    continue;
                }
                if (!"SI".equalsIgnoreCase(leerCelda(row, 24))) {
                    total++;
                }
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo contar las entregas pendientes", ex);
        }
        return total;
    }//contarEntregasPorRevisar

    public static int contarProveedoresActivos() {
        File archivoProveedores = PROVEEDOR_REPOSITORY.obtenerArchivoProveedores();
        if (!archivoProveedores.exists()) {
            return 0;
        }
        int total = 0;
        try (FileInputStream fis = new FileInputStream(archivoProveedores);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row row : hoja) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                if (leerCelda(row, 0).isBlank()) {
                    continue;
                }
                String estado = ProveedorService.normalizarEstado(leerCelda(row, 11));
                if ("1".equals(estado)) {
                    total++;
                }
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo contar los proveedores activos", ex);
        }
        return total;
    }//contarProveedoresActivos

    private static java.util.List<String> obtenerAlertasEntregasPendientes() {
        java.util.List<String> alertasEntrega = new java.util.ArrayList<>();
        java.time.LocalDate hoy = java.time.LocalDate.now();
        File archivoOrdenes = AppConfig.registroCFile();
        if (!archivoOrdenes.exists()) {
            return alertasEntrega;
        }

        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row row : hoja) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                if ("SI".equalsIgnoreCase(leerCelda(row, 24))) {
                    continue;
                }
                java.util.Date fechaEntrega = leerFechaCelda(row, 11);
                if (fechaEntrega == null) {
                    continue;
                }
                java.time.LocalDate limiteEntrega = fechaEntrega.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate();
                long diasFaltantes = java.time.temporal.ChronoUnit.DAYS.between(hoy, limiteEntrega);
                if (diasFaltantes <= 7) {
                    alertasEntrega.add(describirAlertaEntrega(leerCelda(row, 0), diasFaltantes, limiteEntrega));
                }
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo revisar alertas de entregas", ex);
        }
        return alertasEntrega;
    }//obtenerAlertasEntregasPendientes

    private static java.util.List<String> obtenerAlertasPagosPendientes() {
        java.util.List<String> alertasPago = new java.util.ArrayList<>();
        java.time.LocalDate hoy = java.time.LocalDate.now();
        File archivoOrdenes = AppConfig.registroCFile();
        if (!archivoOrdenes.exists()) {
            return alertasPago;
        }

        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row row : hoja) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                String noOrden = leerCelda(row, 0);
                double pendientePago = parseMontoAlerta(leerCelda(row, 21))
                        - parseMontoAlerta(leerCelda(row, 22));
                if (noOrden.isBlank() || pendientePago <= 0.001) {
                    continue;
                }
                java.util.Date fechaPago = leerFechaCelda(row, 12);
                if (fechaPago == null) {
                    continue;
                }
                java.time.LocalDate limitePago = fechaPago.toInstant()
                        .atZone(java.time.ZoneId.systemDefault()).toLocalDate();
                long diasFaltantes = java.time.temporal.ChronoUnit.DAYS.between(hoy, limitePago);
                if (diasFaltantes <= 7) {
                    alertasPago.add(describirAlertaPago(noOrden, diasFaltantes, limitePago, pendientePago));
                }
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudieron revisar alertas de pagos", ex);
        }
        return alertasPago;
    }//obtenerAlertasPagosPendientes

    private static double parseMontoAlerta(String texto) {
        if (texto == null) {
            return 0;
        }
        String limpio = texto.replaceAll("[^0-9.\\-]", "");
        if (limpio.isEmpty()) {
            return 0;
        }
        try {
            return Double.parseDouble(limpio);
        } catch (NumberFormatException ex) {
            return 0;
        }
    }//parseMontoAlerta

    /**
     * Creates new form PRINCIPAL
     */
    public PROVEEDORES() {
        initComponents();
        IconoVentanaUtil.aplicar(this);
        setLocationRelativeTo(null);
        asegurarArchivoProveedores();
        cargarCacheProveedores();
        configurarAutocompleteNombre();
        configurarEventos();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jButton3 = new javax.swing.JButton();
        jButton6 = new javax.swing.JButton();
        jLabelId = new javax.swing.JLabel();
        jTextFieldId = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jTextField1 = new javax.swing.JTextField();
        jTextField2 = new javax.swing.JTextField();
        jTextField3 = new javax.swing.JTextField();
        jTextField4 = new javax.swing.JTextField();
        jTextField5 = new javax.swing.JTextField();
        jTextField6 = new javax.swing.JTextField();
        jButton4 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(255, 255, 255));

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(0, 95, 131));
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("PROVEEDORES");

        jButton1.setBackground(new java.awt.Color(0, 95, 131));
        jButton1.setForeground(new java.awt.Color(255, 255, 255));
        jButton1.setText("CREAR");

        jButton2.setBackground(new java.awt.Color(0, 95, 131));
        jButton2.setForeground(new java.awt.Color(255, 255, 255));
        jButton2.setText("EDITAR");

        jButton3.setBackground(java.awt.Color.WHITE);
        jButton3.setForeground(new java.awt.Color(255, 255, 255));
        jButton3.setText("ELIMINAR");

        jButton6.setBackground(new java.awt.Color(0, 95, 131));
        jButton6.setForeground(new java.awt.Color(255, 255, 255));
        jButton6.setText("< ATRAS");
        jButton6.addActionListener(this::jButton6ActionPerformed);

        jLabelId.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabelId.setText("ID:");

        jTextFieldId.setEditable(false);

        jLabel2.setText("NOMBRE O RAZÓN SOCIAL:");
        jLabel2.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));

        jLabel3.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel3.setText("RFC:");

        jLabel4.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel4.setText("TELÉFONO:");

        jLabel5.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel5.setText("CORREO:");

        jLabel6.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel6.setText("DIRECCIÓN:");

        jLabel7.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel7.setText("CONTACTO:");

        jButton4.setBackground(new java.awt.Color(0, 95, 131));
        jButton4.setForeground(new java.awt.Color(255, 255, 255));
        jButton4.setText("HISTORIAL DE COMPRAS");
        jButton4.addActionListener(this::jButton4ActionPerformed);

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(104, 104, 104)
                .addComponent(jButton4)
                .addGap(18, 18, 18)
                .addComponent(jButton1)
                .addGap(18, 18, 18)
                .addComponent(jButton2)
                .addGap(18, 18, 18)
                .addComponent(jButton3)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(22, 22, 22)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 109, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addComponent(jLabel2))
                            .addComponent(jLabelId, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 145, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel4, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 145, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel5, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 135, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel6, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 150, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel7, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 151, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(18, 18, 18)
                                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jTextFieldId)
                                    .addComponent(jTextField2)
                                    .addComponent(jTextField3)
                                    .addComponent(jTextField4)
                                    .addComponent(jTextField5)
                                    .addComponent(jTextField1, javax.swing.GroupLayout.Alignment.TRAILING))
                                .addGap(60, 60, 60))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(18, 18, 18)
                                .addComponent(jTextField6, javax.swing.GroupLayout.PREFERRED_SIZE, 612, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addContainerGap(60, Short.MAX_VALUE))))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(jButton6)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(21, 21, 21)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel1)
                    .addComponent(jButton6))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabelId)
                    .addComponent(jTextFieldId, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4)
                    .addComponent(jTextField3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel5)
                    .addComponent(jTextField4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel6)
                    .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel7)
                    .addComponent(jTextField6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton1)
                    .addComponent(jButton2)
                    .addComponent(jButton3)
                    .addComponent(jButton4))
                .addContainerGap(22, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
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

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        mostrarHistorialCompras();
    }//GEN-LAST:event_jButton4ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
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
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> new PROVEEDORES().setVisible(true));
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton6;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabelId;
    private javax.swing.JTextField jTextFieldId;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField2;
    private javax.swing.JTextField jTextField3;
    private javax.swing.JTextField jTextField4;
    private javax.swing.JTextField jTextField5;
    private javax.swing.JTextField jTextField6;
    // End of variables declaration//GEN-END:variables

    private void configurarEventos() {
        jButton1.addActionListener(e -> crearProveedor());
        jButton2.addActionListener(e -> editarProveedor());
        jButton3.addActionListener(e -> alternarEstadoProveedor());
        restablecerIndicadorEstado();
    }//configurarEventos

    private void asegurarArchivoProveedores() {
        try {
            proveedorRepository.asegurarArchivoProveedores();
        } catch (IllegalStateException ex) {
            mostrarError("No se pudo preparar el archivo de proveedores.", ex);
        }
    }//asegurarArchivoProveedores

    private void crearProveedor() {
        if (!validarCamposBase()) {
            return;
        }

        String rfc = jTextField2.getText().trim();
        String idCreado = "";

        try {
            File archivo = obtenerArchivoProveedores();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaProveedores(libro);
                asegurarEncabezados(hoja);

                if (buscarFilaProveedorPorRFC(hoja, rfc) >= 0) {
                    JOptionPane.showMessageDialog(this, "Ya existe un proveedor con ese RFC.");
                    return;
                }

                int nuevaFila = primeraFilaDisponible(hoja);
                Row fila = hoja.getRow(nuevaFila);
                if (fila == null) {
                    fila = hoja.createRow(nuevaFila);
                }

                idCreado = generarIdProveedor(hoja, jTextField1.getText().trim());
                escribirCelda(fila, 0, idCreado);
                guardarDatosPrincipalesEnFila(fila);
                completarValoresComercialesPorDefecto(fila);
                escribirCelda(fila, 11, "1");

                guardarLibro(libro, archivo);
                filaProveedorActual = nuevaFila;
                actualizarIndicadorEstado("1");
            }

            jTextFieldId.setText(idCreado);
            JOptionPane.showMessageDialog(this, "Proveedor creado.\nID asignado: " + idCreado + "\nEstado: ACTIVO (1)");
            limpiarFormulario(false);
            cargarCacheProveedores();
        } catch (IOException ex) {
            mostrarError("No se pudo crear el proveedor.", ex);
        }
    }//crearProveedor

    private void editarProveedor() {
        if (!validarCamposBase()) {
            return;
        }

        try {
            File archivo = obtenerArchivoProveedores();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaProveedores(libro);
                asegurarEncabezados(hoja);

                int filaIndex = resolverFilaProveedor(hoja);
                if (filaIndex < 0) {
                    restablecerIndicadorEstado();
                    JOptionPane.showMessageDialog(this, "No se encontro un proveedor para editar. Usa BUSCAR primero.");
                    return;
                }

                String rfcForm = jTextField2.getText().trim();
                int filaMismoRFC = buscarFilaProveedorPorRFC(hoja, rfcForm);
                if (filaMismoRFC >= 0 && filaMismoRFC != filaIndex) {
                    JOptionPane.showMessageDialog(this, "El RFC ya pertenece a otro proveedor.");
                    return;
                }

                Row fila = hoja.getRow(filaIndex);
                if (fila == null) {
                    fila = hoja.createRow(filaIndex);
                }

                if (leerCelda(fila, 0).isEmpty()) {
                    escribirCelda(fila, 0, generarIdProveedor(hoja, jTextField1.getText().trim()));
                }
                guardarDatosPrincipalesEnFila(fila);
                completarValoresComercialesPorDefecto(fila);

                if (leerCelda(fila, 11).isEmpty()) {
                    escribirCelda(fila, 11, "1");
                } else {
                    escribirCelda(fila, 11, normalizarEstado(leerCelda(fila, 11)));
                }

                guardarLibro(libro, archivo);
                filaProveedorActual = filaIndex;
                actualizarIndicadorEstado(leerCelda(fila, 11));

                jTextFieldId.setText(leerCelda(fila, 0));
            }

            JOptionPane.showMessageDialog(this, "Proveedor actualizado correctamente.");
            cargarCacheProveedores();
        } catch (IOException ex) {
            mostrarError("No se pudo editar el proveedor.", ex);
        }
    }//editarProveedor

    private void alternarEstadoProveedor() {
        try {
            File archivo = obtenerArchivoProveedores();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaProveedores(libro);
                asegurarEncabezados(hoja);

                int filaIndex = resolverFilaProveedor(hoja);
                if (filaIndex < 0) {
                    restablecerIndicadorEstado();
                    JOptionPane.showMessageDialog(this, "No se encontro un proveedor para activar/desactivar. Usa BUSCAR primero.");
                    return;
                }

                Row fila = hoja.getRow(filaIndex);
                if (fila == null) {
                    restablecerIndicadorEstado();
                    JOptionPane.showMessageDialog(this, "El registro encontrado esta vacio.");
                    return;
                }

                String estadoActual = normalizarEstado(leerCelda(fila, 11));
                String nuevoEstado = "1".equals(estadoActual) ? "0" : "1";
                escribirCelda(fila, 11, nuevoEstado);

                guardarLibro(libro, archivo);
                actualizarIndicadorEstado(nuevoEstado);

                filaProveedorActual = filaIndex;
            }
        } catch (IOException ex) {
            mostrarError("No se pudo cambiar el estado del proveedor.", ex);
        }
    }//alternarEstadoProveedor

    private void buscarProveedor() {
        String nombre = jTextField1.getText().trim();
        String rfc = jTextField2.getText().trim();

        if (nombre.isEmpty() && rfc.isEmpty()) {
            restablecerIndicadorEstado();
            JOptionPane.showMessageDialog(this, "Ingresa NOMBRE O RAZÓN SOCIAL, o RFC para buscar.");
            return;
        }

        try {
            File archivo = obtenerArchivoProveedores();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaProveedores(libro);
                asegurarEncabezados(hoja);

                int filaIndex = -1;
                if (!nombre.isEmpty()) {
                    filaIndex = buscarFilaProveedorPorNombre(hoja, nombre);
                }
                if (filaIndex < 0 && !rfc.isEmpty()) {
                    filaIndex = buscarFilaProveedorPorRFC(hoja, rfc);
                }

                if (filaIndex < 0) {
                    filaProveedorActual = -1;
                    restablecerIndicadorEstado();
                    JOptionPane.showMessageDialog(this, "No se encontro un proveedor con los datos capturados.");
                    return;
                }

                Row fila = hoja.getRow(filaIndex);
                if (fila == null) {
                    filaProveedorActual = -1;
                    restablecerIndicadorEstado();
                    JOptionPane.showMessageDialog(this, "El registro encontrado esta vacio.");
                    return;
                }

                cargarFilaEnFormulario(fila);
                filaProveedorActual = filaIndex;

                String estado = normalizarEstado(leerCelda(fila, 11));
                actualizarIndicadorEstado(estado);
            }
        } catch (IOException ex) {
            restablecerIndicadorEstado();
            mostrarError("No se pudo buscar el proveedor.", ex);
        }
    }//buscarProveedor

    private void mostrarHistorialCompras() {
        String idProveedor = jTextFieldId.getText().trim();

        java.util.List<Object[]> filas = new java.util.ArrayList<>();
        java.util.Map<String, Double> productosPendientesPorOrden = obtenerProductosPendientesPorOrden();
        java.time.LocalDate hoy = java.time.LocalDate.now();
        double sumaTotal = 0;
        double sumaAbonado = 0;
        String simboloMoneda = "$";
        File archivoOrdenes = new File(ORDENES_FILE_ABSOLUTE);
        if (archivoOrdenes.exists()) {
            try (FileInputStream fis = new FileInputStream(archivoOrdenes);
                 Workbook libro = new XSSFWorkbook(fis)) {
                Sheet hoja = libro.getSheetAt(0);
                for (Row row : hoja) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    String noOrden = leerCelda(row, 0);
                    if (noOrden.isBlank()) {
                        continue;
                    }
                    String idFila = leerCelda(row, 4);
                    if (!idProveedor.isEmpty() && !idProveedor.equalsIgnoreCase(idFila)) {
                        continue;
                    }

                    java.util.Date fecha = leerFechaCelda(row, 1);
                    java.util.Date fechaEntrega = leerFechaCelda(row, 11);
                    String cotizacion = leerCelda(row, 3);
                    String solicitante = leerCelda(row, 6);
                    String total = leerCelda(row, 21);
                    String abonado = leerCelda(row, 22);
                    String estadoRecepcion = leerCelda(row, 24);
                    String diasFaltantesCongelados = leerCelda(row, 25);
                    String estatus = calcularEstatusOrden(total, abonado, estadoRecepcion);
                    String pendiente = calcularPendientePago(total, abonado);
                    sumaTotal += parseMonto(total);
                    sumaAbonado += parseMonto(abonado);
                    if (!total.isEmpty()) {
                        simboloMoneda = obtenerSimboloMoneda(total);
                    }

                    String fechaTexto = "";
                    String diasFaltantesTexto = "";
                    if (fecha != null) {
                        java.time.LocalDate fechaOrden = fecha.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate();
                        fechaTexto = fechaOrden.format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                    }
                    if (fechaEntrega != null) {
                        if ("SI".equalsIgnoreCase(estadoRecepcion) && !diasFaltantesCongelados.isEmpty()) {
                            diasFaltantesTexto = diasFaltantesCongelados;
                        } else {
                            java.time.LocalDate limiteEntrega = fechaEntrega.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate();
                            long diasFaltantes = java.time.temporal.ChronoUnit.DAYS.between(hoy, limiteEntrega);
                            diasFaltantesTexto = String.valueOf(diasFaltantes);
                        }
                    }

                    String productosPendientes = formatearCantidadInventario(
                            productosPendientesPorOrden.getOrDefault(noOrden.toUpperCase(java.util.Locale.ROOT), 0.0));
                    filas.add(new Object[]{idFila, noOrden, fechaTexto, cotizacion, solicitante, total, abonado,
                        pendiente, diasFaltantesTexto, productosPendientes, estatus});
                }
            } catch (Exception ex) {
                mostrarError("No se pudo leer el historial de ordenes de compra.", ex);
                return;
            }
        }

        String[] columnas = {
            "ID PROVEEDOR", "NO. ORDEN", "FECHA", "COTIZACI\u00d3N", "SOLICITANTE", "TOTAL", "ABONADO",
            "PEND X PAGAR", "DIAS FAL_RECIBIR", "PROD X RECIBIR", "ESTATUS"
        };
        javax.swing.table.DefaultTableModel modelo = new javax.swing.table.DefaultTableModel(columnas, 0) {
            @Override
            public boolean isCellEditable(int fila, int columna) {
                return false;
            }
        };
        for (Object[] fila : filas) {
            modelo.addRow(fila);
        }

        javax.swing.JTable tabla = new javax.swing.JTable(modelo) {
            @Override
            public java.awt.Component prepareRenderer(javax.swing.table.TableCellRenderer renderer,
                    int fila, int columna) {
                java.awt.Component componente = super.prepareRenderer(renderer, fila, columna);
                if (!isCellSelected(fila, columna)) {
                    int filaModelo = convertRowIndexToModel(fila);
                    int columnaModelo = convertColumnIndexToModel(columna);
                    String estatus = String.valueOf(getModel().getValueAt(filaModelo, 10)).trim();
                    componente.setBackground(getBackground());
                    componente.setForeground(columnaModelo == 1
                            ? new java.awt.Color(0, 90, 200)
                            : getForeground());
                    try {
                        long dias = Long.parseLong(String.valueOf(getModel().getValueAt(filaModelo, 8)).trim());
                        boolean columnaPendiente = columnaModelo == 7 || columnaModelo == 9;
                        double valorPendiente = columnaPendiente
                            ? parseNumeroOrdenable(getModel().getValueAt(filaModelo, columnaModelo))
                            : 0;
                        if (columnaPendiente && valorPendiente > 0 && dias <= 7
                            && !"FINALIZADO".equalsIgnoreCase(estatus)) {
                            componente.setForeground(new java.awt.Color(170, 0, 0));
                            componente.setBackground(new java.awt.Color(255, 235, 235));
                        }
                    } catch (NumberFormatException ignored) {
                    }
                }
                return componente;
            }
        };
        configurarOrdenamientoHistorial(tabla, modelo);
        tabla.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        configurarColumnaNoOrdenComoEnlace(tabla);
        tabla.removeColumn(tabla.getColumn("DIAS FAL_RECIBIR"));
        ajustarAnchoColumnas(tabla);
        java.awt.Rectangle areaUtil = java.awt.GraphicsEnvironment.getLocalGraphicsEnvironment()
            .getMaximumWindowBounds();
        ajustarAnchoHistorial(tabla, areaUtil.width - 40);

        javax.swing.JLabel etiquetaTotal = new javax.swing.JLabel(
                (idProveedor.isEmpty() ? "Total de compras: " : "Total de compras del proveedor: ") + filas.size());
        etiquetaTotal.setBorder(javax.swing.BorderFactory.createEmptyBorder(8, 8, 4, 8));

        javax.swing.JPanel panelEncabezado = new javax.swing.JPanel();
        panelEncabezado.setLayout(new javax.swing.BoxLayout(panelEncabezado, javax.swing.BoxLayout.Y_AXIS));
        panelEncabezado.add(etiquetaTotal);

        if (!idProveedor.isEmpty()) {
            javax.swing.JLabel etiquetaSumaTotal = new javax.swing.JLabel(
                    String.format("Total de todas las órdenes: %s %,.2f", simboloMoneda, sumaTotal));
            etiquetaSumaTotal.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 8, 4, 8));
            panelEncabezado.add(etiquetaSumaTotal);

            javax.swing.JLabel etiquetaSumaAbonado = new javax.swing.JLabel(
                    String.format("Total abonado: %s %,.2f", simboloMoneda, sumaAbonado));
            etiquetaSumaAbonado.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 8, 8, 8));
            panelEncabezado.add(etiquetaSumaAbonado);
        }

        javax.swing.JButton botonAbonar = new javax.swing.JButton("ABONAR");
        botonAbonar.addActionListener(e -> registrarAbono(tabla));
        javax.swing.JPanel panelBoton = new javax.swing.JPanel(new java.awt.BorderLayout());
        panelBoton.setBorder(javax.swing.BorderFactory.createEmptyBorder(8, 8, 4, 8));
        panelBoton.add(botonAbonar, java.awt.BorderLayout.NORTH);

        javax.swing.JPanel panelSuperior = new javax.swing.JPanel(new java.awt.BorderLayout());
        panelSuperior.add(panelEncabezado, java.awt.BorderLayout.WEST);
        panelSuperior.add(panelBoton, java.awt.BorderLayout.EAST);

        javax.swing.JPanel panel = new javax.swing.JPanel(new java.awt.BorderLayout());
        panel.add(panelSuperior, java.awt.BorderLayout.NORTH);
        panel.add(new javax.swing.JScrollPane(tabla), java.awt.BorderLayout.CENTER);

        int anchoTotal = tabla.getColumnModel().getTotalColumnWidth() + 60;
        int altoTotal = (tabla.getRowCount() + 1) * tabla.getRowHeight() + 120;
        int anchoPanel = Math.max(1200, Math.min(anchoTotal, areaUtil.width - 16));
        int altoPanel = Math.max(600, Math.min(altoTotal, areaUtil.height - 60));
        panel.setPreferredSize(new java.awt.Dimension(anchoPanel, altoPanel));

        javax.swing.JDialog dialogo = new javax.swing.JDialog(this,
                idProveedor.isEmpty() ? "Historial de \u00d3rdenes de Compra - Todos los proveedores"
                        : "Historial de \u00d3rdenes de Compra - " + idProveedor,
                true);
        dialogo.getContentPane().add(panel);
        dialogo.setResizable(true);
        dialogo.pack();
        int anchoDialogo = Math.min(Math.max(dialogo.getWidth(), anchoTotal + 20), areaUtil.width - 16);
        dialogo.setSize(anchoDialogo, dialogo.getHeight());
        dialogo.setLocation(areaUtil.x + (areaUtil.width - anchoDialogo) / 2,
            areaUtil.y + (areaUtil.height - dialogo.getHeight()) / 2);
        dialogo.setVisible(true);
    }//mostrarHistorialCompras

    private static String describirAlertaEntrega(String noOrden, long diasFaltantes, java.time.LocalDate limiteEntrega) {
        String fechaLimite = limiteEntrega.format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        if (diasFaltantes < 0) {
            return "Orden " + noOrden + ": entrega vencida hace " + Math.abs(diasFaltantes) + " día(s). Límite: " + fechaLimite;
        }
        if (diasFaltantes == 0) {
            return "Orden " + noOrden + ": la entrega vence hoy. Límite: " + fechaLimite;
        }
        return "Orden " + noOrden + ": faltan " + diasFaltantes + " día(s) para recibir. Límite: " + fechaLimite;
    }//describirAlertaEntrega

    private static String describirAlertaPago(String noOrden, long diasFaltantes,
            java.time.LocalDate limitePago, double pendientePago) {
        String fechaLimite = limitePago.format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        String saldo = String.format("$ %,.2f", pendientePago);
        if (diasFaltantes < 0) {
            return "Orden " + noOrden + ": pago vencido hace " + Math.abs(diasFaltantes)
                    + " día(s). Límite: " + fechaLimite + ". Pendiente: " + saldo;
        }
        if (diasFaltantes == 0) {
            return "Orden " + noOrden + ": el pago vence hoy. Límite: " + fechaLimite
                    + ". Pendiente: " + saldo;
        }
        return "Orden " + noOrden + ": faltan " + diasFaltantes + " día(s) para pagar. Límite: "
                + fechaLimite + ". Pendiente: " + saldo;
    }//describirAlertaPago

    private void configurarAlertaDiasFaltantes(javax.swing.JTable tabla) {
        final int columnaDias = 8;
        javax.swing.table.DefaultTableCellRenderer renderer = new javax.swing.table.DefaultTableCellRenderer() {
            @Override
            public java.awt.Component getTableCellRendererComponent(javax.swing.JTable t, Object valor,
                    boolean seleccionado, boolean foco, int fila, int columna) {
                java.awt.Component comp = super.getTableCellRendererComponent(t, valor, seleccionado, foco, fila, columna);
                if (!seleccionado) {
                    comp.setForeground(t.getForeground());
                    comp.setBackground(t.getBackground());
                    try {
                        long dias = Long.parseLong(String.valueOf(valor).trim());
                        if (dias <= 7) {
                            comp.setForeground(new java.awt.Color(170, 0, 0));
                            comp.setBackground(new java.awt.Color(255, 235, 235));
                        }
                    } catch (NumberFormatException ignored) {
                    }
                }
                setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
                return comp;
            }
        };
        tabla.getColumnModel().getColumn(columnaDias).setCellRenderer(renderer);
    }//configurarAlertaDiasFaltantes

    private void configurarOrdenamientoHistorial(javax.swing.JTable tabla, javax.swing.table.DefaultTableModel modelo) {
        javax.swing.table.TableRowSorter<javax.swing.table.DefaultTableModel> sorter = new javax.swing.table.TableRowSorter<>(modelo);
        java.util.Comparator<String> textoComparator = String.CASE_INSENSITIVE_ORDER;
        java.util.Comparator<String> numeroComparator = java.util.Comparator.comparingDouble(this::parseNumeroOrdenable);
        java.util.Comparator<String> fechaComparator = java.util.Comparator.comparing(this::parseFechaOrdenable);

        sorter.setComparator(0, textoComparator);
        sorter.setComparator(1, numeroComparator);
        sorter.setComparator(2, fechaComparator);
        sorter.setComparator(3, textoComparator);
        sorter.setComparator(4, textoComparator);
        sorter.setComparator(5, numeroComparator);
        sorter.setComparator(6, numeroComparator);
        sorter.setComparator(7, numeroComparator);
        sorter.setComparator(8, numeroComparator);
        sorter.setComparator(9, numeroComparator);
        sorter.setComparator(10, textoComparator);
        tabla.setRowSorter(sorter);
    }//configurarOrdenamientoHistorial

    private double parseNumeroOrdenable(Object valor) {
        if (valor == null) {
            return 0;
        }
        String texto = valor.toString().trim();
        if (texto.isEmpty()) {
            return 0;
        }
        try {
            return Double.parseDouble(texto.replaceAll("[^0-9.\\-]", ""));
        } catch (NumberFormatException ex) {
            return 0;
        }
    }//parseNumeroOrdenable

    private java.time.LocalDate parseFechaOrdenable(Object valor) {
        if (valor == null || valor.toString().trim().isEmpty()) {
            return java.time.LocalDate.MIN;
        }
        try {
            return java.time.LocalDate.parse(valor.toString().trim(), java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        } catch (java.time.format.DateTimeParseException ex) {
            return java.time.LocalDate.MIN;
        }
    }//parseFechaOrdenable

    private void registrarAbono(javax.swing.JTable tabla) {
        int filaSeleccionada = tabla.getSelectedRow();
        if (filaSeleccionada < 0) {
            JOptionPane.showMessageDialog(this, "Selecciona una orden de la tabla primero.");
            return;
        }

        String noOrden = String.valueOf(tabla.getValueAt(filaSeleccionada, 1)).trim();
        String totalTexto = String.valueOf(tabla.getValueAt(filaSeleccionada, 5)).trim();
        String abonadoTexto = String.valueOf(tabla.getValueAt(filaSeleccionada, 6)).trim();

        double totalNumero = parseMonto(totalTexto);
        double abonadoActual = parseMonto(abonadoTexto);
        double pendiente = totalNumero - abonadoActual;

        if (pendiente <= 0) {
            JOptionPane.showMessageDialog(this, "Esta orden ya est\u00e1 totalmente abonada.");
            return;
        }

        String simbolo = obtenerSimboloMoneda(totalTexto);
        javax.swing.JTextField campoMonto = new javax.swing.JTextField(18);
        javax.swing.JButton botonPdf = new javax.swing.JButton("SUBIR PDF");
        botonPdf.addActionListener(e -> cargarImporteDesdePdf(campoMonto, pendiente));

        javax.swing.JPanel panelMonto = new javax.swing.JPanel(new java.awt.GridBagLayout());
        java.awt.GridBagConstraints restricciones = new java.awt.GridBagConstraints();
        restricciones.gridx = 0;
        restricciones.gridy = 0;
        restricciones.gridwidth = 2;
        restricciones.anchor = java.awt.GridBagConstraints.WEST;
        restricciones.insets = new java.awt.Insets(4, 4, 8, 4);
        panelMonto.add(new javax.swing.JLabel(
            String.format("Pendiente por pagar: %s %,.2f", simbolo, pendiente)), restricciones);
        restricciones.gridy = 1;
        restricciones.gridwidth = 1;
        restricciones.insets = new java.awt.Insets(4, 4, 4, 8);
        panelMonto.add(new javax.swing.JLabel("Cantidad a abonar:"), restricciones);
        restricciones.gridx = 1;
        restricciones.fill = java.awt.GridBagConstraints.HORIZONTAL;
        restricciones.weightx = 1;
        panelMonto.add(campoMonto, restricciones);
        restricciones.gridx = 1;
        restricciones.gridy = 2;
        restricciones.fill = java.awt.GridBagConstraints.NONE;
        restricciones.anchor = java.awt.GridBagConstraints.EAST;
        restricciones.weightx = 0;
        panelMonto.add(botonPdf, restricciones);

        Object[] opciones = {"REGISTRAR ABONO", "CANCELAR"};
        int seleccion = JOptionPane.showOptionDialog(this, panelMonto, "Registrar abono",
            JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, null, opciones, opciones[0]);
        if (seleccion != 0) {
            return;
        }

        double monto = parseMonto(campoMonto.getText());
        if (monto <= 0) {
            JOptionPane.showMessageDialog(this, "La cantidad a abonar debe ser mayor a 0.");
            return;
        }
        if (monto > pendiente + 0.001) {
            JOptionPane.showMessageDialog(this,
                    String.format("La cantidad no puede exceder el pendiente (%s %,.2f).", simbolo, pendiente));
            return;
        }

        double nuevoAbonado = abonadoActual + monto;
        File archivoOrdenes = new File(ORDENES_FILE_ABSOLUTE);
        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            boolean actualizado = false;
            for (Row row : hoja) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                if (noOrden.equalsIgnoreCase(leerCelda(row, 0))) {
                    escribirCelda(row, 22, String.format("%s %,.2f", simbolo, nuevoAbonado));
                    escribirCelda(row, 23, calcularEstatusOrden(
                            leerCelda(row, 21), String.valueOf(nuevoAbonado), leerCelda(row, 24)));
                    actualizado = true;
                    break;
                }
            }
            if (!actualizado) {
                JOptionPane.showMessageDialog(this, "No se encontr\u00f3 la orden en el archivo.");
                return;
            }
            try (FileOutputStream fos = new FileOutputStream(archivoOrdenes)) {
                libro.write(fos);
            }
        } catch (Exception ex) {
            mostrarError("No se pudo registrar el abono.", ex);
            return;
        }

        JOptionPane.showMessageDialog(this, "Abono registrado correctamente.");
        int filaModelo = tabla.convertRowIndexToModel(filaSeleccionada);
        javax.swing.table.TableModel modelo = tabla.getModel();
        String nuevoAbonadoTexto = String.format("%s %,.2f", simbolo, nuevoAbonado);
        modelo.setValueAt(nuevoAbonadoTexto, filaModelo, 6);
        modelo.setValueAt(String.format("%s %,.2f", simbolo, Math.max(0, totalNumero - nuevoAbonado)), filaModelo, 7);
        modelo.setValueAt(calcularEstatusOrden(totalTexto, nuevoAbonadoTexto,
            obtenerEstadoRecepcionOrden(noOrden)), filaModelo, 10);
        tabla.repaint();
        PRINCIPAL.actualizarAlarmasEnTiempoReal();
    }//registrarAbono

    private void cargarImporteDesdePdf(javax.swing.JTextField campoMonto, double pendiente) {
        javax.swing.JFileChooser selector = new javax.swing.JFileChooser();
        selector.setDialogTitle("Seleccionar comprobante de pago");
        selector.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Archivos PDF", "pdf"));
        if (selector.showOpenDialog(this) != javax.swing.JFileChooser.APPROVE_OPTION) {
            return;
        }
        try {
            double importe = extraerImportePdf(selector.getSelectedFile(), pendiente);
            campoMonto.setText(String.format(java.util.Locale.US, "%.2f", importe));
            campoMonto.requestFocusInWindow();
            campoMonto.selectAll();
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo extraer el importe del PDF", ex);
            JOptionPane.showMessageDialog(this,
                    "No se encontró un importe válido en el PDF. Puedes capturarlo manualmente.",
                    "Importe no encontrado", JOptionPane.WARNING_MESSAGE);
        }
    }//cargarImporteDesdePdf

    private double extraerImportePdf(File archivoPdf, double pendiente) throws IOException {
        StringBuilder texto = new StringBuilder();
        try (com.itextpdf.kernel.pdf.PdfReader lector = new com.itextpdf.kernel.pdf.PdfReader(archivoPdf);
             com.itextpdf.kernel.pdf.PdfDocument documento = new com.itextpdf.kernel.pdf.PdfDocument(lector)) {
            for (int pagina = 1; pagina <= documento.getNumberOfPages(); pagina++) {
                texto.append(com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor
                        .getTextFromPage(documento.getPage(pagina))).append('\n');
            }
        }

        java.util.regex.Pattern importeEtiquetado = java.util.regex.Pattern.compile(
                "(?i)(?:importe|monto|total|cantidad)(?:\\s+(?:pagad[oa]|de\\s+pago))?\\s*[:$]?\\s*"
                + "(?:MXN|M\\.?N\\.?)?\\s*\\$?\\s*([0-9][0-9,.]*[0-9]|[0-9])");
        java.util.regex.Matcher coincidencia = importeEtiquetado.matcher(texto);
        while (coincidencia.find()) {
            double importe = parseMontoDocumento(coincidencia.group(1));
            if (importe > 0 && importe <= pendiente + 0.001) {
                return importe;
            }
        }

        java.util.regex.Matcher importeMoneda = java.util.regex.Pattern
                .compile("\\$\\s*([0-9][0-9,.]*[0-9]|[0-9])")
                .matcher(texto);
        double mejorImporte = 0;
        while (importeMoneda.find()) {
            double importe = parseMontoDocumento(importeMoneda.group(1));
            if (importe > mejorImporte && importe <= pendiente + 0.001) {
                mejorImporte = importe;
            }
        }
        if (mejorImporte <= 0) {
            throw new IOException("El PDF no contiene texto con un importe reconocible.");
        }
        return mejorImporte;
    }//extraerImportePdf

    private double parseMontoDocumento(String valor) {
        String limpio = valor == null ? "" : valor.replaceAll("[^0-9,.]", "");
        int ultimaComa = limpio.lastIndexOf(',');
        int ultimoPunto = limpio.lastIndexOf('.');
        if (ultimaComa > ultimoPunto) {
            limpio = limpio.replace(".", "").replace(',', '.');
        } else {
            limpio = limpio.replace(",", "");
        }
        try {
            return Double.parseDouble(limpio);
        } catch (NumberFormatException ex) {
            return 0;
        }
    }//parseMontoDocumento

    private String obtenerEstadoRecepcionOrden(String noOrden) {
        File archivoOrdenes = new File(ORDENES_FILE_ABSOLUTE);
        if (!archivoOrdenes.exists()) {
            return "";
        }
        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            for (Row row : libro.getSheetAt(0)) {
                if (row.getRowNum() > 0 && noOrden.equalsIgnoreCase(leerCelda(row, 0))) {
                    return leerCelda(row, 24);
                }
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo consultar el estado de recepción", ex);
        }
        return "";
    }//obtenerEstadoRecepcionOrden

    private java.util.Map<String, Double> obtenerProductosPendientesPorOrden() {
        java.util.Map<String, Double> pendientes = new java.util.HashMap<>();
        File archivoDetalle = new File(ORDENES_DETALLE_FILE_ABSOLUTE);
        if (!archivoDetalle.exists()) {
            return pendientes;
        }
        try (FileInputStream fis = new FileInputStream(archivoDetalle);
             Workbook libro = new XSSFWorkbook(fis)) {
            for (Row row : libro.getSheetAt(0)) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                String noOrden = leerCelda(row, 0).toUpperCase(java.util.Locale.ROOT);
                if (noOrden.isBlank()) {
                    continue;
                }
                double cantidad = parseNumeroInventario(leerCelda(row, 4));
                double recibido = parseNumeroInventario(leerCelda(row, 7));
                pendientes.merge(noOrden, Math.max(0, cantidad - recibido), Double::sum);
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudieron calcular productos pendientes", ex);
        }
        return pendientes;
    }//obtenerProductosPendientesPorOrden

    private double obtenerProductosPendientesOrden(String noOrden) {
        return obtenerProductosPendientesPorOrden()
                .getOrDefault(noOrden.toUpperCase(java.util.Locale.ROOT), 0.0);
    }//obtenerProductosPendientesOrden

    private void actualizarFilaHistorial(javax.swing.JTable tabla, String noOrden) {
        File archivoOrdenes = new File(ORDENES_FILE_ABSOLUTE);
        if (!archivoOrdenes.exists()) {
            return;
        }
        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            for (Row row : libro.getSheetAt(0)) {
                if (row.getRowNum() == 0 || !noOrden.equalsIgnoreCase(leerCelda(row, 0))) {
                    continue;
                }

                String total = leerCelda(row, 21);
                String abonado = leerCelda(row, 22);
                String estadoRecepcion = leerCelda(row, 24);
                String diasFaltantesTexto = leerCelda(row, 25);
                if (!"SI".equalsIgnoreCase(estadoRecepcion) || diasFaltantesTexto.isEmpty()) {
                    java.util.Date fechaEntrega = leerFechaCelda(row, 11);
                    if (fechaEntrega != null) {
                        java.time.LocalDate limiteEntrega = fechaEntrega.toInstant()
                                .atZone(java.time.ZoneId.systemDefault()).toLocalDate();
                        diasFaltantesTexto = String.valueOf(java.time.temporal.ChronoUnit.DAYS.between(
                                java.time.LocalDate.now(), limiteEntrega));
                    }
                }

                javax.swing.table.TableModel modelo = tabla.getModel();
                for (int filaModelo = 0; filaModelo < modelo.getRowCount(); filaModelo++) {
                    if (noOrden.equalsIgnoreCase(String.valueOf(modelo.getValueAt(filaModelo, 1)).trim())) {
                        modelo.setValueAt(abonado, filaModelo, 6);
                        modelo.setValueAt(calcularPendientePago(total, abonado), filaModelo, 7);
                        modelo.setValueAt(diasFaltantesTexto, filaModelo, 8);
                        modelo.setValueAt(formatearCantidadInventario(
                            obtenerProductosPendientesOrden(noOrden)), filaModelo, 9);
                        modelo.setValueAt(calcularEstatusOrden(total, abonado, estadoRecepcion), filaModelo, 10);
                        tabla.repaint();
                        return;
                    }
                }
                return;
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo actualizar la fila del historial", ex);
        }
    }//actualizarFilaHistorial

    private void configurarColumnaNoOrdenComoEnlace(javax.swing.JTable tabla) {
        final int columnaNoOrden = 1;
        javax.swing.table.DefaultTableCellRenderer rendererEnlace = new javax.swing.table.DefaultTableCellRenderer() {
            @Override
            public java.awt.Component getTableCellRendererComponent(javax.swing.JTable t, Object valor,
                    boolean seleccionado, boolean foco, int fila, int columna) {
                java.awt.Component comp = super.getTableCellRendererComponent(t, valor, seleccionado, foco, fila, columna);
                java.util.Map<java.awt.font.TextAttribute, Object> atributos = new java.util.HashMap<>(comp.getFont().getAttributes());
                atributos.put(java.awt.font.TextAttribute.UNDERLINE, java.awt.font.TextAttribute.UNDERLINE_ON);
                comp.setFont(comp.getFont().deriveFont(atributos));
                if (!seleccionado) {
                    comp.setForeground(new java.awt.Color(0, 90, 200));
                }
                return comp;
            }
        };
        tabla.getColumnModel().getColumn(columnaNoOrden).setCellRenderer(rendererEnlace);

        tabla.addMouseMotionListener(new java.awt.event.MouseMotionAdapter() {
            @Override
            public void mouseMoved(java.awt.event.MouseEvent e) {
                int columna = tabla.columnAtPoint(e.getPoint());
                tabla.setCursor(columna == columnaNoOrden
                        ? java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR)
                        : java.awt.Cursor.getDefaultCursor());
            }
        });

        tabla.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mouseClicked(java.awt.event.MouseEvent e) {
                int fila = tabla.rowAtPoint(e.getPoint());
                int columna = tabla.columnAtPoint(e.getPoint());
                if (fila >= 0 && columna == columnaNoOrden) {
                    mostrarDetalleOrden(String.valueOf(tabla.getValueAt(fila, columnaNoOrden)).trim(), tabla);
                }
            }
        });
    }//configurarColumnaNoOrdenComoEnlace

    private void mostrarDetalleOrden(String noOrden, javax.swing.JTable tablaHistorial) {
        if (noOrden.isEmpty()) {
            return;
        }

        java.util.Map<String, String> datos = new java.util.LinkedHashMap<>();
        File archivoOrdenes = new File(ORDENES_FILE_ABSOLUTE);
        if (archivoOrdenes.exists()) {
            try (FileInputStream fis = new FileInputStream(archivoOrdenes);
                 Workbook libro = new XSSFWorkbook(fis)) {
                Sheet hoja = libro.getSheetAt(0);
                for (Row row : hoja) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    if (!noOrden.equalsIgnoreCase(leerCelda(row, 0))) {
                        continue;
                    }
                    datos.put("NO. ORDEN", leerCelda(row, 0));
                    datos.put("PROVEEDOR", obtenerNombreProveedorPorId(leerCelda(row, 4)));
                    datos.put("FECHA", formatearFechaCelda(leerFechaCelda(row, 1)));
                    datos.put("DOCUMENTO", leerCelda(row, 2));
                    datos.put("COTIZACI\u00d3N", leerCelda(row, 3));
                    datos.put("CFDI", leerCelda(row, 5));
                    datos.put("SOLICITANTE", leerCelda(row, 6));
                    datos.put("PROYECTO", leerCelda(row, 7));
                    datos.put("FORMA DE PAGO", leerCelda(row, 9));
                    datos.put("M\u00c9TODO DE PAGO", leerCelda(row, 10));
                    datos.put("FECHA DE ENTREGA", formatearFechaCelda(leerFechaCelda(row, 11)));
                    datos.put("ELABOR\u00d3", leerCelda(row, 15));
                    datos.put("FECHA LÍMITE DE PAGO", formatearFechaCelda(leerFechaCelda(row, 12)));
                    datos.put("AUTORIZ\u00d3", leerCelda(row, 16));
                    datos.put("MONEDA", leerCelda(row, 17));
                    datos.put("TOTAL", leerCelda(row, 21));
                    datos.put("ABONADO", leerCelda(row, 22));
                    datos.put("PENDIENTE POR PAGAR", calcularPendientePago(leerCelda(row, 21), leerCelda(row, 22)));
                        datos.put("ESTATUS", calcularEstatusOrden(
                            leerCelda(row, 21), leerCelda(row, 22), leerCelda(row, 24)));
                    break;
                }
            } catch (Exception ex) {
                mostrarError("No se pudo leer la orden de compra.", ex);
                return;
            }
        }

        if (datos.isEmpty()) {
            JOptionPane.showMessageDialog(this, "No se encontr\u00f3 informaci\u00f3n de la orden " + noOrden + ".");
            return;
        }

        javax.swing.JPanel panelDatos = new javax.swing.JPanel(new java.awt.GridLayout(0, 4, 12, 6));
        panelDatos.setBorder(javax.swing.BorderFactory.createEmptyBorder(10, 10, 10, 10));
        for (java.util.Map.Entry<String, String> dato : datos.entrySet()) {
            panelDatos.add(new javax.swing.JLabel(dato.getKey() + ":"));
            javax.swing.JLabel valor = new javax.swing.JLabel(dato.getValue().isEmpty() ? "-" : dato.getValue());
            valor.setFont(valor.getFont().deriveFont(java.awt.Font.BOLD));
            panelDatos.add(valor);
        }
        if (datos.size() % 2 != 0) {
            panelDatos.add(new javax.swing.JLabel(""));
            panelDatos.add(new javax.swing.JLabel(""));
        }

        String[] columnasProductos = {"C\u00d3DIGO", "DESCRIPCI\u00d3N", "UNIDAD", "CANTIDAD", "RECIBIDO", "PENDIENTE", "RECIBIR AHORA", "PRECIO", "IMPORTE"};
        javax.swing.table.DefaultTableModel modeloProductos = new javax.swing.table.DefaultTableModel(columnasProductos, 0) {
            @Override
            public boolean isCellEditable(int fila, int columna) {
                return columna == 6 && parseNumeroInventario(String.valueOf(getValueAt(fila, 5))) > 0;
            }
        };

        File archivoDetalle = new File(ORDENES_DETALLE_FILE_ABSOLUTE);
        if (archivoDetalle.exists()) {
            try (FileInputStream fis = new FileInputStream(archivoDetalle);
                 Workbook libro = new XSSFWorkbook(fis)) {
                Sheet hoja = libro.getSheetAt(0);
                asegurarEncabezadosRecepcionDetalle(hoja);
                for (Row row : hoja) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    if (noOrden.equalsIgnoreCase(leerCelda(row, 0))) {
                        double cantidad = parseNumeroInventario(leerCelda(row, 4));
                        double recibido = parseNumeroInventario(leerCelda(row, 7));
                        double pendiente = Math.max(0, cantidad - recibido);
                        modeloProductos.addRow(new Object[]{
                            leerCelda(row, 1), leerCelda(row, 2), leerCelda(row, 3),
                            formatearCantidadInventario(cantidad),
                            formatearCantidadInventario(recibido),
                            formatearCantidadInventario(pendiente),
                            "",
                            leerCelda(row, 5), leerCelda(row, 6)
                        });
                    }
                }
            } catch (Exception ex) {
                mostrarError("No se pudo leer los productos de la orden.", ex);
            }
        }

        javax.swing.JTable tablaProductos = new javax.swing.JTable(modeloProductos) {
            @Override
            public java.awt.Component prepareRenderer(javax.swing.table.TableCellRenderer renderer, int fila, int columna) {
                java.awt.Component componente = super.prepareRenderer(renderer, fila, columna);
                int filaModelo = convertRowIndexToModel(fila);
                boolean recibidoCompleto = parseNumeroInventario(String.valueOf(getModel().getValueAt(filaModelo, 5))) <= 0;
                if (!isCellSelected(fila, columna)) {
                    componente.setBackground(recibidoCompleto ? new java.awt.Color(225, 225, 225) : getBackground());
                    componente.setForeground(recibidoCompleto ? new java.awt.Color(120, 120, 120) : getForeground());
                }
                return componente;
            }
        };
        tablaProductos.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_ALL_COLUMNS);
        tablaProductos.setRowHeight(24);
        tablaProductos.setFillsViewportHeight(true);
        tablaProductos.setGridColor(new java.awt.Color(220, 220, 220));
        tablaProductos.setShowGrid(true);
        tablaProductos.getTableHeader().setFont(tablaProductos.getTableHeader().getFont().deriveFont(java.awt.Font.BOLD));
        ocultarColumnasProductos(tablaProductos, "PRECIO", "IMPORTE");
        ajustarAnchoColumnas(tablaProductos);
        configurarNavegacionRecepcion(tablaProductos);

        javax.swing.JPanel panelProductos = new javax.swing.JPanel(new java.awt.BorderLayout());
        panelProductos.setBorder(javax.swing.BorderFactory.createTitledBorder("PRODUCTOS"));
        javax.swing.JScrollPane scrollProductos = new javax.swing.JScrollPane(tablaProductos);
        scrollProductos.setPreferredSize(new java.awt.Dimension(860, 320));
        panelProductos.add(scrollProductos, java.awt.BorderLayout.CENTER);

        javax.swing.JButton botonRecibir = new javax.swing.JButton("RECIBIR MERCANCIA");
        botonRecibir.setBackground(new java.awt.Color(0, 95, 131));
        botonRecibir.setForeground(java.awt.Color.WHITE);
        botonRecibir.setEnabled(!ordenMercanciaRecibida(noOrden) && hayPendienteRecibir(modeloProductos));
        botonRecibir.addActionListener(e -> {
            if (recibirMercancia(noOrden, modeloProductos)) {
                botonRecibir.setEnabled(!ordenMercanciaRecibida(noOrden) && hayPendienteRecibir(modeloProductos));
                actualizarFilaHistorial(tablaHistorial, noOrden);
            }
        });
        javax.swing.JPanel panelRecibir = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT));
        panelRecibir.add(botonRecibir);
        panelProductos.add(panelRecibir, java.awt.BorderLayout.SOUTH);

        javax.swing.JPanel panelContenido = new javax.swing.JPanel(new java.awt.BorderLayout(10, 10));
        panelContenido.add(panelDatos, java.awt.BorderLayout.NORTH);
        panelContenido.add(panelProductos, java.awt.BorderLayout.CENTER);
        panelContenido.setPreferredSize(new java.awt.Dimension(900, 600));

        javax.swing.JDialog dialogo = new javax.swing.JDialog(this, "Orden de Compra No. " + noOrden, true);
        dialogo.getContentPane().add(panelContenido);
        dialogo.setResizable(true);
        dialogo.pack();
        dialogo.setLocationRelativeTo(this);
        dialogo.setVisible(true);
    }//mostrarDetalleOrden

    private String obtenerNombreProveedorPorId(String idProveedor) {
        if (idProveedor == null || idProveedor.isBlank()) {
            return "";
        }
        File archivo = obtenerArchivoProveedores();
        if (!archivo.exists()) {
            return idProveedor;
        }
        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaProveedores(libro);
            for (Row row : hoja) {
                if (row.getRowNum() > 0 && idProveedor.equalsIgnoreCase(leerCelda(row, 0))) {
                    String nombre = leerCelda(row, 1);
                    return nombre.isEmpty() ? idProveedor : nombre;
                }
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo consultar el nombre del proveedor", ex);
        }
        return idProveedor;
    }//obtenerNombreProveedorPorId

    private void configurarNavegacionRecepcion(javax.swing.JTable tabla) {
        final int columnaRecibirModelo = 6;
        javax.swing.Action avanzarRenglon = new javax.swing.AbstractAction() {
            @Override
            public void actionPerformed(java.awt.event.ActionEvent evento) {
                int filaActual = tabla.getEditingRow() >= 0 ? tabla.getEditingRow() : tabla.getSelectedRow();
                if (tabla.isEditing() && !tabla.getCellEditor().stopCellEditing()) {
                    return;
                }
                int columnaRecibirVista = tabla.convertColumnIndexToView(columnaRecibirModelo);
                int filaSiguiente = filaActual + 1;
                while (filaSiguiente < tabla.getRowCount()) {
                    int filaModelo = tabla.convertRowIndexToModel(filaSiguiente);
                    double pendiente = parseNumeroInventario(String.valueOf(tabla.getModel().getValueAt(filaModelo, 5)));
                    if (pendiente > 0) {
                        break;
                    }
                    filaSiguiente++;
                }
                if (filaSiguiente >= tabla.getRowCount() || columnaRecibirVista < 0) {
                    return;
                }
                tabla.changeSelection(filaSiguiente, columnaRecibirVista, false, false);
                tabla.editCellAt(filaSiguiente, columnaRecibirVista);
                java.awt.Component editor = tabla.getEditorComponent();
                if (editor != null) {
                    editor.requestFocusInWindow();
                }
            }
        };

        String accionAvanzar = "avanzarRenglonRecepcion";
        javax.swing.InputMap mapaTeclas = tabla.getInputMap(javax.swing.JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT);
        mapaTeclas.put(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_ENTER, 0), accionAvanzar);
        mapaTeclas.put(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_TAB, 0), accionAvanzar);
        tabla.getActionMap().put(accionAvanzar, avanzarRenglon);
    }//configurarNavegacionRecepcion

    private String formatearFechaCelda(java.util.Date fecha) {
        if (fecha == null) {
            return "";
        }
        return fecha.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate()
                .format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
    }//formatearFechaCelda

    private void ocultarColumnasProductos(javax.swing.JTable tabla, String... columnas) {
        for (String columna : columnas) {
            try {
                tabla.removeColumn(tabla.getColumn(columna));
            } catch (IllegalArgumentException ignored) {
            }
        }
    }//ocultarColumnasProductos

    private boolean recibirMercancia(String noOrden, javax.swing.table.DefaultTableModel modeloProductos) {
        if (modeloProductos.getRowCount() == 0) {
            JOptionPane.showMessageDialog(this, "La orden no tiene productos para recibir.");
            return false;
        }
        if (!hayPendienteRecibir(modeloProductos)) {
            JOptionPane.showMessageDialog(this, "La mercancia de esta orden ya fue recibida por completo.");
            return false;
        }
        if (ordenMercanciaRecibida(noOrden)) {
            JOptionPane.showMessageDialog(this, "La mercancia de esta orden ya fue recibida por completo.");
            return false;
        }

        String errorRecepcion = validarCantidadesARecibir(modeloProductos);
        if (errorRecepcion != null) {
            JOptionPane.showMessageDialog(this, errorRecepcion, "Recibir mercancia", JOptionPane.WARNING_MESSAGE);
            return false;
        }

        int respuesta = JOptionPane.showConfirmDialog(this,
                "¿Deseas registrar las cantidades capturadas y sumarlas al inventario?",
                "Recibir mercancia",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.QUESTION_MESSAGE);
        if (respuesta != JOptionPane.YES_OPTION) {
            return false;
        }

        try {
            verificarRegistroRecepcionDisponible();
            File archivoInventarios = INVENTARIO_REPOSITORY.obtenerArchivoInventarios();
            try (Workbook libroInventarios = INVENTARIO_REPOSITORY.cargarLibro(archivoInventarios)) {
                Sheet hojaInventarios = INVENTARIO_REPOSITORY.obtenerHojaInventarios(libroInventarios);
                INVENTARIO_REPOSITORY.asegurarEncabezados(hojaInventarios);

                for (int fila = 0; fila < modeloProductos.getRowCount(); fila++) {
                    String codigo = obtenerValorModelo(modeloProductos, fila, 0);
                    if (codigo.isEmpty()) {
                        continue;
                    }
                    String descripcion = obtenerValorModelo(modeloProductos, fila, 1);
                    String unidad = obtenerValorModelo(modeloProductos, fila, 2);
                    double cantidadRecibida = parseNumeroInventario(obtenerValorModelo(modeloProductos, fila, 6));
                    String precio = obtenerValorModelo(modeloProductos, fila, 7);
                    if (cantidadRecibida <= 0) {
                        continue;
                    }

                    Row filaInventario = obtenerOCrearFilaInventario(hojaInventarios, codigo);
                    if (INVENTARIO_REPOSITORY.leerCelda(filaInventario, 1).isEmpty()) {
                        INVENTARIO_REPOSITORY.escribirCelda(filaInventario, 1, descripcion);
                    }
                    if (INVENTARIO_REPOSITORY.leerCelda(filaInventario, 2).isEmpty()) {
                        INVENTARIO_REPOSITORY.escribirCelda(filaInventario, 2, unidad);
                    }
                    if (INVENTARIO_REPOSITORY.leerCelda(filaInventario, 3).isEmpty()) {
                        INVENTARIO_REPOSITORY.escribirCelda(filaInventario, 3, precio);
                    }
                    double cantidadActual = parseNumeroInventario(INVENTARIO_REPOSITORY.leerCelda(filaInventario, 4));
                    INVENTARIO_REPOSITORY.escribirCelda(filaInventario, 4, formatearCantidadInventario(cantidadActual + cantidadRecibida));
                    if (INVENTARIO_REPOSITORY.leerCelda(filaInventario, 5).isEmpty()) {
                        INVENTARIO_REPOSITORY.escribirCelda(filaInventario, 5, "1");
                    }
                }

                INVENTARIO_REPOSITORY.guardarLibro(libroInventarios, archivoInventarios);
            }

            registrarMovimientosRecepcion(noOrden, modeloProductos);
            actualizarRecepcionOrden(noOrden, modeloProductos);
            PRINCIPAL.actualizarAlarmasEnTiempoReal();
            JOptionPane.showMessageDialog(this, "Recepcion registrada en el historial y sumada al inventario correctamente.");
            return true;
        } catch (Exception ex) {
            mostrarError("No se pudo recibir la mercancia.", ex);
            return false;
        }
    }//recibirMercancia

    private void verificarRegistroRecepcionDisponible() throws IOException {
        File archivoRegistro = AppConfig.resolveBaseFile("REGISTRORECIBIR.xlsx");
        if (!archivoRegistro.exists()) {
            return;
        }
        try (java.nio.channels.SeekableByteChannel canal = java.nio.file.Files.newByteChannel(
                archivoRegistro.toPath(), java.nio.file.StandardOpenOption.WRITE)) {
            // Abrir el canal sin modificarlo confirma que Excel no mantiene bloqueado el archivo.
        } catch (IOException ex) {
            throw new IOException("Cierra REGISTRORECIBIR.xlsx antes de registrar la recepción.", ex);
        }
    }//verificarRegistroRecepcionDisponible

    private void registrarMovimientosRecepcion(String noOrden,
            javax.swing.table.DefaultTableModel modeloProductos) throws IOException {
        File archivoRegistro = AppConfig.resolveBaseFile("REGISTRORECIBIR.xlsx");
        Workbook libro;
        if (archivoRegistro.exists() && archivoRegistro.length() > 0) {
            try (FileInputStream fis = new FileInputStream(archivoRegistro)) {
                libro = new XSSFWorkbook(fis);
            }
        } else {
            libro = new XSSFWorkbook();
        }

        try (libro) {
            Sheet hoja = libro.getNumberOfSheets() == 0
                    ? libro.createSheet("REGISTRORECIBIR")
                    : libro.getSheetAt(0);
            Row encabezado = hoja.getRow(0);
            if (encabezado == null) {
                encabezado = hoja.createRow(0);
            }
            for (int columna = 0; columna < ENCABEZADOS_REGISTRO_RECEPCION.length; columna++) {
                encabezado.getCell(columna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        .setCellValue(ENCABEZADOS_REGISTRO_RECEPCION[columna]);
            }

            String proveedor = obtenerProveedorOrden(noOrden);
            java.time.LocalDate fechaRecibido = java.time.LocalDate.now();
            String fechaTexto = fechaRecibido.format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
            int siguienteFila = Math.max(1, hoja.getLastRowNum() + 1);
            for (int fila = 0; fila < modeloProductos.getRowCount(); fila++) {
                double cantidadRecibida = parseNumeroInventario(obtenerValorModelo(modeloProductos, fila, 6));
                if (cantidadRecibida <= 0) {
                    continue;
                }
                Row movimiento = hoja.createRow(siguienteFila++);
                escribirCelda(movimiento, 0, noOrden);
                escribirCelda(movimiento, 1, proveedor);
                escribirCelda(movimiento, 2, fechaTexto);
                escribirCelda(movimiento, 3, obtenerValorModelo(modeloProductos, fila, 0));
                escribirCelda(movimiento, 4, formatearCantidadInventario(cantidadRecibida));
            }

            try (FileOutputStream fos = new FileOutputStream(archivoRegistro)) {
                libro.write(fos);
            }
        }
    }//registrarMovimientosRecepcion

    private String obtenerProveedorOrden(String noOrden) throws IOException {
        File archivoOrdenes = new File(ORDENES_FILE_ABSOLUTE);
        if (!archivoOrdenes.exists()) {
            return "";
        }
        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row row : hoja) {
                if (row.getRowNum() > 0 && noOrden.equalsIgnoreCase(leerCelda(row, 0))) {
                    return leerCelda(row, 4);
                }
            }
        }
        return "";
    }//obtenerProveedorOrden

    private String validarCantidadesARecibir(javax.swing.table.DefaultTableModel modeloProductos) {
        boolean tieneCantidad = false;
        for (int fila = 0; fila < modeloProductos.getRowCount(); fila++) {
            double pendiente = parseNumeroInventario(obtenerValorModelo(modeloProductos, fila, 5));
            double recibirAhora = parseNumeroInventario(obtenerValorModelo(modeloProductos, fila, 6));
            if (recibirAhora <= 0) {
                continue;
            }
            tieneCantidad = true;
            if (recibirAhora > pendiente) {
                return "La cantidad a recibir del producto " + obtenerValorModelo(modeloProductos, fila, 0)
                        + " no puede ser mayor al pendiente.";
            }
        }
        return tieneCantidad ? null : "Captura al menos una cantidad en la columna RECIBIR AHORA.";
    }//validarCantidadesARecibir

    private boolean hayPendienteRecibir(javax.swing.table.DefaultTableModel modeloProductos) {
        for (int fila = 0; fila < modeloProductos.getRowCount(); fila++) {
            if (parseNumeroInventario(obtenerValorModelo(modeloProductos, fila, 5)) > 0) {
                return true;
            }
        }
        return false;
    }//hayPendienteRecibir

    private boolean ordenMercanciaRecibida(String noOrden) {
        File archivoOrdenes = new File(ORDENES_FILE_ABSOLUTE);
        if (!archivoOrdenes.exists()) {
            return false;
        }
        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row row : hoja) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                if (noOrden.equalsIgnoreCase(leerCelda(row, 0))) {
                    return "SI".equalsIgnoreCase(leerCelda(row, 24));
                }
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo verificar recepcion de mercancia", ex);
        }
        return false;
    }//ordenMercanciaRecibida

    private void actualizarRecepcionOrden(String noOrden, javax.swing.table.DefaultTableModel modeloProductos) throws IOException {
        File archivoDetalle = new File(ORDENES_DETALLE_FILE_ABSOLUTE);
        try (FileInputStream fis = new FileInputStream(archivoDetalle);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            asegurarEncabezadosRecepcionDetalle(hoja);

            for (Row row : hoja) {
                if (row.getRowNum() == 0 || !noOrden.equalsIgnoreCase(leerCelda(row, 0))) {
                    continue;
                }
                String codigo = leerCelda(row, 1);
                int filaModelo = buscarFilaModeloPorCodigo(modeloProductos, codigo);
                if (filaModelo < 0) {
                    continue;
                }

                double cantidadOrdenada = parseNumeroInventario(leerCelda(row, 4));
                double recibidoAnterior = parseNumeroInventario(leerCelda(row, 7));
                double recibirAhora = parseNumeroInventario(obtenerValorModelo(modeloProductos, filaModelo, 6));
                double recibidoNuevo = recibidoAnterior + recibirAhora;
                double pendienteNuevo = Math.max(0, cantidadOrdenada - recibidoNuevo);

                escribirCelda(row, 7, formatearCantidadInventario(recibidoNuevo));
                escribirCelda(row, 8, formatearCantidadInventario(pendienteNuevo));
                modeloProductos.setValueAt(formatearCantidadInventario(recibidoNuevo), filaModelo, 4);
                modeloProductos.setValueAt(formatearCantidadInventario(pendienteNuevo), filaModelo, 5);
                modeloProductos.setValueAt("", filaModelo, 6);
            }

            try (FileOutputStream fos = new FileOutputStream(archivoDetalle)) {
                libro.write(fos);
            }
        }

        actualizarEstadoRecepcionOrden(noOrden, hayPendienteRecibir(modeloProductos) ? "PARCIAL" : "SI");
    }//actualizarRecepcionOrden

    private void actualizarEstadoRecepcionOrden(String noOrden, String estadoRecepcion) throws IOException {
        File archivoOrdenes = new File(ORDENES_FILE_ABSOLUTE);
        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            Row encabezado = hoja.getRow(0);
            if (encabezado == null) {
                encabezado = hoja.createRow(0);
            }
            encabezado.getCell(24, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue("MERCANCIA_RECIBIDA");
            encabezado.getCell(25, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue("DIAS_FALTANTES_AL_RECIBIR");

            for (Row row : hoja) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                if (noOrden.equalsIgnoreCase(leerCelda(row, 0))) {
                    escribirCelda(row, 24, estadoRecepcion);
                    escribirCelda(row, 23, calcularEstatusOrden(
                            leerCelda(row, 21), leerCelda(row, 22), estadoRecepcion));
                    if ("SI".equalsIgnoreCase(estadoRecepcion) && leerCelda(row, 25).isEmpty()) {
                        java.util.Date fechaEntrega = leerFechaCelda(row, 11);
                        if (fechaEntrega != null) {
                            java.time.LocalDate limiteEntrega = fechaEntrega.toInstant()
                                    .atZone(java.time.ZoneId.systemDefault()).toLocalDate();
                            long diasFaltantes = java.time.temporal.ChronoUnit.DAYS.between(
                                    java.time.LocalDate.now(), limiteEntrega);
                            escribirCelda(row, 25, String.valueOf(diasFaltantes));
                        }
                    }
                    break;
                }
            }

            try (FileOutputStream fos = new FileOutputStream(archivoOrdenes)) {
                libro.write(fos);
            }
        }
    }//actualizarEstadoRecepcionOrden

    private String calcularEstatusOrden(String total, String abonado, String estadoRecepcion) {
        double totalNumero = parseMonto(total);
        double abonadoNumero = parseMonto(abonado);
        boolean pagoCompleto = totalNumero > 0 && abonadoNumero >= totalNumero - 0.001;
        boolean mercanciaCompleta = "SI".equalsIgnoreCase(estadoRecepcion);

        if (pagoCompleto && mercanciaCompleta) {
            return "FINALIZADO";
        }
        if (mercanciaCompleta) {
            return "PENDIENTE DE PAGO";
        }
        if (pagoCompleto) {
            return "PENDIENTE DE MERCANCIA";
        }

        boolean huboAbono = abonadoNumero > 0;
        boolean huboRecepcion = "PARCIAL".equalsIgnoreCase(estadoRecepcion);
        return huboAbono || huboRecepcion ? "PENDIENTE PAGO Y MERCANCIA" : "PROGRESO";
    }//calcularEstatusOrden

    private void asegurarEncabezadosRecepcionDetalle(Sheet hoja) {
        Row encabezado = hoja.getRow(0);
        if (encabezado == null) {
            encabezado = hoja.createRow(0);
        }
        encabezado.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue("CANTIDAD_RECIBIDA");
        encabezado.getCell(8, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue("PENDIENTE_RECIBIR");
    }//asegurarEncabezadosRecepcionDetalle

    private int buscarFilaModeloPorCodigo(javax.swing.table.DefaultTableModel modelo, String codigo) {
        for (int fila = 0; fila < modelo.getRowCount(); fila++) {
            if (codigo.equalsIgnoreCase(obtenerValorModelo(modelo, fila, 0))) {
                return fila;
            }
        }
        return -1;
    }//buscarFilaModeloPorCodigo

    private Row obtenerOCrearFilaInventario(Sheet hojaInventarios, String codigo) {
        for (int fila = 1; fila <= hojaInventarios.getLastRowNum(); fila++) {
            Row row = hojaInventarios.getRow(fila);
            if (row == null) {
                continue;
            }
            if (codigo.equalsIgnoreCase(INVENTARIO_REPOSITORY.leerCelda(row, 0))) {
                return row;
            }
        }

        int nuevaFila = hojaInventarios.getLastRowNum() + 1;
        Row row = hojaInventarios.createRow(nuevaFila);
        INVENTARIO_REPOSITORY.escribirCelda(row, 0, codigo);
        return row;
    }//obtenerOCrearFilaInventario

    private String obtenerValorModelo(javax.swing.table.DefaultTableModel modelo, int fila, int columna) {
        Object valor = modelo.getValueAt(fila, columna);
        return valor == null ? "" : valor.toString().trim();
    }//obtenerValorModelo

    private double parseNumeroInventario(String valor) {
        if (valor == null || valor.trim().isEmpty()) {
            return 0;
        }
        try {
            return Double.parseDouble(valor.trim().replace(",", ""));
        } catch (NumberFormatException ex) {
            return 0;
        }
    }//parseNumeroInventario

    private String formatearCantidadInventario(double cantidad) {
        if (Math.rint(cantidad) == cantidad) {
            return String.valueOf((long) cantidad);
        }
        return String.valueOf(cantidad);
    }//formatearCantidadInventario

    private void ajustarAnchoColumnas(javax.swing.JTable tabla) {
        javax.swing.table.TableColumnModel modeloColumnas = tabla.getColumnModel();
        for (int columna = 0; columna < tabla.getColumnCount(); columna++) {
            javax.swing.table.TableColumn tableColumn = modeloColumnas.getColumn(columna);
            javax.swing.table.TableCellRenderer rendererEncabezado = tabla.getTableHeader().getDefaultRenderer();
            int ancho = rendererEncabezado
                    .getTableCellRendererComponent(tabla, tableColumn.getHeaderValue(), false, false, 0, columna)
                    .getPreferredSize().width;

            for (int fila = 0; fila < tabla.getRowCount(); fila++) {
                javax.swing.table.TableCellRenderer renderer = tabla.getCellRenderer(fila, columna);
                int anchoCelda = tabla.prepareRenderer(renderer, fila, columna).getPreferredSize().width;
                ancho = Math.max(ancho, anchoCelda);
            }

            tableColumn.setPreferredWidth(ancho + 16);
        }
    }//ajustarAnchoColumnas

    private void ajustarAnchoHistorial(javax.swing.JTable tabla, int anchoDisponible) {
        javax.swing.table.TableColumnModel columnas = tabla.getColumnModel();
        int ultimaColumna = columnas.getColumnCount() - 1;
        int anchoEstatus = columnas.getColumn(ultimaColumna).getPreferredWidth();
        int anchoOtrasColumnas = 0;

        for (int columna = 0; columna < ultimaColumna; columna++) {
            anchoOtrasColumnas += columnas.getColumn(columna).getPreferredWidth();
        }

        double factor = Math.min(1.0, (double) (anchoDisponible - anchoEstatus) / anchoOtrasColumnas);
        for (int columna = 0; columna < ultimaColumna; columna++) {
            javax.swing.table.TableColumn tableColumn = columnas.getColumn(columna);
            int anchoEncabezado = tabla.getTableHeader().getDefaultRenderer()
                    .getTableCellRendererComponent(tabla, tableColumn.getHeaderValue(), false, false, 0, columna)
                    .getPreferredSize().width + 12;
            int anchoAjustado = Math.max(anchoEncabezado,
                    (int) Math.floor(tableColumn.getPreferredWidth() * factor));
            tableColumn.setPreferredWidth(anchoAjustado);
            tableColumn.setWidth(anchoAjustado);
        }
        columnas.getColumn(ultimaColumna).setWidth(anchoEstatus);
    }//ajustarAnchoHistorial

    private String calcularPendientePago(String total, String abonado) {
        double totalNumero = parseMonto(total);
        double abonadoNumero = parseMonto(abonado);
        double pendiente = totalNumero - abonadoNumero;
        String simbolo = obtenerSimboloMoneda(total);
        return String.format("%s %,.2f", simbolo, pendiente);
    }//calcularPendientePago

    private double parseMonto(String texto) {
        if (texto == null) {
            return 0;
        }
        String limpio = texto.replaceAll("[^0-9.\\-]", "");
        if (limpio.isEmpty()) {
            return 0;
        }
        try {
            return Double.parseDouble(limpio);
        } catch (NumberFormatException ex) {
            return 0;
        }
    }//parseMonto

    private String obtenerSimboloMoneda(String texto) {
        if (texto == null) {
            return "$";
        }
        String simbolo = texto.replaceAll("[0-9.,\\s]", "").trim();
        return simbolo.isEmpty() ? "$" : simbolo;
    }//obtenerSimboloMoneda

    private static java.util.Date leerFechaCelda(Row row, int columna) {
        Cell celda = row.getCell(columna, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (celda == null) {
            return null;
        }
        try {
            if (celda.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC
                    && org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(celda)) {
                return celda.getDateCellValue();
            }
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo leer fecha de orden", ex);
        }
        return null;
    }//leerFechaCelda

    private static String leerCelda(Row fila, int columna) {
        return PROVEEDOR_REPOSITORY.leerCelda(fila, columna);
    }//leerCelda

    private boolean validarCamposBase() {
        if (jTextField1.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "NOMBRE O RAZON SOCIAL es obligatorio.");
            return false;
        }
        if (jTextField2.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "RFC es obligatorio.");
            return false;
        }
        return true;
    }//validarCamposBase

    private int resolverFilaProveedor(Sheet hoja) {
        if (filaProveedorActual > 0 && filaProveedorActual <= hoja.getLastRowNum()) {
            Row fila = hoja.getRow(filaProveedorActual);
            if (fila != null) {
                String rfcCapturado = jTextField2.getText().trim();
                String rfcFila = leerCelda(fila, 2);
                if (rfcCapturado.isEmpty() || rfcCapturado.equalsIgnoreCase(rfcFila)) {
                    return filaProveedorActual;
                }
            }
        }

        String rfc = jTextField2.getText().trim();
        if (!rfc.isEmpty()) {
            int filaPorRFC = buscarFilaProveedorPorRFC(hoja, rfc);
            if (filaPorRFC >= 0) {
                return filaPorRFC;
            }
        }

        String nombre = jTextField1.getText().trim();
        if (!nombre.isEmpty()) {
            int filaPorNombre = buscarFilaProveedorPorNombre(hoja, nombre);
            if (filaPorNombre >= 0) {
                return filaPorNombre;
            }
        }

        return -1;
    }//resolverFilaProveedor

    private void guardarDatosPrincipalesEnFila(Row fila) {
        escribirCelda(fila, 1, jTextField1.getText().trim());
        escribirCelda(fila, 2, jTextField2.getText().trim());
        escribirCelda(fila, 3, jTextField3.getText().trim());
        escribirCelda(fila, 4, jTextField4.getText().trim());
        escribirCelda(fila, 5, jTextField5.getText().trim());
        escribirCelda(fila, 6, jTextField6.getText().trim());
    }//guardarDatosPrincipalesEnFila

    private void completarValoresComercialesPorDefecto(Row fila) {
        if (leerCelda(fila, 7).isEmpty()) {
            escribirCelda(fila, 7, "CONTADO");
        }
        if (leerCelda(fila, 8).isEmpty()) {
            escribirCelda(fila, 8, "PUE(PAGO EN UNA SOLA EXHIBICION)");
        }
        if (leerCelda(fila, 9).isEmpty()) {
            escribirCelda(fila, 9, "De 3 a 5 dias");
        }
        if (leerCelda(fila, 10).isEmpty()) {
            escribirCelda(fila, 10, "Transferencia Electronica Fondos");
        }
    }//completarValoresComercialesPorDefecto

    private void cargarFilaEnFormulario(Row fila) {
        jTextFieldId.setText(leerCelda(fila, 0));
        jTextField1.setText(leerCelda(fila, 1));
        jTextField2.setText(leerCelda(fila, 2));
        jTextField3.setText(leerCelda(fila, 3));
        jTextField4.setText(leerCelda(fila, 4));
        jTextField5.setText(leerCelda(fila, 5));
        jTextField6.setText(leerCelda(fila, 6));
    }//cargarFilaEnFormulario

    private void limpiarFormulario(boolean conservarNombre) {
        jTextFieldId.setText("");
        if (!conservarNombre) {
            jTextField1.setText("");
        }
        jTextField2.setText("");
        jTextField3.setText("");
        jTextField4.setText("");
        jTextField5.setText("");
        jTextField6.setText("");
        filaProveedorActual = -1;
        restablecerIndicadorEstado();
    }//limpiarFormulario

    private void configurarAutocompleteNombre() {
        jTextField1.addFocusListener(new java.awt.event.FocusAdapter() {
            @Override
            public void focusLost(java.awt.event.FocusEvent e) {
                if (!e.isTemporary()) {
                    popupProveedores.setVisible(false);
                    if (jTextField1.getText().trim().isEmpty()) {
                        limpiarFormulario(false);
                    }
                }
            }
        });

        jTextField1.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
            public void insertUpdate(javax.swing.event.DocumentEvent e) { filtrarProveedoresNombre(); }
            public void removeUpdate(javax.swing.event.DocumentEvent e) { filtrarProveedoresNombre(); }
            public void changedUpdate(javax.swing.event.DocumentEvent e) { filtrarProveedoresNombre(); }
        });
    }//configurarAutocompleteNombre

    private void cargarCacheProveedores() {
        cacheProveedores.clear();
        try {
            File archivo = obtenerArchivoProveedores();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaProveedores(libro);
                for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
                    Row row = hoja.getRow(fila);
                    if (row == null) {
                        continue;
                    }
                    String id = leerCelda(row, 0);
                    String nombre = leerCelda(row, 1);
                    if (!nombre.isEmpty()) {
                        cacheProveedores.add(new String[]{id, nombre});
                    }
                }
            }
        } catch (IOException ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo cargar cache de proveedores", ex);
        }
    }//cargarCacheProveedores

    private void filtrarProveedoresNombre() {
        if (seleccionandoProveedor) {
            return;
        }
        String texto = jTextField1.getText().trim().toLowerCase();
        popupProveedores.setVisible(false);
        popupProveedores.removeAll();
        if (texto.isEmpty()) {
            return;
        }

        int count = 0;
        for (String[] prov : cacheProveedores) {
            if (prov[1].toLowerCase().contains(texto)) {
                String etiqueta = prov[0] + "  |  " + prov[1];
                javax.swing.JMenuItem item = new javax.swing.JMenuItem(etiqueta);
                final String nombre = prov[1];
                item.addActionListener(ev -> {
                    seleccionandoProveedor = true;
                    popupProveedores.setVisible(false);
                    jTextField1.setText(nombre);
                    buscarProveedor();
                    seleccionandoProveedor = false;
                });
                popupProveedores.add(item);
                if (++count == 10) {
                    break;
                }
            }
        }

        if (count > 0) {
            popupProveedores.show(jTextField1, 0, jTextField1.getHeight());
            jTextField1.requestFocusInWindow();
        }
    }//filtrarProveedoresNombre

    private int buscarFilaProveedorPorId(Sheet hoja, String id) {
        for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
            Row row = hoja.getRow(fila);
            if (row == null) {
                continue;
            }
            String idActual = leerCelda(row, 0);
            if (id.equalsIgnoreCase(idActual)) {
                return fila;
            }
        }
        return -1;
    }//buscarFilaProveedorPorId

    private int buscarFilaProveedorPorNombre(Sheet hoja, String nombre) {
        for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
            Row row = hoja.getRow(fila);
            if (row == null) {
                continue;
            }
            String nombreActual = leerCelda(row, 1);
            if (nombre.equalsIgnoreCase(nombreActual)) {
                return fila;
            }
        }
        return -1;
    }//buscarFilaProveedorPorNombre

    private int buscarFilaProveedorPorRFC(Sheet hoja, String rfc) {
        for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
            Row row = hoja.getRow(fila);
            if (row == null) {
                continue;
            }
            String rfcActual = leerCelda(row, 2);
            if (rfc.equalsIgnoreCase(rfcActual)) {
                return fila;
            }
        }
        return -1;
    }//buscarFilaProveedorPorRFC

    private int primeraFilaDisponible(Sheet hoja) {
        for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
            Row row = hoja.getRow(fila);
            if (row == null || leerCelda(row, 0).isEmpty()) {
                return fila;
            }
        }
        return hoja.getLastRowNum() + 1;
    }//primeraFilaDisponible

    private String generarIdProveedor(Sheet hoja, String nombreProveedor) {
        String base = nombreProveedor == null ? "" : nombreProveedor.replaceAll("[^A-Za-z0-9]", "").toLowerCase();
        if (base.isEmpty()) {
            base = "prov";
        }

        if (base.length() > 4) {
            base = base.substring(0, 4);
        }

        while (base.length() < 4) {
            base = base + "x";
        }

        for (int i = 1; i <= 999; i++) {
            String sufijo = i <= 99 ? String.format("%02d", i) : String.valueOf(i);
            String candidato = base + sufijo;
            if (buscarFilaProveedorPorId(hoja, candidato) < 0) {
                return candidato;
            }
        }

        return base + System.currentTimeMillis();
    }//generarIdProveedor

    private boolean asegurarEncabezados(Sheet hoja) {
        return proveedorRepository.asegurarEncabezados(hoja);
    }//asegurarEncabezados

    private Sheet obtenerHojaProveedores(Workbook libro) {
        return proveedorRepository.obtenerHojaProveedores(libro);
    }//obtenerHojaProveedores

    private Workbook cargarLibro(File archivo) throws IOException {
        return proveedorRepository.cargarLibro(archivo);
    }//cargarLibro

    private void guardarLibro(Workbook libro, File archivo) throws IOException {
        proveedorRepository.guardarLibro(libro, archivo);
    }//guardarLibro

    private File obtenerArchivoProveedores() {
        return proveedorRepository.obtenerArchivoProveedores();
    }//obtenerArchivoProveedores

    private void escribirCelda(Row fila, int columna, String valor) {
        proveedorRepository.escribirCelda(fila, columna, valor);
    }//escribirCelda

    private String normalizarEstado(String valor) {
        return ProveedorService.normalizarEstado(valor);
    }//normalizarEstado

    private void actualizarIndicadorEstado(String estado) {
        String estadoNormalizado = normalizarEstado(estado);
        jButton3.setBackground(java.awt.Color.WHITE);
        if ("1".equals(estadoNormalizado)) {
            jButton3.setText("DESACTIVAR (ACTIVO)");
            jButton3.setForeground(new java.awt.Color(0, 120, 0));
            jButton3.setToolTipText("Estado actual: ACTIVO (1)");
            return;
        }

        jButton3.setText("ACTIVAR (INACTIVO)");
        jButton3.setForeground(new java.awt.Color(180, 0, 0));
        jButton3.setToolTipText("Estado actual: INACTIVO (0)");
    }//actualizarIndicadorEstado

    private void restablecerIndicadorEstado() {
        jButton3.setText("ACTIVAR/DESACTIVAR");
        jButton3.setBackground(java.awt.Color.WHITE);
        jButton3.setForeground(new java.awt.Color(0, 0, 0));
        jButton3.setToolTipText("Estado actual: sin seleccionar");
    }//restablecerIndicadorEstado

    private void mostrarError(String mensajeUsuario, Exception ex) {
        logger.log(java.util.logging.Level.SEVERE, mensajeUsuario, ex);
        JOptionPane.showMessageDialog(this, mensajeUsuario + "\nDetalle tecnico: " + ex.getMessage());
    }//mostrarError
}
