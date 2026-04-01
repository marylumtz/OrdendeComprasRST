/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.ocxrst;

import java.awt.Image;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import javax.swing.ImageIcon;
import javax.swing.JComboBox;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Maria Luisa Martinez
 */
public class USUARIOS extends javax.swing.JFrame {

    private static final java.util.logging.Logger logger = java.util.logging.Logger.getLogger(USUARIOS.class.getName());
    private static final String USUARIOS_FILE_ABSOLUTE = "C:\\OCXRST\\OrdendeComprasRST\\src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\USUARIOS.xlsx";
    private static final String USUARIOS_FILE_RELATIVE = "src\\main\\java\\com\\mycompany\\ocxrst\\BASES\\USUARIOS.xlsx";
    private static final String[] ENCABEZADOS = {"USUARIO", "CONTRASENA", "NOMBRE", "AREA", "PUESTO", "CORREO", "FIRMA"};
    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    /**
     * Creates new form PRINCIPAL
     */
    public USUARIOS() {
        initComponents();
        IconoVentanaUtil.aplicar(this);
        setLocationRelativeTo(null);
        asegurarArchivoUsuarios();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jButton3 = new javax.swing.JButton();
        jButton6 = new javax.swing.JButton();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jTextField1 = new javax.swing.JTextField();
        jTextField2 = new javax.swing.JTextField();
        jTextField5 = new javax.swing.JTextField();
        jTextField6 = new javax.swing.JTextField();
        jButton4 = new javax.swing.JButton();
        jComboBox1 = new javax.swing.JComboBox<>();
        jComboBox2 = new javax.swing.JComboBox<>();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(0, 0, 153));
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("USUARIOS");

        jButton1.setText("CREAR");
        jButton1.addActionListener(this::jButton1ActionPerformed);

        jButton2.setText("EDITAR");
        jButton2.addActionListener(this::jButton2ActionPerformed);

        jButton3.setText("ELIMINAR");
        jButton3.addActionListener(this::jButton3ActionPerformed);

        jButton6.setBackground(new java.awt.Color(0, 0, 153));
        jButton6.setForeground(new java.awt.Color(255, 255, 255));
        jButton6.setText("< ATRAS");
        jButton6.addActionListener(this::jButton6ActionPerformed);

        jLabel2.setText("NOMBRE COMPLETO:");
        jLabel2.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));

        jLabel3.setText("AREA O DEPARTAMENTO:");

        jLabel4.setText("PUESTO:");

        jLabel5.setText("CORREO:");

        jLabel6.setText("ID_USUARIO:");

        jLabel7.setText("FIRMA:");

        jButton4.setText("BUSCAR");
        jButton4.addActionListener(this::jButton4ActionPerformed);

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "DIRECCION", "GERENCIA", "CORDINACION", "RECURSOS HUMANOS", "TI" }));
        jComboBox1.addActionListener(this::jComboBox1ActionPerformed);

        jComboBox2.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "DIRECTOR", "GERENTE", "SUB-GERENTE", "CORDINADOR", "AUXILIAR", "AYUDANTE GERENCIAL", "DESARROLLADOR", "RH" }));

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
                .addComponent(jButton6)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGap(82, 82, 82))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel7, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel5, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel4, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel3, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel2, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel6)
                                .addGap(24, 24, 24)))
                        .addGap(18, 18, 18)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, 255, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                .addComponent(jTextField2)
                                .addComponent(jTextField5)
                                .addComponent(jTextField6, javax.swing.GroupLayout.DEFAULT_SIZE, 255, Short.MAX_VALUE)
                                .addComponent(jComboBox1, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jComboBox2, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                        .addGap(60, 60, 60))))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(21, 21, 21)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel1)
                    .addComponent(jButton6))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel6))
                .addGap(15, 15, 15)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4)
                    .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel5))
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
                .addContainerGap(19, Short.MAX_VALUE))
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
        ventana.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_jButton6ActionPerformed

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        buscarUsuario();
    }//GEN-LAST:event_jButton4ActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        crearUsuario();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        editarUsuario();
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        eliminarUsuario();
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
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> new USUARIOS().setVisible(true));
    }//GEN-LAST:event_jButton3ActionPerformed

    private void jComboBox1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox1ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jComboBox1ActionPerformed

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton6;
    private javax.swing.JComboBox<String> jComboBox1;
    private javax.swing.JComboBox<String> jComboBox2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField2;
    private javax.swing.JTextField jTextField5;
    private javax.swing.JTextField jTextField6;
    // End of variables declaration//GEN-END:variables

    private void aplicarIconoVentana() {
        URL logoUrl = getClass().getResource("/com/mycompany/ocxrst/IMAGENES/LOGO.jpg");
        if (logoUrl == null) {
            return;
        }
        Image icono = new ImageIcon(logoUrl).getImage();
        setIconImage(icono);
    }//aplicarIconoVentana

    private void asegurarArchivoUsuarios() {
        File archivo = obtenerArchivoUsuarios();
        if (!archivo.exists()) {
            try (Workbook libro = new XSSFWorkbook()) {
                Sheet hoja = libro.createSheet("USUARIOS");
                asegurarEncabezados(hoja);
                guardarLibro(libro, archivo);
            } catch (IOException ex) {
                mostrarError("No se pudo crear el archivo de usuarios.", ex);
            }
            return;
        }

        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaUsuarios(libro);
            if (asegurarEncabezados(hoja)) {
                guardarLibro(libro, archivo);
            }
        } catch (IOException ex) {
            mostrarError("No se pudo preparar el archivo de usuarios.", ex);
        }
    }//asegurarArchivoUsuarios

    private void crearUsuario() {
        if (!validarCamposBase()) {
            return;
        }

        String idUsuario = jTextField1.getText().trim();

        try {
            File archivo = obtenerArchivoUsuarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaUsuarios(libro);
                asegurarEncabezados(hoja);

                if (buscarFilaUsuario(hoja, idUsuario) >= 0) {
                    JOptionPane.showMessageDialog(this, "Ya existe un usuario con ese nombre de usuario.");
                    return;
                }

                int nuevaFila = primeraFilaDisponible(hoja);
                Row fila = hoja.getRow(nuevaFila);
                if (fila == null) {
                    fila = hoja.createRow(nuevaFila);
                }
                guardarDatosFormularioEnFila(fila, true);
                guardarLibro(libro, archivo);
            }

            JOptionPane.showMessageDialog(this, "Usuario creado. Contraseña inicial: mismo valor que nombre de usuario.");
            limpiarFormulario(false);
        } catch (IOException ex) {
            mostrarError("No se pudo crear el usuario.", ex);
        }
    }//crearUsuario

    private void editarUsuario() {
        if (!validarCamposBase()) {
            return;
        }

        String idUsuario = jTextField1.getText().trim();

        try {
            File archivo = obtenerArchivoUsuarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaUsuarios(libro);
                asegurarEncabezados(hoja);

                int filaIndex = buscarFilaUsuario(hoja, idUsuario);
                if (filaIndex < 0) {
                    JOptionPane.showMessageDialog(this, "No se encontro un usuario con ese NOMBRE DE USUARIO.");
                    return;
                }

                Row fila = hoja.getRow(filaIndex);
                if (fila == null) {
                    fila = hoja.createRow(filaIndex);
                }
                guardarDatosFormularioEnFila(fila, false);
                guardarLibro(libro, archivo);
            }

            JOptionPane.showMessageDialog(this, "Usuario actualizado correctamente.");
        } catch (IOException ex) {
            mostrarError("No se pudo editar el usuario.", ex);
        }
    }//editarUsuario

    private void eliminarUsuario() {
        String idUsuario = jTextField1.getText().trim();
        if (idUsuario.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Ingresa NOMBRE DE USUARIO para eliminar.");
            return;
        }

        int confirmacion = JOptionPane.showConfirmDialog(
                this,
                "Se eliminara el usuario " + idUsuario + ". Deseas continuar?",
                "Confirmar eliminacion",
                JOptionPane.YES_NO_OPTION
        );

        if (confirmacion != JOptionPane.YES_OPTION) {
            return;
        }

        try {
            File archivo = obtenerArchivoUsuarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaUsuarios(libro);
                asegurarEncabezados(hoja);

                int filaIndex = buscarFilaUsuario(hoja, idUsuario);
                if (filaIndex < 0) {
                    JOptionPane.showMessageDialog(this, "No se encontro un usuario con ese NOMBRE DE USUARIO.");
                    return;
                }

                Row fila = hoja.getRow(filaIndex);
                if (fila != null) {
                    hoja.removeRow(fila);
                }
                int ultimaFila = hoja.getLastRowNum();
                if (filaIndex < ultimaFila) {
                    hoja.shiftRows(filaIndex + 1, ultimaFila, -1);
                }

                guardarLibro(libro, archivo);
            }

            JOptionPane.showMessageDialog(this, "Usuario eliminado correctamente.");
            limpiarFormulario(false);
        } catch (IOException ex) {
            mostrarError("No se pudo eliminar el usuario.", ex);
        }
    }//eliminarUsuario

    private void buscarUsuario() {
        String idUsuario = jTextField1.getText().trim();
        if (idUsuario.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Ingresa NOMBRE DE USUARIO para buscar.");
            return;
        }

        try {
            File archivo = obtenerArchivoUsuarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaUsuarios(libro);
                asegurarEncabezados(hoja);

                int filaIndex = buscarFilaUsuario(hoja, idUsuario);
                if (filaIndex < 0) {
                    JOptionPane.showMessageDialog(this, "No se encontro un usuario con ese NOMBRE DE USUARIO.");
                    return;
                }

                Row fila = hoja.getRow(filaIndex);
                if (fila == null) {
                    JOptionPane.showMessageDialog(this, "El registro encontrado esta vacio.");
                    return;
                }

                cargarFilaEnFormulario(fila);
                JOptionPane.showMessageDialog(this, "Usuario encontrado.");
            }
        } catch (IOException ex) {
            mostrarError("No se pudo buscar el usuario.", ex);
        }
    }//buscarUsuario

    private boolean validarCamposBase() {
        if (jTextField1.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "NOMBRE DE USUARIO es obligatorio.");
            return false;
        }
        if (jTextField2.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(this, "NOMBRE COMPLETO es obligatorio.");
            return false;
        }
        if (seleccionCombo(jComboBox1).isEmpty()) {
            JOptionPane.showMessageDialog(this, "Selecciona AREA O DEPARTAMENTO.");
            return false;
        }
        if (seleccionCombo(jComboBox2).isEmpty()) {
            JOptionPane.showMessageDialog(this, "Selecciona PUESTO.");
            return false;
        }
        return true;
    }//validarCamposBase

    private void guardarDatosFormularioEnFila(Row fila, boolean esNuevo) {
        String idUsuario = jTextField1.getText().trim();
        String contrasenaActual = leerCelda(fila, 1);

        escribirCelda(fila, 0, idUsuario);
        if (esNuevo || contrasenaActual.isEmpty()) {
            escribirCelda(fila, 1, idUsuario);
        } else {
            escribirCelda(fila, 1, contrasenaActual);
        }
        escribirCelda(fila, 2, jTextField2.getText().trim());
        escribirCelda(fila, 3, seleccionCombo(jComboBox1));
        escribirCelda(fila, 4, seleccionCombo(jComboBox2));
        escribirCelda(fila, 5, jTextField5.getText().trim());
        escribirCelda(fila, 6, jTextField6.getText().trim());
    }// guardarDatosFormularioEnFila

    private void cargarFilaEnFormulario(Row fila) {
        jTextField1.setText(leerCelda(fila, 0));
        jTextField2.setText(leerCelda(fila, 2));
        seleccionarCombo(jComboBox1, leerCelda(fila, 3));
        seleccionarCombo(jComboBox2, leerCelda(fila, 4));
        jTextField5.setText(leerCelda(fila, 5));
        jTextField6.setText(leerCelda(fila, 6));
    }//cargarFilaEnFormulario

    private void limpiarFormulario(boolean conservarId) {
        if (!conservarId) {
            jTextField1.setText("");
        }
        jTextField2.setText("");
        jComboBox1.setSelectedIndex(0);
        jComboBox2.setSelectedIndex(0);
        jTextField5.setText("");
        jTextField6.setText("");
    }//limpiarFormulario

    private int buscarFilaUsuario(Sheet hoja, String idUsuario) {
        for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
            Row row = hoja.getRow(fila);
            if (row == null) {
                continue;
            }

            String idActual = leerCelda(row, 0);
            if (idUsuario.equalsIgnoreCase(idActual)) {
                return fila;
            }
        }
        return -1;
    }//buscarFilaUsuario

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
    }//asegurarEncabezados

    private Sheet obtenerHojaUsuarios(Workbook libro) {
        if (libro.getNumberOfSheets() == 0) {
            return libro.createSheet("USUARIOS");
        }
        return libro.getSheetAt(0);
    }//obtenerHojaUsuarios

    private Workbook cargarLibro(File archivo) throws IOException {
        if (!archivo.exists()) {
            Workbook libroNuevo = new XSSFWorkbook();
            Sheet hojaNueva = libroNuevo.createSheet("USUARIOS");
            asegurarEncabezados(hojaNueva);
            guardarLibro(libroNuevo, archivo);
            return libroNuevo;
        }

        try (FileInputStream fis = new FileInputStream(archivo)) {
            return new XSSFWorkbook(fis);
        }
    }//cargarLibro

    private void guardarLibro(Workbook libro, File archivo) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(archivo)) {
            libro.write(fos);
        }
    }//guardarLibro

    private File obtenerArchivoUsuarios() {
        File archivoAbsoluto = new File(USUARIOS_FILE_ABSOLUTE);
        if (archivoAbsoluto.exists()) {
            return archivoAbsoluto;
        }

        File archivoRelativo = new File(USUARIOS_FILE_RELATIVE);
        if (archivoRelativo.exists()) {
            return archivoRelativo;
        }

        File parent = archivoAbsoluto.getParentFile();
        if (parent != null && !parent.exists()) {
            parent.mkdirs();
        }
        return archivoAbsoluto;
    }//obtenerArchivoUsuarios

    private String leerCelda(Row fila, int columna) {
        if (fila == null) {
            return "";
        }
        Cell celda = fila.getCell(columna, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (celda == null) {
            return "";
        }
        return DATA_FORMATTER.formatCellValue(celda).trim();
    }//leerCelda

    private void escribirCelda(Row fila, int columna, String valor) {
        Cell celda = fila.getCell(columna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        celda.setCellValue(valor == null ? "" : valor.trim());
    }//escribirCelda

    private String seleccionCombo(JComboBox<String> combo) {
        Object item = combo.getSelectedItem();
        if (item == null) {
            return "";
        }
        String valor = item.toString().trim();
        if (valor.isEmpty() || valor.startsWith("**")) {
            return "";
        }
        return valor;
    }//escribirCelda

    private void seleccionarCombo(JComboBox<String> combo, String valor) {
        String limpio = valor == null ? "" : valor.trim();
        if (limpio.isEmpty()) {
            combo.setSelectedIndex(0);
            return;
        }

        for (int i = 0; i < combo.getItemCount(); i++) {
            String item = combo.getItemAt(i);
            if (item.equalsIgnoreCase(limpio)) {
                combo.setSelectedIndex(i);
                return;
            }
        }//seleccioncombo

        combo.addItem(limpio);
        combo.setSelectedItem(limpio);
    }//seleccionarCombo

    private void mostrarError(String mensajeUsuario, Exception ex) {
        logger.log(java.util.logging.Level.SEVERE, mensajeUsuario, ex);
        JOptionPane.showMessageDialog(this, mensajeUsuario + "\nDetalle tecnico: " + ex.getMessage());
    }//mostrarError
}
