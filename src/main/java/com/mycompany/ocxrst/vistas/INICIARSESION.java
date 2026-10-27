
package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.IconoVentanaUtil;
import com.mycompany.ocxrst.EstiloAplicacion;
import com.mycompany.ocxrst.config.AppConfig;
import java.awt.Graphics2D;
import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.awt.image.BaseMultiResolutionImage;
import java.awt.image.BufferedImage;
import java.io.IOException;
import javax.imageio.ImageIO;
import javax.swing.ImageIcon;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author Maria Luisa Martinez
 */
public class INICIARSESION extends javax.swing.JFrame {

    // =====================================================================
    // Estado de la sesion actual
    // =====================================================================

    /** Usuario que inició sesión actualmente */
    public static String usuarioActual = "";
    /** Nombre completo asociado al usuario autenticado */
    public static String nombreUsuarioActual = "";
    
    private static final java.util.logging.Logger logger = java.util.logging.Logger.getLogger(INICIARSESION.class.getName());

    // =====================================================================
    // Construccion y recursos visuales
    // =====================================================================

    /** Inicializa el formulario de acceso y carga sus imagenes. */
    public INICIARSESION() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        initComponents();
        LOGO.setIcon(cargarImagenEscalada("/com/mycompany/ocxrst/IMAGENES/FONDO LOGIN.png", 550, 410));
        IconoVentanaUtil.aplicar(this);
        setLocationRelativeTo(null); 
        EstiloAplicacion.aplicarEstiloBotones(getContentPane());
        java.awt.EventQueue.invokeLater(() -> jTextField1.requestFocusInWindow());
    
    }

    /**
     * Carga una imagen desde el classpath y la escala exactamente al tamaño
     * indicado (sin conservar relación de aspecto), útil para fondos que
     * deben cubrir por completo un área determinada.
     */
    private static ImageIcon cargarImagenEscalada(String rutaClasspath, int ancho, int alto) {
        URL url = INICIARSESION.class.getResource(rutaClasspath);
        if (url == null) {
            return IconoVentanaUtil.obtenerIconoSeguro(rutaClasspath);
        }
        BufferedImage imagenOriginal;
        try {
            imagenOriginal = ImageIO.read(url);
        } catch (IOException ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo leer el fondo de inicio de sesión", ex);
            return IconoVentanaUtil.obtenerIconoSeguro(rutaClasspath);
        }
        if (imagenOriginal == null) {
            return IconoVentanaUtil.obtenerIconoSeguro(rutaClasspath);
        }

        BufferedImage imagenBase = redimensionarImagen(imagenOriginal, ancho, alto);
        BufferedImage imagenAltaDensidad = redimensionarImagen(imagenOriginal, ancho * 2, alto * 2);
        BufferedImage imagenMuyAltaDensidad = redimensionarImagen(imagenOriginal, ancho * 3, alto * 3);
        BaseMultiResolutionImage imagenMultirresolucion = new BaseMultiResolutionImage(
                imagenBase,
                imagenAltaDensidad,
                imagenMuyAltaDensidad);
        return new ImageIcon(imagenMultirresolucion);
    }

    private static BufferedImage redimensionarImagen(BufferedImage imagenOriginal, int ancho, int alto) {
        BufferedImage imagenEscalada = new BufferedImage(ancho, alto, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = imagenEscalada.createGraphics();
        try {
            graphics.setRenderingHint(
                    java.awt.RenderingHints.KEY_INTERPOLATION,
                    java.awt.RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            graphics.setRenderingHint(
                    java.awt.RenderingHints.KEY_RENDERING,
                    java.awt.RenderingHints.VALUE_RENDER_QUALITY);
            graphics.setRenderingHint(
                    java.awt.RenderingHints.KEY_COLOR_RENDERING,
                    java.awt.RenderingHints.VALUE_COLOR_RENDER_QUALITY);
            graphics.setRenderingHint(
                    java.awt.RenderingHints.KEY_DITHERING,
                    java.awt.RenderingHints.VALUE_DITHER_DISABLE);
            graphics.setRenderingHint(
                    java.awt.RenderingHints.KEY_ALPHA_INTERPOLATION,
                    java.awt.RenderingHints.VALUE_ALPHA_INTERPOLATION_QUALITY);
            graphics.drawImage(imagenOriginal, 0, 0, ancho, alto, null);
        } finally {
            graphics.dispose();
        }
        return imagenEscalada;
    }

    // =====================================================================
    // Componentes generados por NetBeans: no modificar manualmente
    // =====================================================================

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPasswordField1 = new javax.swing.JPasswordField();
        jTextField1 = new javax.swing.JTextField();
        jButton1 = new javax.swing.JButton();
        jLabel4 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        salir = new javax.swing.JButton();
        LOGO = new javax.swing.JLabel();
        lblBienvenido = new javax.swing.JLabel();
        lblSubtitulo = new javax.swing.JLabel();
        lblTagline = new javax.swing.JLabel();
        lblFooter = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setUndecorated(true);
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        java.awt.Color colorTitulo = new java.awt.Color(13, 42, 84);
        java.awt.Color colorSubtitulo = new java.awt.Color(110, 122, 133);
        java.awt.Color colorEtiquetaCampo = new java.awt.Color(90, 105, 120);
        java.awt.Color colorBordeCampo = new java.awt.Color(200, 210, 220);
        java.awt.Color colorBoton = new java.awt.Color(8, 84, 132);
        javax.swing.border.Border bordeCampo = javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(colorBordeCampo, 1),
                javax.swing.BorderFactory.createEmptyBorder(0, 14, 0, 14));

        lblTagline.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD | java.awt.Font.ITALIC, 10));
        lblTagline.setForeground(new java.awt.Color(205, 226, 238));
        lblTagline.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblTagline.setText("CONECTANDO PORTUNIDADES");
        getContentPane().add(lblTagline, new org.netbeans.lib.awtextra.AbsoluteConstraints(18, 132, 170, 18));

        lblBienvenido.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 26));
        lblBienvenido.setForeground(colorTitulo);
        lblBienvenido.setText("Bienvenido");
        getContentPane().add(lblBienvenido, new org.netbeans.lib.awtextra.AbsoluteConstraints(225, 42, 290, 40));

        lblSubtitulo.setFont(new java.awt.Font("Segoe UI", 0, 13));
        lblSubtitulo.setForeground(colorSubtitulo);
        lblSubtitulo.setText("Accede a tu cuenta para continuar");
        getContentPane().add(lblSubtitulo, new org.netbeans.lib.awtextra.AbsoluteConstraints(225, 85, 300, 20));

        jLabel4.setFont(new java.awt.Font("Segoe UI", 1, 11)); // NOI18N
        jLabel4.setForeground(colorEtiquetaCampo);
        jLabel4.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        jLabel4.setText("USUARIO");
        getContentPane().add(jLabel4, new org.netbeans.lib.awtextra.AbsoluteConstraints(225, 125, 200, 16));

        jTextField1.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jTextField1.setAutoscrolls(false);
        jTextField1.setBorder(bordeCampo);
        jTextField1.setBackground(new java.awt.Color(255, 255, 255));
        jTextField1.setDoubleBuffered(true);
        jTextField1.addActionListener(this::jTextField1ActionPerformed);
        getContentPane().add(jTextField1, new org.netbeans.lib.awtextra.AbsoluteConstraints(225, 143, 290, 36));

        jLabel3.setFont(new java.awt.Font("Segoe UI", 1, 11)); // NOI18N
        jLabel3.setForeground(colorEtiquetaCampo);
        jLabel3.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        jLabel3.setText("CONTRASEÑA");
        getContentPane().add(jLabel3, new org.netbeans.lib.awtextra.AbsoluteConstraints(225, 190, 200, 16));

        jPasswordField1.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jPasswordField1.setBorder(bordeCampo);
        jPasswordField1.setBackground(new java.awt.Color(255, 255, 255));
        jPasswordField1.addActionListener(this::jPasswordField1ActionPerformed);
        getContentPane().add(jPasswordField1, new org.netbeans.lib.awtextra.AbsoluteConstraints(225, 208, 290, 36));

        jButton1.setBackground(colorBoton);
        jButton1.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jButton1.setForeground(new java.awt.Color(255, 255, 255));
        jButton1.setText("INICIA SESIÓN   \u2192");
        jButton1.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 0, 0, 0));
        jButton1.setFocusPainted(false);
        jButton1.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jButton1.addActionListener(this::jButton1ActionPerformed);
        getContentPane().add(jButton1, new org.netbeans.lib.awtextra.AbsoluteConstraints(225, 262, 290, 42));

        lblFooter.setFont(new java.awt.Font("Segoe UI", 1, 9)); // NOI18N
        lblFooter.setForeground(new java.awt.Color(150, 160, 170));
        lblFooter.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        lblFooter.setText("TELECOMUNICACIONES PARA UN MEJOR FUTURO");
        getContentPane().add(lblFooter, new org.netbeans.lib.awtextra.AbsoluteConstraints(225, 322, 290, 16));

        salir.setText("\u00D7");
        salir.setFont(new java.awt.Font("Segoe UI", 1, 16)); // NOI18N
        salir.setForeground(new java.awt.Color(60, 70, 80));
        salir.setBorder(javax.swing.BorderFactory.createEmptyBorder());
        salir.setContentAreaFilled(false);
        salir.setFocusPainted(false);
        salir.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        salir.addActionListener(this::salirActionPerformed);
        getContentPane().add(salir, new org.netbeans.lib.awtextra.AbsoluteConstraints(514, 12, 24, 24));

        LOGO.setBackground(new java.awt.Color(255, 255, 255));
        LOGO.setHorizontalAlignment(javax.swing.SwingConstants.TRAILING);
        LOGO.setIcon(IconoVentanaUtil.obtenerIconoSeguro("/com/mycompany/ocxrst/IMAGENES/FONDO LOGIN.png")); // NOI18N
        LOGO.setLabelFor(LOGO);
        LOGO.setDoubleBuffered(true);
        LOGO.setFocusCycleRoot(true);
        LOGO.setFocusTraversalPolicyProvider(true);
        LOGO.setOpaque(true);
        getContentPane().add(LOGO, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 550, 410));

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jTextField1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField1ActionPerformed
        autenticar();
    }//GEN-LAST:event_jTextField1ActionPerformed

    // =====================================================================
    // Autenticacion y navegacion
    // =====================================================================

    /** Valida las credenciales y abre la pantalla principal si son correctas. */
    private void autenticar() {
        String usuario = jTextField1.getText();
        String contraseña = new String(jPasswordField1.getPassword());

        if (validarUsuarioExcel(usuario, contraseña)) {
            usuarioActual = usuario;
            nombreUsuarioActual = obtenerNombreUsuarioExcel(usuario, contraseña);
            String nombreMostrado = nombreUsuarioActual == null || nombreUsuarioActual.isBlank()
                    ? usuario
                    : nombreUsuarioActual;
            mostrarMensaje("Bienvenido " + nombreMostrado, javax.swing.JOptionPane.INFORMATION_MESSAGE);

            PRINCIPAL.reiniciarAlertaEntregasSesion();
            PRINCIPAL ventana = new PRINCIPAL();
            ventana.setVisible(true);

            this.dispose();
        } else {
            mostrarMensaje("Usuario o contraseña incorrectos", javax.swing.JOptionPane.WARNING_MESSAGE);
        }
    }

    private void mostrarMensaje(String mensaje, int tipoMensaje) {
        javax.swing.JOptionPane panel = new javax.swing.JOptionPane(
                mensaje,
                tipoMensaje,
                javax.swing.JOptionPane.DEFAULT_OPTION);
        panel.setBackground(java.awt.Color.WHITE);
        panel.setOpaque(true);
        aplicarFondoBlanco(panel);

        javax.swing.JDialog dialogo = panel.createDialog(this, "RST");
        dialogo.getContentPane().setBackground(java.awt.Color.WHITE);
        ((javax.swing.JComponent) dialogo.getContentPane()).setOpaque(true);
        aplicarFondoBlanco(dialogo.getContentPane());
        dialogo.setResizable(false);
        dialogo.setLocationRelativeTo(this);
        dialogo.setVisible(true);
    }

    private static void aplicarFondoBlanco(java.awt.Container contenedor) {
        if (contenedor instanceof javax.swing.JComponent componente) {
            componente.setBackground(java.awt.Color.WHITE);
            componente.setOpaque(true);
        }
        for (java.awt.Component componente : contenedor.getComponents()) {
            if (componente instanceof javax.swing.JComponent componenteSwing) {
                componenteSwing.setBackground(java.awt.Color.WHITE);
                componenteSwing.setOpaque(true);
            }
            if (componente instanceof java.awt.Container) {
                aplicarFondoBlanco((java.awt.Container) componente);
            }
        }
    }

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        autenticar();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void salirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_salirActionPerformed
        System.exit(0);
    }//GEN-LAST:event_salirActionPerformed

    private void jPasswordField1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jPasswordField1ActionPerformed
        autenticar();
    }//GEN-LAST:event_jPasswordField1ActionPerformed

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

        // Mantiene el mensaje blanco y la barra de botones gris claro, como las alertas del sistema.
        javax.swing.UIManager.put("OptionPane.background", java.awt.Color.WHITE);
        javax.swing.UIManager.put("OptionPane.messageAreaBackground", java.awt.Color.WHITE);
        javax.swing.UIManager.put("OptionPane.buttonAreaBackground", new java.awt.Color(238, 238, 238));
        javax.swing.UIManager.put("Panel.background", java.awt.Color.WHITE);

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> new INICIARSESION().setVisible(true));
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JLabel LOGO;
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JPasswordField jPasswordField1;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JButton salir;
    private javax.swing.JLabel lblBienvenido;
    private javax.swing.JLabel lblSubtitulo;
    private javax.swing.JLabel lblTagline;
    private javax.swing.JLabel lblFooter;
    // End of variables declaration//GEN-END:variables



    // =====================================================================
    // Lectura de credenciales desde Excel
    // =====================================================================

    public static String obtenerNombreUsuarioExcel(String usuario, String contraseña) {
        try {
            File archivo = AppConfig.resolveBaseFile("USUARIOS.xlsx");
            if (!archivo.exists()) {
                return "";
            }

            try (FileInputStream fis = new FileInputStream(archivo);
                 Workbook libro = new XSSFWorkbook(fis)) {
                return obtenerNombreUsuarioExcel(libro, usuario, contraseña);
            }
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }

    public static String obtenerNombreUsuarioExcel(Workbook libro, String usuario, String contraseña) {
        if (libro == null) {
            return "";
        }

        Sheet hoja = libro.getSheetAt(0);
        for (int i = 1; i <= hoja.getLastRowNum(); i++) {
            Row fila = hoja.getRow(i);
            if (fila == null) {
                continue;
            }

            String userExcel = obtenerValorCelda(fila, 0);
            String passExcel = obtenerValorCelda(fila, 1);
            String nombreExcel = obtenerValorCelda(fila, 2);

            if (usuario.equals(userExcel) && contraseña.equals(passExcel)) {
                return nombreExcel;
            }
        }

        return "";
    }

    private static String obtenerValorCelda(Row fila, int columna) {
        if (fila == null) {
            return "";
        }
        Cell celda = fila.getCell(columna, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (celda == null) {
            return "";
        }
        return celda.toString().trim();
    }

    public boolean validarUsuarioExcel(String usuario, String contraseña) {
        boolean encontrado = false;

        try {
            File archivo = AppConfig.resolveBaseFile("USUARIOS.xlsx");
            FileInputStream fis = new FileInputStream(archivo);
            Workbook libro = new XSSFWorkbook(fis);
            Sheet hoja = libro.getSheetAt(0);

            for (int i = 1; i <= hoja.getLastRowNum(); i++) {
                Row fila = hoja.getRow(i);

                if (fila != null) {
                    Cell celdaUsuario = fila.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    Cell celdaPass = fila.getCell(1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                    String userExcel = celdaUsuario == null ? "" : celdaUsuario.toString();
                    String passExcel = celdaPass == null ? "" : celdaPass.toString();

                    if (usuario.equals(userExcel) && contraseña.equals(passExcel)) {
                        encontrado = true;
                        break;
                    }
                }
            }

            libro.close();
            fis.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

        return encontrado;
    }
}
