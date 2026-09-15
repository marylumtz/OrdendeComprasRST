
package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.IconoVentanaUtil;

public class PRINCIPAL extends javax.swing.JFrame {

    private static boolean alertaEntregasMostradaEnSesion = false;
    private final java.util.Map<String, javax.swing.JLabel> etiquetasValorIndicadores = new java.util.HashMap<>();

    public PRINCIPAL() {
        initComponents();
        setSize(800, 550);
        setResizable(false);
        configurarMenuSesion();
        actualizarBotonNotificaciones();
        iniciarReloj();
        IconoVentanaUtil.aplicar(this);
        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        setLocationRelativeTo(null); 
        mostrarAlertaEntregasUnaVezPorSesion();
    }

    private void configurarMenuSesion() {
        menuSesion = new javax.swing.JPopupMenu();
        javax.swing.JMenuItem cerrarSesion = new javax.swing.JMenuItem("CERRAR SESION");
        cerrarSesion.addActionListener(e -> cerrarSesion());
        menuSesion.add(cerrarSesion);
        javax.swing.JMenuItem salirPrograma = new javax.swing.JMenuItem("SALIR DEL PROGRAMA");
        salirPrograma.addActionListener(e -> salirDelPrograma());
        menuSesion.add(salirPrograma);

        String nombreUsuario = INICIARSESION.nombreUsuarioActual == null
            || INICIARSESION.nombreUsuarioActual.isBlank()
            ? INICIARSESION.usuarioActual
            : INICIARSESION.nombreUsuarioActual;
        jButtonMenu = new javax.swing.JButton(nombreUsuario);
        jButtonMenu.setBackground(java.awt.Color.WHITE);
        jButtonMenu.setForeground(java.awt.Color.BLACK);
        jButtonMenu.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        jButtonMenu.setFocusPainted(false);
        jButtonMenu.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 101, 151)));
        int anchoBoton = Math.max(110, jButtonMenu.getPreferredSize().width + 12);
        jButtonMenu.setPreferredSize(new java.awt.Dimension(anchoBoton, 34));
        jButtonMenu.addActionListener(e -> menuSesion.show(jButtonMenu, 0, jButtonMenu.getHeight()));
        panelUsuario.add(jButtonMenu);
        java.awt.Dimension tamanoMenu = panelUsuario.getPreferredSize();
        panelUsuario.setPreferredSize(tamanoMenu);
        panelEspaciador.setPreferredSize(tamanoMenu);
        panelUsuario.revalidate();
        panelEspaciador.revalidate();
    }

    private void cerrarSesion() {
        INICIARSESION.usuarioActual = "";
        INICIARSESION.nombreUsuarioActual = "";
        reiniciarAlertaEntregasSesion();
        INICIARSESION inicioSesion = new INICIARSESION();
        inicioSesion.setVisible(true);
        dispose();
    }

    private void salirDelPrograma() {
        dispose();
        System.exit(0);
    }

    public static void reiniciarAlertaEntregasSesion() {
        alertaEntregasMostradaEnSesion = false;
    }

    public static void actualizarAlarmasEnTiempoReal() {
        Runnable actualizar = () -> {
            for (java.awt.Window ventana : java.awt.Window.getWindows()) {
                if (ventana instanceof PRINCIPAL principal && principal.isDisplayable()) {
                    principal.actualizarBotonNotificaciones();
                }
            }
        };
        if (java.awt.EventQueue.isDispatchThread()) {
            actualizar.run();
        } else {
            java.awt.EventQueue.invokeLater(actualizar);
        }
    }

    private void mostrarAlertaEntregasUnaVezPorSesion() {
        if (alertaEntregasMostradaEnSesion) {
            return;
        }
        alertaEntregasMostradaEnSesion = true;
        java.awt.EventQueue.invokeLater(() -> PROVEEDORES.mostrarAlertaEntregasPendientes(this));
    }

    private void actualizarBotonNotificaciones() {
        int notificaciones = PROVEEDORES.contarAlertasEntregasPendientes();
        jButtonNotificaciones.setText("NOTIFICACIONES (" + notificaciones + ")");
        jButtonNotificaciones.setBackground(notificaciones > 0
                ? new java.awt.Color(180, 0, 0)
                : new java.awt.Color(0, 95, 131));
        jButtonNotificaciones.setForeground(java.awt.Color.WHITE);
    }

    private void initComponents() {
        jPanel1 = crearPanelConFondo();
        jLabel1 = new javax.swing.JLabel();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jButton3 = new javax.swing.JButton();
        jButton4 = new javax.swing.JButton();
        jButton5 = new javax.swing.JButton();
        jButtonNotificaciones = new javax.swing.JButton();
        jButton6 = new javax.swing.JButton();
        jLabel3 = new javax.swing.JLabel();
        panelUsuario = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 0, 0));
        panelEspaciador = new javax.swing.JPanel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        jPanel1.setLayout(new java.awt.BorderLayout());

        java.awt.Color azul = new java.awt.Color(0, 101, 151);
        javax.swing.JPanel encabezado = new javax.swing.JPanel(new java.awt.BorderLayout(16, 0));
        encabezado.setOpaque(false);
        encabezado.setBorder(javax.swing.BorderFactory.createEmptyBorder(8, 24, 8, 24));
        panelEspaciador.setOpaque(false);
        encabezado.add(panelEspaciador, java.awt.BorderLayout.WEST);

        jButtonNotificaciones.setBackground(java.awt.Color.WHITE);
        jButtonNotificaciones.setForeground(azul);
        jButtonNotificaciones.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 12));
        jButtonNotificaciones.setBorder(javax.swing.BorderFactory.createEmptyBorder(6, 10, 6, 10));
        jButtonNotificaciones.setFocusPainted(false);
        jButtonNotificaciones.setContentAreaFilled(false);
        jButtonNotificaciones.setOpaque(false);
        jButtonNotificaciones.addActionListener(this::jButtonNotificacionesActionPerformed);
        encabezado.add(jButtonNotificaciones, java.awt.BorderLayout.CENTER);

        panelUsuario.setOpaque(false);
        encabezado.add(panelUsuario, java.awt.BorderLayout.EAST);
        jPanel1.add(encabezado, java.awt.BorderLayout.NORTH);

        javax.swing.JPanel contenido = new javax.swing.JPanel();
        contenido.setOpaque(false);
        contenido.setBorder(javax.swing.BorderFactory.createEmptyBorder(18, 120, 14, 120));
        contenido.setLayout(new javax.swing.BoxLayout(contenido, javax.swing.BoxLayout.Y_AXIS));

        javax.swing.JPanel tituloSistema = new javax.swing.JPanel();
        tituloSistema.setOpaque(false);
        tituloSistema.setLayout(new javax.swing.BoxLayout(tituloSistema, javax.swing.BoxLayout.Y_AXIS));
        tituloSistema.setAlignmentX(java.awt.Component.CENTER_ALIGNMENT);

        jLabel1.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 14));
        jLabel1.setForeground(new java.awt.Color(82, 111, 133));
        jLabel1.setAlignmentX(java.awt.Component.CENTER_ALIGNMENT);
        jLabel1.setText("SISTEMA DE");
        tituloSistema.add(jLabel1);

        javax.swing.JLabel tituloOrdenes = new javax.swing.JLabel("ÓRDENES DE COMPRAS");
        tituloOrdenes.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 30));
        tituloOrdenes.setForeground(new java.awt.Color(0, 82, 135));
        tituloOrdenes.setAlignmentX(java.awt.Component.CENTER_ALIGNMENT);
        tituloSistema.add(tituloOrdenes);

        javax.swing.JSeparator separadorTitulo = new javax.swing.JSeparator();
        separadorTitulo.setForeground(new java.awt.Color(0, 150, 185));
        separadorTitulo.setMaximumSize(new java.awt.Dimension(60, 2));
        separadorTitulo.setPreferredSize(new java.awt.Dimension(60, 2));
        separadorTitulo.setAlignmentX(java.awt.Component.CENTER_ALIGNMENT);
        tituloSistema.add(javax.swing.Box.createVerticalStrut(5));
        tituloSistema.add(separadorTitulo);
        contenido.add(tituloSistema);

        javax.swing.JLabel subtitulo = new javax.swing.JLabel("Gestiona compras, proveedores e inventarios");
        subtitulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 14));
        subtitulo.setForeground(new java.awt.Color(82, 111, 133));
        subtitulo.setAlignmentX(java.awt.Component.CENTER_ALIGNMENT);
        contenido.add(javax.swing.Box.createVerticalStrut(5));
        contenido.add(subtitulo);
        contenido.add(javax.swing.Box.createVerticalStrut(18));

        javax.swing.JPanel tarjetas = new javax.swing.JPanel(new java.awt.GridLayout(2, 2, 12, 12));
        tarjetas.setOpaque(false);
        tarjetas.setMaximumSize(new java.awt.Dimension(560, 178));

        jButton1.setText("ORDEN DE COMPRAS");
        jButton1.addActionListener(this::jButton1ActionPerformed);
        jButton2.setText("PROVEEDORES");
        jButton2.addActionListener(this::jButton2ActionPerformed);
        jButton3.setText("USUARIOS");
        jButton3.addActionListener(this::jButton3ActionPerformed);
        jButton4.setText("INVENTARIOS");
        jButton4.addActionListener(this::jButton4ActionPerformed);
        jButton5.setText("TERMINOS Y CONDICIONES");
        jButton5.addActionListener(this::jButton5ActionPerformed);
        configurarTarjeta(jButton1, new java.awt.Color(191, 229, 248));
        configurarTarjeta(jButton2, new java.awt.Color(190, 238, 226));
        configurarTarjeta(jButton3, new java.awt.Color(218, 215, 247));
        configurarTarjeta(jButton4, new java.awt.Color(247, 224, 180));
        tarjetas.add(jButton1);
        tarjetas.add(jButton2);
        tarjetas.add(jButton3);
        tarjetas.add(jButton4);
        tarjetas.setAlignmentX(java.awt.Component.CENTER_ALIGNMENT);
        contenido.add(tarjetas);
        contenido.add(javax.swing.Box.createVerticalStrut(12));

        configurarTarjeta(jButton5, new java.awt.Color(246, 207, 218));
        jButton5.setMaximumSize(new java.awt.Dimension(560, 58));
        jButton5.setAlignmentX(java.awt.Component.CENTER_ALIGNMENT);
        contenido.add(jButton5);
        contenido.add(javax.swing.Box.createVerticalStrut(12));

        javax.swing.JPanel indicadores = new javax.swing.JPanel(new java.awt.GridLayout(1, 3));
        indicadores.setBackground(java.awt.Color.WHITE);
        indicadores.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(215, 227, 235)));
        indicadores.setMaximumSize(new java.awt.Dimension(560, 62));
        indicadores.add(crearIndicador("ORDENES", "Gestion de compras"));
        indicadores.add(crearIndicador("PENDIENTES", "Entregas por revisar"));
        indicadores.add(crearIndicador("PROVEEDORES", "Registros activos"));
        indicadores.setAlignmentX(java.awt.Component.CENTER_ALIGNMENT);
        contenido.add(indicadores);
        jPanel1.add(contenido, java.awt.BorderLayout.CENTER);
        actualizarIndicadores();

        etiquetaFechaHora = new javax.swing.JLabel();
        etiquetaFechaHora.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        etiquetaFechaHora.setForeground(new java.awt.Color(55, 103, 139));
        etiquetaFechaHora.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        etiquetaFechaHora.setBorder(javax.swing.BorderFactory.createEmptyBorder(7, 20, 7, 20));
        javax.swing.JPanel barraInferior = new javax.swing.JPanel(new java.awt.BorderLayout());
        barraInferior.setBackground(new java.awt.Color(241, 248, 252));
        barraInferior.setBorder(javax.swing.BorderFactory.createMatteBorder(1, 0, 0, 0, new java.awt.Color(193, 218, 235)));
        barraInferior.add(etiquetaFechaHora, java.awt.BorderLayout.EAST);
        jPanel1.add(barraInferior, java.awt.BorderLayout.SOUTH);

        getContentPane().setLayout(new java.awt.BorderLayout());
        getContentPane().add(jPanel1, java.awt.BorderLayout.CENTER);
    }

    private void iniciarReloj() {
        actualizarFechaHora();
        reloj = new javax.swing.Timer(1000, evento -> actualizarFechaHora());
        reloj.start();
    }

    private void actualizarFechaHora() {
        java.time.format.DateTimeFormatter formato = java.time.format.DateTimeFormatter
                .ofPattern("EEEE, d 'de' MMMM 'de' uuuu | HH:mm", new java.util.Locale("es", "MX"));
        etiquetaFechaHora.setText(java.time.LocalDateTime.now().format(formato));
    }

    private javax.swing.JPanel crearPanelConFondo() {
        java.awt.Image fondo = IconoVentanaUtil.obtenerIconoSeguro(
                "/com/mycompany/ocxrst/IMAGENES/FONDO PRINCIPAL.png").getImage();
        return new javax.swing.JPanel() {
            @Override
            protected void paintComponent(java.awt.Graphics graphics) {
                super.paintComponent(graphics);
                if (fondo != null) {
                    graphics.drawImage(fondo, 0, 0, getWidth(), getHeight(), this);
                }
            }
        };
    }

    private void configurarTarjeta(javax.swing.JButton boton, java.awt.Color color) {
        boton.setBackground(color);
        boton.setForeground(new java.awt.Color(0, 82, 135));
        boton.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 14));
        boton.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        boton.setContentAreaFilled(true);
        boton.setOpaque(true);
        boton.setBorder(javax.swing.BorderFactory.createCompoundBorder(
            javax.swing.BorderFactory.createLineBorder(color.darker()),
                javax.swing.BorderFactory.createEmptyBorder(8, 18, 8, 18)));
        boton.setFocusPainted(false);
    }

    private javax.swing.JPanel crearIndicador(String titulo, String descripcion) {
        javax.swing.JPanel indicador = new javax.swing.JPanel(new java.awt.GridLayout(3, 1));
        indicador.setOpaque(false);
        javax.swing.JLabel etiquetaValor = new javax.swing.JLabel("0", javax.swing.SwingConstants.CENTER);
        etiquetaValor.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 18));
        etiquetaValor.setForeground(new java.awt.Color(0, 101, 151));
        javax.swing.JLabel etiquetaTitulo = new javax.swing.JLabel(titulo, javax.swing.SwingConstants.CENTER);
        etiquetaTitulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        etiquetaTitulo.setForeground(new java.awt.Color(0, 82, 135));
        javax.swing.JLabel etiquetaDescripcion = new javax.swing.JLabel(descripcion, javax.swing.SwingConstants.CENTER);
        etiquetaDescripcion.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 11));
        etiquetaDescripcion.setForeground(new java.awt.Color(82, 111, 133));
        indicador.add(etiquetaValor);
        indicador.add(etiquetaTitulo);
        indicador.add(etiquetaDescripcion);
        etiquetasValorIndicadores.put(titulo, etiquetaValor);
        return indicador;
    }

    private void actualizarIndicadores() {
        int ordenes = PROVEEDORES.contarOrdenesRegistradas();
        int pendientes = PROVEEDORES.contarEntregasPorRevisar();
        int proveedores = PROVEEDORES.contarProveedoresActivos();
        actualizarValorIndicador("ORDENES", ordenes);
        actualizarValorIndicador("PENDIENTES", pendientes);
        actualizarValorIndicador("PROVEEDORES", proveedores);
    }

    private void actualizarValorIndicador(String titulo, int valor) {
        javax.swing.JLabel etiqueta = etiquetasValorIndicadores.get(titulo);
        if (etiqueta != null) {
            etiqueta.setText(String.valueOf(valor));
        }
    }

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        PROVEEDORES ventana = new PROVEEDORES();
        ventana.setLocationRelativeTo(this);
        ventana.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        USUARIOS ventana = new USUARIOS();
        ventana.setLocationRelativeTo(this);
        ventana.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_jButton3ActionPerformed

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        INVENTARIOS ventana =  new INVENTARIOS();
        ventana.setVisible(true);
        
        this.dispose();
    }//GEN-LAST:event_jButton4ActionPerformed

    private void jButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton5ActionPerformed
        TERMINOS ventana =  new TERMINOS();
        ventana.setLocationRelativeTo(this);
        ventana.setVisible(true);
        this.dispose();
    }//GEN-LAST:event_jButton5ActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        ORDENCOMPRA ventana =  new ORDENCOMPRA ();
        ventana.setVisible(true);
        
        this.dispose();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButtonNotificacionesActionPerformed(java.awt.event.ActionEvent evt) {
        actualizarBotonNotificaciones();
        PROVEEDORES.mostrarAlertaEntregasPendientes(this);
    }//jButtonNotificacionesActionPerformed

    public static void main(String args[]) {

        java.awt.EventQueue.invokeLater(() -> new PRINCIPAL().setVisible(true));
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton5;
    private javax.swing.JButton jButton6;
    private javax.swing.JButton jButtonMenu;
    private javax.swing.JButton jButtonNotificaciones;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel etiquetaFechaHora;
    private javax.swing.JPopupMenu menuSesion;
    private javax.swing.JPanel panelEspaciador;
    private javax.swing.JPanel panelUsuario;
    private javax.swing.Timer reloj;
    private javax.swing.JPanel jPanel1;
    // End of variables declaration//GEN-END:variables
}
