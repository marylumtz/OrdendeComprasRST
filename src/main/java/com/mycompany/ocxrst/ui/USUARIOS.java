/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.util.IconoVentanaUtil;
import com.mycompany.ocxrst.util.EstiloAplicacion;
import com.mycompany.ocxrst.repository.UsuarioRepository;
import com.mycompany.ocxrst.validation.UsuarioValidator;
import java.io.File;
import java.io.IOException;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author Maria Luisa Martinez
 */
public class USUARIOS extends javax.swing.JFrame {

    // =====================================================================
    // Repositorio y configuracion de usuarios
    // =====================================================================

    private static final java.util.logging.Logger logger = java.util.logging.Logger.getLogger(USUARIOS.class.getName());
    private final UsuarioRepository usuarioRepository = new UsuarioRepository();
    private java.awt.CardLayout configuracionVistasLayout;
    private javax.swing.JPanel configuracionVistas;
    private javax.swing.JButton navegacionDatos;
    private javax.swing.JButton navegacionPermisos;

    /** Inicializa el formulario y prepara el archivo de usuarios. */
    public USUARIOS() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        initComponents();
        configurarDisenoVisual();
        IconoVentanaUtil.aplicar(this);
        setSize(1050, 650);
        setResizable(false);
        setLocationRelativeTo(null);
        asegurarArchivoUsuarios();
        EstiloAplicacion.aplicarEstiloBotonesRestoPantallas(jPanel1);
        EstiloAplicacion.aplicarTipografia(jPanel1);
        EstiloAplicacion.aplicarEstiloCamposTexto(jPanel1);
    }

    private static final java.awt.Color COLOR_AZUL_OSCURO = new java.awt.Color(28, 82, 132);
    private static final java.awt.Color COLOR_AZUL = new java.awt.Color(31, 112, 184);
    private static final java.awt.Color COLOR_FONDO = EstiloAplicacion.FONDO_PANTALLA;
    private static final java.awt.Color COLOR_BORDE = new java.awt.Color(198, 210, 225);

    private void configurarDisenoVisual() {
        jPanel1.removeAll();
        jPanel1.setOpaque(false);
        jPanel1.setBackground(COLOR_FONDO);
        jPanel1.setLayout(new java.awt.BorderLayout(0, 14));

        javax.swing.JPanel contenido = new javax.swing.JPanel(new java.awt.BorderLayout(12, 12));
        contenido.setOpaque(false);
        contenido.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 8, 8, 8));

        javax.swing.JPanel acciones = new javax.swing.JPanel(new java.awt.FlowLayout(
                java.awt.FlowLayout.RIGHT, 10, 8));
        acciones.setOpaque(false);
        estilizarBoton(jButton1, "NUEVO USUARIO", COLOR_AZUL, java.awt.Color.WHITE, 165);
        jButton1.setIcon(iconoMas(java.awt.Color.WHITE));
        estilizarBoton(jButton2, "GUARDAR", java.awt.Color.WHITE, COLOR_AZUL, 130);
        jButton2.setIcon(iconoGuardar(COLOR_AZUL));
        estilizarBoton(jButton3, "ELIMINAR", java.awt.Color.WHITE, COLOR_AZUL, 135);
        jButton3.setIcon(iconoEliminar(COLOR_AZUL));
        estilizarBoton(jButton4, "ACTUALIZAR", java.awt.Color.WHITE, COLOR_AZUL, 145);
        jButton4.setIcon(iconoActualizar(COLOR_AZUL));
        acciones.add(jButton1);
        acciones.add(jButton2);
        acciones.add(jButton3);
        javax.swing.JSeparator separador = new javax.swing.JSeparator(javax.swing.SwingConstants.VERTICAL);
        separador.setPreferredSize(new java.awt.Dimension(1, 28));
        acciones.add(separador);
        acciones.add(jButton4);
        contenido.add(acciones, java.awt.BorderLayout.NORTH);

        javax.swing.JPanel cuerpo = new javax.swing.JPanel(new java.awt.BorderLayout(12, 0));
        cuerpo.setOpaque(false);
        cuerpo.add(crearNavegacion(), java.awt.BorderLayout.WEST);

        TarjetaPanel tarjeta = new TarjetaPanel();
        tarjeta.setBorder(javax.swing.BorderFactory.createEmptyBorder(14, 20, 16, 20));
        tarjeta.setLayout(new java.awt.BorderLayout(0, 12));
        javax.swing.JPanel barra = new javax.swing.JPanel(new java.awt.BorderLayout(8, 0));
        barra.setBackground(new java.awt.Color(232, 242, 255));
        barra.setBorder(javax.swing.BorderFactory.createEmptyBorder(9, 12, 9, 12));
        javax.swing.JLabel registro = new javax.swing.JLabel("Configuración");
        registro.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 18));
        registro.setForeground(COLOR_AZUL_OSCURO);
        barra.add(registro, java.awt.BorderLayout.WEST);
        javax.swing.JLabel aviso = new javax.swing.JLabel("Complete los datos del usuario para continuar");
        aviso.setForeground(COLOR_AZUL_OSCURO);
        barra.add(aviso, java.awt.BorderLayout.EAST);
        tarjeta.add(barra, java.awt.BorderLayout.NORTH);

        javax.swing.JPanel formulario = new javax.swing.JPanel(new java.awt.GridBagLayout());
        formulario.setOpaque(false);
        agregarCampo(formulario, 0, jLabel6, jTextField1, iconoId(COLOR_AZUL_OSCURO));
        agregarCampo(formulario, 1, jLabel2, jTextField2, iconoContacto(COLOR_AZUL_OSCURO));
        agregarCampo(formulario, 2, jLabel3, jTextField3, iconoEdificio(COLOR_AZUL_OSCURO, 18));
        agregarCampo(formulario, 3, jLabel4, jTextField4, iconoDocumento(COLOR_AZUL_OSCURO));
        agregarCampo(formulario, 4, jLabel5, jTextField5, iconoCorreo(COLOR_AZUL_OSCURO));
        agregarCampo(formulario, 5, jLabel7, jTextField6, iconoDocumento(COLOR_AZUL_OSCURO));
        tarjeta.add(formulario, java.awt.BorderLayout.CENTER);
        configuracionVistasLayout = new java.awt.CardLayout();
        configuracionVistas = new javax.swing.JPanel(configuracionVistasLayout);
        configuracionVistas.setOpaque(false);
        configuracionVistas.add(tarjeta, "DATOS");
        configuracionVistas.add(crearPanelPermisos(), "PERMISOS");
        cuerpo.add(configuracionVistas, java.awt.BorderLayout.CENTER);
        contenido.add(cuerpo, java.awt.BorderLayout.CENTER);
        jPanel1.add(contenido, java.awt.BorderLayout.CENTER);
        getContentPane().setBackground(COLOR_FONDO);
        getContentPane().removeAll();
        getContentPane().setLayout(new java.awt.BorderLayout());
        getContentPane().add(jPanel1, java.awt.BorderLayout.CENTER);
        getContentPane().revalidate();
        getContentPane().repaint();
    }

    private javax.swing.JPanel crearNavegacion() {
        javax.swing.JPanel navegacion = new javax.swing.JPanel();
        navegacion.setPreferredSize(new java.awt.Dimension(98, 0));
        navegacion.setBackground(new java.awt.Color(239, 247, 255));
        navegacion.setBorder(javax.swing.BorderFactory.createEmptyBorder(8, 6, 8, 6));
        navegacion.setLayout(new java.awt.GridLayout(4, 1, 0, 8));
        navegacionDatos = crearElementoNavegacion("Datos", iconoContacto(java.awt.Color.WHITE), true);
        navegacionPermisos = crearElementoNavegacion("Permisos", iconoDocumento(COLOR_AZUL), false);
        navegacionDatos.addActionListener(e -> mostrarVistaConfiguracion("DATOS"));
        navegacionPermisos.addActionListener(e -> mostrarVistaConfiguracion("PERMISOS"));
        navegacion.add(navegacionDatos);
        navegacion.add(navegacionPermisos);
        navegacion.add(crearElementoNavegacion("Accesos", iconoId(COLOR_AZUL), false));
        navegacion.add(crearElementoNavegacion("Historial", iconoDocumento(COLOR_AZUL), false));
        return navegacion;
    }

    private javax.swing.JButton crearElementoNavegacion(String texto, javax.swing.Icon icono, boolean activo) {
        javax.swing.JButton boton = new javax.swing.JButton("<html><center>" + texto + "</center></html>", icono);
        boton.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        boton.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        boton.setIconTextGap(5);
        boton.setBackground(activo ? COLOR_AZUL : java.awt.Color.WHITE);
        boton.setForeground(activo ? java.awt.Color.WHITE : COLOR_AZUL_OSCURO);
        boton.setBorder(javax.swing.BorderFactory.createEmptyBorder(7, 3, 7, 3));
        boton.setFocusPainted(false);
        return boton;
    }

    private void mostrarVistaConfiguracion(String vista) {
        configuracionVistasLayout.show(configuracionVistas, vista);
        boolean permisos = "PERMISOS".equals(vista);
        actualizarNavegacion(navegacionDatos, !permisos, iconoContacto(!permisos
                ? java.awt.Color.WHITE : COLOR_AZUL));
        actualizarNavegacion(navegacionPermisos, permisos, iconoDocumento(permisos
                ? java.awt.Color.WHITE : COLOR_AZUL));
    }

    private void actualizarNavegacion(javax.swing.JButton boton, boolean activo, javax.swing.Icon icono) {
        boton.setBackground(activo ? COLOR_AZUL : java.awt.Color.WHITE);
        boton.setForeground(activo ? java.awt.Color.WHITE : COLOR_AZUL_OSCURO);
        boton.setIcon(icono);
    }

    private javax.swing.JPanel crearPanelPermisos() {
        javax.swing.JPanel panel = new javax.swing.JPanel(new java.awt.BorderLayout(12, 12));
        panel.setOpaque(false);
        panel.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 0, 0, 0));

        javax.swing.JPanel encabezado = new javax.swing.JPanel(new java.awt.BorderLayout());
        encabezado.setBackground(new java.awt.Color(232, 242, 255));
        encabezado.setBorder(javax.swing.BorderFactory.createEmptyBorder(9, 12, 9, 12));
        javax.swing.JLabel titulo = new javax.swing.JLabel("Permisos del Usuario");
        titulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 18));
        titulo.setForeground(COLOR_AZUL_OSCURO);
        encabezado.add(titulo, java.awt.BorderLayout.WEST);
        javax.swing.JLabel ayuda = new javax.swing.JLabel("Seleccione los módulos y acciones que podrá utilizar el usuario");
        ayuda.setForeground(COLOR_AZUL_OSCURO);
        encabezado.add(ayuda, java.awt.BorderLayout.EAST);
        panel.add(encabezado, java.awt.BorderLayout.NORTH);

        javax.swing.JPanel contenido = new javax.swing.JPanel(new java.awt.BorderLayout(12, 0));
        contenido.setOpaque(false);
        javax.swing.JPanel tabla = crearTablaPermisos();
        contenido.add(tabla, java.awt.BorderLayout.CENTER);

        javax.swing.JPanel descripcion = new javax.swing.JPanel(new java.awt.BorderLayout(0, 14));
        descripcion.setBackground(new java.awt.Color(239, 247, 255));
        descripcion.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(new java.awt.Color(198, 218, 243)),
                javax.swing.BorderFactory.createEmptyBorder(12, 12, 12, 12)));
        javax.swing.JLabel tituloDescripcion = new javax.swing.JLabel("ⓘ  Descripción");
        tituloDescripcion.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 14));
        tituloDescripcion.setForeground(COLOR_AZUL_OSCURO);
        javax.swing.JLabel texto = new javax.swing.JLabel(
                "<html><div style='text-align:center'>Otorga los permisos<br>necesarios para que el usuario<br>pueda acceder y utilizar las<br>funciones del sistema.</div></html>",
                javax.swing.SwingConstants.CENTER);
        texto.setForeground(COLOR_AZUL_OSCURO);
        descripcion.add(tituloDescripcion, java.awt.BorderLayout.NORTH);
        descripcion.add(texto, java.awt.BorderLayout.CENTER);
        descripcion.setPreferredSize(new java.awt.Dimension(205, 0));
        contenido.add(descripcion, java.awt.BorderLayout.EAST);
        panel.add(contenido, java.awt.BorderLayout.CENTER);
        return panel;
    }

    private javax.swing.JPanel crearTablaPermisos() {
        String[] columnas = {"Módulo / Opción", "Consultar", "Crear", "Editar", "Eliminar"};
        String[][] filas = {
            {"Sistema", "", "", "", ""},
            {"    Usuarios", "✓", "✓", "✓", "✓"},
            {"    Permisos", "✓", "✓", "✓", "✓"},
            {"    Configuración", "", "", "", ""},
            {"Inventarios", "", "", "", ""},
            {"    Productos", "✓", "✓", "✓", "✓"},
            {"    Entradas", "✓", "✓", "✓", "✓"},
            {"    Salidas", "✓", "✓", "✓", "✓"},
            {"    Reportes", "✓", "✓", "✓", "✓"},
            {"Compras", "", "", "", ""},
            {"    Proveedores", "✓", "✓", "✓", "✓"},
            {"    Órdenes de Compra", "✓", "✓", "✓", "✓"},
            {"    Reportes", "", "", "", ""},
            {"Ventas", "", "", "", ""},
            {"    Clientes", "✓", "✓", "✓", "✓"},
            {"    Facturación", "✓", "✓", "✓", "✓"},
            {"    Reportes", "", "", "", ""},
            {"Reportes Generales", "✓", "✓", "✓", "✓"}
        };
        javax.swing.JPanel tabla = new javax.swing.JPanel(new java.awt.GridLayout(filas.length + 1, columnas.length));
        tabla.setBackground(java.awt.Color.WHITE);
        tabla.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(190, 210, 235)));
        for (String columna : columnas) {
            javax.swing.JLabel etiqueta = new javax.swing.JLabel(columna, javax.swing.SwingConstants.CENTER);
            etiqueta.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 12));
            etiqueta.setForeground(COLOR_AZUL_OSCURO);
            etiqueta.setBackground(new java.awt.Color(224, 236, 252));
            etiqueta.setOpaque(true);
            etiqueta.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(205, 220, 239)));
            tabla.add(etiqueta);
        }
        for (String[] fila : filas) {
            boolean modulo = !fila[0].startsWith("    ");
            for (int indice = 0; indice < fila.length; indice++) {
                if (indice == 0) {
                    javax.swing.JLabel etiqueta = new javax.swing.JLabel(fila[indice]);
                    etiqueta.setFont(new java.awt.Font("Segoe UI", modulo ? java.awt.Font.BOLD : java.awt.Font.PLAIN, 12));
                    etiqueta.setForeground(COLOR_AZUL_OSCURO);
                    etiqueta.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 8, 0, 2));
                    tabla.add(etiqueta);
                } else if (modulo) {
                    javax.swing.JLabel espacio = new javax.swing.JLabel();
                    espacio.setBackground(java.awt.Color.WHITE);
                    espacio.setOpaque(true);
                    tabla.add(espacio);
                } else {
                    javax.swing.JCheckBox casilla = new javax.swing.JCheckBox();
                    casilla.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
                    casilla.setSelected("✓".equals(fila[indice]));
                    casilla.setBackground(java.awt.Color.WHITE);
                    tabla.add(casilla);
                }
            }
        }
        return tabla;
    }

    private void agregarCampo(javax.swing.JPanel formulario, int fila,
            javax.swing.JLabel etiqueta, javax.swing.JTextField campo, javax.swing.Icon icono) {
        etiqueta.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        etiqueta.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        etiqueta.setForeground(COLOR_AZUL_OSCURO);
        etiqueta.setText("<html>" + etiqueta.getText() + " <font color='#d92323'>*</font></html>");
        campo.setBorder(javax.swing.BorderFactory.createLineBorder(COLOR_BORDE));
        campo.setPreferredSize(new java.awt.Dimension(400, 32));
        java.awt.GridBagConstraints gc = new java.awt.GridBagConstraints();
        gc.gridy = fila;
        gc.gridx = 0;
        gc.weightx = 0;
        gc.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gc.insets = new java.awt.Insets(4, 0, 4, 10);
        formulario.add(new javax.swing.JLabel(icono), gc);
        gc.gridx = 1;
        gc.insets = new java.awt.Insets(4, 0, 4, 14);
        formulario.add(etiqueta, gc);
        gc.gridx = 2;
        gc.weightx = 1;
        gc.insets = new java.awt.Insets(4, 0, 4, 0);
        formulario.add(campo, gc);
    }

    @FunctionalInterface
    private interface DibujoIcono {
        void dibujar(java.awt.Graphics2D graphics, int tamano);
    }

    private static final class VectorIcon implements javax.swing.Icon {
        private final int tamano;
        private final java.awt.Color color;
        private final DibujoIcono dibujo;

        private VectorIcon(int tamano, java.awt.Color color, DibujoIcono dibujo) {
            this.tamano = tamano;
            this.color = color;
            this.dibujo = dibujo;
        }

        @Override
        public void paintIcon(java.awt.Component componente, java.awt.Graphics graphics, int x, int y) {
            java.awt.Graphics2D g2 = (java.awt.Graphics2D) graphics.create();
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

    private static javax.swing.Icon iconoId(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            int celda = 7;
            g2.drawRoundRect(1, 1, celda, celda, 2, 2);
            g2.drawRoundRect(t - celda - 3, 1, celda, celda, 2, 2);
            g2.drawRoundRect(1, t - celda - 3, celda, celda, 2, 2);
            g2.drawRoundRect(t - celda - 3, t - celda - 3, celda, celda, 2, 2);
        });
    }

    private static javax.swing.Icon iconoEdificio(java.awt.Color color, int tamano) {
        return new VectorIcon(tamano, color, (g2, t) -> {
            g2.drawRoundRect(2, 1, t - 6, t - 3, 2, 2);
            for (int fila = 0; fila < 3; fila++) {
                for (int columna = 0; columna < 2; columna++) {
                    g2.drawRect(4 + columna * 5, 3 + fila * 4, 2, 2);
                }
            }
            g2.drawLine(1, t - 2, t - 2, t - 2);
        });
    }

    private static javax.swing.Icon iconoDocumento(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawRoundRect(3, 1, t - 8, t - 2, 2, 2);
            for (int i = 0; i < 3; i++) {
                int y = 5 + i * 3;
                g2.drawLine(5, y, t - 7, y);
            }
        });
    }

    private static javax.swing.Icon iconoCorreo(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawRoundRect(1, 3, t - 3, t - 7, 2, 2);
            g2.drawLine(1, 3, t / 2, t / 2);
            g2.drawLine(t - 2, 3, t / 2, t / 2);
        });
    }

    private static javax.swing.Icon iconoContacto(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawOval(t / 2 - 3, 1, 6, 6);
            java.awt.geom.Arc2D cuerpo = new java.awt.geom.Arc2D.Double(2, 8, t - 4, 11, 0, 180, java.awt.geom.Arc2D.OPEN);
            g2.draw(cuerpo);
        });
    }

    private static javax.swing.Icon iconoMas(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawLine(t / 2, 2, t / 2, t - 2);
            g2.drawLine(2, t / 2, t - 2, t / 2);
        });
    }

    private static javax.swing.Icon iconoGuardar(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawRoundRect(2, 1, t - 4, t - 2, 2, 2);
            g2.drawRect(5, 2, t - 10, 5);
            g2.drawRect(5, t - 6, t - 10, 4);
        });
    }

    private static javax.swing.Icon iconoEliminar(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawRect(5, 5, t - 10, t - 6);
            g2.drawLine(4, 4, t - 4, 4);
            g2.drawLine(7, 2, t - 7, 2);
        });
    }

    private static javax.swing.Icon iconoActualizar(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawArc(2, 2, t - 4, t - 4, 40, 275);
            g2.drawLine(t - 3, 2, t - 3, 7);
            g2.drawLine(t - 3, 2, t - 8, 2);
        });
    }

    private void estilizarBoton(javax.swing.JButton boton, String texto,
            java.awt.Color fondo, java.awt.Color textoColor, int ancho) {
        boton.setText(texto);
        boton.setBackground(fondo);
        boton.setPreferredSize(new java.awt.Dimension(ancho, 38));
        boton.setMinimumSize(new java.awt.Dimension(ancho, 38));
        EstiloAplicacion.aplicarEstiloBoton(boton);
    }

    @SuppressWarnings("unused")
    private void prepararBotonAtras(javax.swing.JButton boton) {
        boton.setText("ATRÁS");
        boton.setIcon(iconoChevronIzquierda(java.awt.Color.WHITE));
        boton.setIconTextGap(8);
        boton.setBackground(COLOR_AZUL);
        EstiloAplicacion.aplicarEstiloBoton(boton);
    }

    private static javax.swing.Icon iconoChevronIzquierda(java.awt.Color color) {
        return new VectorIcon(10, color, (g2, tamano) -> {
            g2.drawLine(tamano - 3, 1, 2, tamano / 2);
            g2.drawLine(2, tamano / 2, tamano - 3, tamano - 1);
        });
    }

    private static final class TarjetaPanel extends javax.swing.JPanel {
        private TarjetaPanel() {
            setOpaque(false);
        }

        @Override
        protected void paintComponent(java.awt.Graphics g) {
            java.awt.Graphics2D g2 = (java.awt.Graphics2D) g.create();
            g2.setColor(java.awt.Color.WHITE);
            g2.fillRoundRect(0, 0, getWidth() - 1, getHeight() - 1, 14, 14);
            g2.dispose();
            super.paintComponent(g);
        }
    }

    // =====================================================================
    // Componentes generados por NetBeans: no modificar manualmente
    // =====================================================================

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
        jLabel1.setText("CONFIGURACIÓN");

        jButton1.setText("CREAR");
        jButton1.addActionListener(this::jButton1ActionPerformed);

        jButton2.setText("EDITAR");
        jButton2.addActionListener(this::jButton2ActionPerformed);

        jButton3.setText("ELIMINAR");
        jButton3.addActionListener(this::jButton3ActionPerformed);

        jButton6.setBackground(new java.awt.Color(0, 95, 131));
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

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(22, 22, 22)
                .addComponent(jButton6)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGap(82, 82, 82))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel7, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jLabel5, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jLabel4, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jLabel3, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jLabel2, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jLabel6, javax.swing.GroupLayout.Alignment.TRAILING))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jTextField1)
                    .addComponent(jTextField2)
                    .addComponent(jTextField3)
                    .addComponent(jTextField4)
                    .addComponent(jTextField5)
                    .addComponent(jTextField6, javax.swing.GroupLayout.DEFAULT_SIZE, 255, Short.MAX_VALUE))
                .addGap(60, 60, 60))
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(104, 104, 104)
                .addComponent(jButton4)
                .addGap(18, 18, 18)
                .addComponent(jButton1)
                .addGap(18, 18, 18)
                .addComponent(jButton2)
                .addGap(18, 18, 18)
                .addComponent(jButton3)
                .addContainerGap(106, Short.MAX_VALUE))
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
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel6))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(jTextField3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4)
                    .addComponent(jTextField4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
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
                .addContainerGap(22, Short.MAX_VALUE))
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
        ventana.setLocationRelativeTo(this);
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
    }//GEN-LAST:event_jButton3ActionPerformed

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

        java.awt.EventQueue.invokeLater(() -> new USUARIOS().setVisible(true));
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton6;
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
    private javax.swing.JTextField jTextField3;
    private javax.swing.JTextField jTextField4;
    private javax.swing.JTextField jTextField5;
    private javax.swing.JTextField jTextField6;
    // End of variables declaration//GEN-END:variables

    // =====================================================================
    // Operaciones CRUD y persistencia de usuarios
    // =====================================================================

    private void asegurarArchivoUsuarios() {
        try {
            usuarioRepository.asegurarArchivoUsuarios();
        } catch (IllegalStateException ex) {
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

            JOptionPane.showMessageDialog(this, "Usuario creado. Contrasena inicial: mismo valor que nombre de usuario.");
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
        String error = UsuarioValidator.validar(
                jTextField1.getText(),
                jTextField2.getText(),
                jTextField3.getText(),
                jTextField4.getText());
        if (error == null) {
            return true;
        }
        JOptionPane.showMessageDialog(this, error);
        return false;
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
        escribirCelda(fila, 3, jTextField3.getText().trim());
        escribirCelda(fila, 4, jTextField4.getText().trim());
        escribirCelda(fila, 5, jTextField5.getText().trim());
        escribirCelda(fila, 6, jTextField6.getText().trim());
    }//guardarDatosFormularioEnFila

    private void cargarFilaEnFormulario(Row fila) {
        jTextField1.setText(leerCelda(fila, 0));
        jTextField2.setText(leerCelda(fila, 2));
        jTextField3.setText(leerCelda(fila, 3));
        jTextField4.setText(leerCelda(fila, 4));
        jTextField5.setText(leerCelda(fila, 5));
        jTextField6.setText(leerCelda(fila, 6));
    }//cargarFilaEnFormulario

    private void limpiarFormulario(boolean conservarId) {
        if (!conservarId) {
            jTextField1.setText("");
        }
        jTextField2.setText("");
        jTextField3.setText("");
        jTextField4.setText("");
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
        return usuarioRepository.asegurarEncabezados(hoja);
    }//asegurarEncabezados

    private Sheet obtenerHojaUsuarios(Workbook libro) {
        return usuarioRepository.obtenerHojaUsuarios(libro);
    }//obtenerHojaUsuarios

    private Workbook cargarLibro(File archivo) throws IOException {
        return usuarioRepository.cargarLibro(archivo);
    }//cargarLibro

    private void guardarLibro(Workbook libro, File archivo) throws IOException {
        usuarioRepository.guardarLibro(libro, archivo);
    }//guardarLibro

    private File obtenerArchivoUsuarios() {
        return usuarioRepository.obtenerArchivoUsuarios();
    }//obtenerArchivoUsuarios

    private String leerCelda(Row fila, int columna) {
        return usuarioRepository.leerCelda(fila, columna);
    }//leerCelda

    private void escribirCelda(Row fila, int columna, String valor) {
        usuarioRepository.escribirCelda(fila, columna, valor);
    }//escribirCelda

    private void mostrarError(String mensajeUsuario, Exception ex) {
        logger.log(java.util.logging.Level.SEVERE, mensajeUsuario, ex);
        JOptionPane.showMessageDialog(this, mensajeUsuario + "\nDetalle tecnico: " + ex.getMessage());
    }//mostrarError
}
