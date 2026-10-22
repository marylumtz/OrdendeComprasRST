package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.EstiloAplicacion;
import com.mycompany.ocxrst.IconoVentanaUtil;
import com.mycompany.ocxrst.config.AppConfig;
import com.mycompany.ocxrst.repository.ActividadRepository;
import com.mycompany.ocxrst.repository.ProveedorRepository;
import com.mycompany.ocxrst.service.InventarioService;
import java.awt.BorderLayout;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Image;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.SwingConstants;
import javax.swing.Timer;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PRINCIPAL extends javax.swing.JFrame {

    private static boolean alertaEntregasMostradaEnSesion;
    private static final DataFormatter FORMATEADOR = new DataFormatter();
    private static final Color AZUL = new Color(30, 57, 130);
    private static final Color TEXTO = new Color(38, 52, 88);
    private final ProveedorRepository proveedorRepository = new ProveedorRepository();
    private final ActividadRepository actividadRepository = new ActividadRepository();
    private JButton jButtonNotificaciones;
    private JButton jButtonMenu;
    private JButton botonMenuActivo;
    private JLabel etiquetaEncabezado;
    private JLabel etiquetaFechaHora;
    private javax.swing.JPopupMenu menuSesion;
    private Timer reloj;
    private Timer actualizadorTablero;
    private int totalNotificaciones;
    private JPanel panelVistas;
    private JPanel cuerpoTablero;
    private JScrollPane scrollTablero;

    public PRINCIPAL() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        setUndecorated(true);
        initComponents();
        configurarMenuSesion();
        actualizarBotonNotificaciones();
        iniciarReloj();
        IconoVentanaUtil.aplicar(this);
        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        setSize(1180, 760);
        setMinimumSize(new Dimension(980, 650));
        setLocationRelativeTo(null);
        mostrarAlertaEntregasUnaVezPorSesion();
    }

    private void initComponents() {
        JPanel fondo = crearPanelConFondo();
        fondo.setLayout(new BorderLayout());
        setContentPane(fondo);
        fondo.add(crearMenuLateral(), BorderLayout.WEST);
        fondo.add(crearContenido(), BorderLayout.CENTER);
    }

    private JPanel crearMenuLateral() {
        JPanel lateral = new JPanel(new BorderLayout());
        lateral.setOpaque(false);
        lateral.setBorder(BorderFactory.createEmptyBorder(86, 0, 0, 0));
        JPanel menu = new JPanel();
        menu.setPreferredSize(new Dimension(186, 627));
        menu.setOpaque(false);
        menu.setBorder(BorderFactory.createEmptyBorder(14, 4, 14, 4));
        menu.setLayout(new BoxLayout(menu, BoxLayout.Y_AXIS));

        menu.add(Box.createVerticalStrut(4));
        JButton inicio = crearBotonMenu("Inicio", "inicio", e -> refrescarTablero());
        marcarBotonActivo(inicio);
        menu.add(inicio);
        menu.add(Box.createVerticalStrut(5));
        menu.add(crearBotonMenu("Mesa de trabajo", "tablero", e -> abrirMesaTrabajo()));
        menu.add(Box.createVerticalStrut(5));
        menu.add(crearBotonMenu("Órdenes de compra", "orden", e -> abrirOrdenes()));
        menu.add(Box.createVerticalStrut(5));
        menu.add(crearBotonMenu("Proveedores", "proveedor", e -> abrirProveedores()));
        menu.add(Box.createVerticalStrut(5));
        menu.add(crearBotonMenu("Inventarios", "inventario", e -> abrirInventarios()));
        menu.add(Box.createVerticalStrut(5));
        menu.add(crearBotonMenu("Usuarios", "usuarios", e -> abrirUsuarios()));
        menu.add(Box.createVerticalStrut(5));
        menu.add(crearBotonMenu("Reportes", "reportes", e -> mostrarPaginaConstruccion()));
        menu.add(Box.createVerticalStrut(5));
        menu.add(crearBotonMenu("Términos y condiciones", "terminos", e -> abrirTerminos()));
        lateral.add(menu, BorderLayout.NORTH);
        etiquetaFechaHora = new JLabel("", SwingConstants.CENTER);
        etiquetaFechaHora.setForeground(AZUL);
        etiquetaFechaHora.setOpaque(false);
        etiquetaFechaHora.setFont(new Font("Segoe UI", Font.PLAIN, 10));
        etiquetaFechaHora.setPreferredSize(new Dimension(340, 20));
        etiquetaFechaHora.setHorizontalAlignment(SwingConstants.RIGHT);
        etiquetaFechaHora.setBorder(BorderFactory.createEmptyBorder(0, 0, 2, 2));
        return lateral;
    }

    private JButton crearBotonMenu(String texto, String tipoIcono, ActionListener accion) {
        JButton boton = new JButton(texto) {
            @Override
            protected void paintComponent(java.awt.Graphics g) {
                java.awt.Graphics2D grafico = (java.awt.Graphics2D) g.create();
                grafico.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                        java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
                grafico.setColor(getBackground());
                grafico.fillRoundRect(0, 0, getWidth(), getHeight(), 12, 12);
                grafico.dispose();
                super.paintComponent(g);
            }
        };
        boton.setAlignmentX(Component.CENTER_ALIGNMENT);
        boton.setMaximumSize(new Dimension(Integer.MAX_VALUE, 38));
        boton.setPreferredSize(new Dimension(167, 38));
        boton.setHorizontalAlignment(SwingConstants.LEFT);
        boton.setForeground(Color.WHITE);
        boton.setBackground(new Color(75, 116, 218, 235));
        boton.setIcon(new IconoMenu(tipoIcono));
        boton.setIconTextGap(10);
        boton.setFont(new Font("Segoe UI", Font.BOLD, 11));
        boton.setBorder(BorderFactory.createEmptyBorder(6, 12, 6, 4));
        boton.setFocusPainted(false);
        boton.setContentAreaFilled(false);
        boton.setOpaque(false);
        boton.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        boton.addActionListener(e -> {
            marcarBotonActivo(boton);
            accion.actionPerformed(e);
        });
        return boton;
    }

    private static final class IconoMenu implements javax.swing.Icon {
        private final String tipo;

        private IconoMenu(String tipo) {
            this.tipo = tipo;
        }

        @Override
        public int getIconWidth() {
            return 20;
        }

        @Override
        public int getIconHeight() {
            return 20;
        }

        @Override
        public void paintIcon(Component componente, java.awt.Graphics g, int x, int y) {
            java.awt.Graphics2D grafico = (java.awt.Graphics2D) g.create();
            grafico.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                    java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
            grafico.setColor(componente.getForeground());
            grafico.setStroke(new java.awt.BasicStroke(1.8f));

            switch (tipo) {
                case "inicio":
                    int[] techoX = {x + 2, x + 10, x + 18};
                    int[] techoY = {y + 9, y + 2, y + 9};
                    grafico.fillPolygon(techoX, techoY, 3);
                    grafico.fillRect(x + 5, y + 8, 10, 9);
                    grafico.setColor(componente.getBackground());
                    grafico.fillRect(x + 9, y + 12, 3, 5);
                    break;
                case "tablero":
                    for (int fila = 0; fila < 2; fila++) {
                        for (int columna = 0; columna < 2; columna++) {
                            grafico.fillRoundRect(x + 2 + columna * 8, y + 2 + fila * 8, 6, 6, 2, 2);
                        }
                    }
                    break;
                case "orden":
                case "terminos":
                    grafico.drawRoundRect(x + 4, y + 2, 12, 16, 2, 2);
                    grafico.drawLine(x + 7, y + 7, x + 13, y + 7);
                    grafico.drawLine(x + 7, y + 10, x + 13, y + 10);
                    grafico.drawLine(x + 7, y + 13, x + 11, y + 13);
                    break;
                case "pendiente":
                    grafico.drawOval(x + 2, y + 2, 16, 16);
                    grafico.drawLine(x + 10, y + 5, x + 10, y + 10);
                    grafico.drawLine(x + 10, y + 10, x + 14, y + 12);
                    break;
                case "alerta":
                    int[] alertaX = {x + 10, x + 2, x + 18};
                    int[] alertaY = {y + 2, y + 18, y + 18};
                    grafico.fillPolygon(alertaX, alertaY, 3);
                    grafico.setColor(componente.getBackground());
                    grafico.fillRect(x + 9, y + 7, 2, 6);
                    grafico.fillRect(x + 9, y + 15, 2, 2);
                    break;
                case "proveedor":
                    grafico.drawRect(x + 2, y + 7, 13, 7);
                    grafico.drawLine(x + 15, y + 9, x + 18, y + 9);
                    grafico.drawLine(x + 18, y + 9, x + 18, y + 14);
                    grafico.drawOval(x + 4, y + 13, 4, 4);
                    grafico.drawOval(x + 14, y + 13, 4, 4);
                    break;
                case "inventario":
                    int[] cuboX = {x + 2, x + 10, x + 18, x + 10};
                    int[] cuboY = {y + 6, y + 2, y + 6, y + 10};
                    grafico.fillPolygon(cuboX, cuboY, 4);
                    grafico.drawLine(x + 2, y + 6, x + 2, y + 15);
                    grafico.drawLine(x + 2, y + 15, x + 10, y + 19);
                    grafico.drawLine(x + 18, y + 6, x + 18, y + 15);
                    grafico.drawLine(x + 18, y + 15, x + 10, y + 19);
                    grafico.drawLine(x + 10, y + 10, x + 10, y + 19);
                    break;
                case "usuarios":
                    grafico.fillOval(x + 7, y + 2, 6, 6);
                    grafico.fillOval(x + 1, y + 5, 5, 5);
                    grafico.fillOval(x + 14, y + 5, 5, 5);
                    grafico.fillRoundRect(x + 5, y + 9, 10, 8, 4, 4);
                    grafico.fillRoundRect(x, y + 11, 7, 6, 3, 3);
                    grafico.fillRoundRect(x + 13, y + 11, 7, 6, 3, 3);
                    break;
                case "reportes":
                    grafico.fillRect(x + 2, y + 10, 4, 8);
                    grafico.fillRect(x + 8, y + 6, 4, 12);
                    grafico.fillRect(x + 14, y + 2, 4, 16);
                    break;
                default:
                    break;
            }
            grafico.dispose();
        }
    }

    private void marcarBotonActivo(JButton boton) {
        if (botonMenuActivo != null) {
            botonMenuActivo.setBackground(new Color(75, 116, 218, 235));
            botonMenuActivo.setForeground(Color.WHITE);
        }
        botonMenuActivo = boton;
        botonMenuActivo.setBackground(new Color(255, 255, 255, 235));
        botonMenuActivo.setForeground(AZUL);
    }

    private JPanel crearContenido() {
        JPanel contenido = new JPanel(new BorderLayout(0, 12));
        contenido.setOpaque(false);
        contenido.setBorder(BorderFactory.createEmptyBorder(14, 8, 10, 8));

        JPanel encabezado = new JPanel(new BorderLayout());
        encabezado.setOpaque(false);
        etiquetaEncabezado = new JLabel();
        etiquetaEncabezado.setForeground(AZUL);
        etiquetaEncabezado.setFont(new Font("Segoe UI", Font.BOLD, 18));
        etiquetaEncabezado.setBorder(BorderFactory.createEmptyBorder(48, 0, 0, 8));
        actualizarEncabezado("INICIO");
        encabezado.add(etiquetaEncabezado, BorderLayout.WEST);

        etiquetaFechaHora.setForeground(Color.WHITE);
        etiquetaFechaHora.setFont(new Font("Segoe UI", Font.BOLD, 14));
        etiquetaFechaHora.setPreferredSize(new Dimension(310, 38));
        etiquetaFechaHora.setHorizontalAlignment(SwingConstants.CENTER);
        etiquetaFechaHora.setBorder(BorderFactory.createEmptyBorder(0, 0, 0, 0));

        JPanel acciones = new JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 8, 0));
        acciones.setOpaque(false);
        acciones.add(etiquetaFechaHora);
        jButtonNotificaciones = new JButton();
        jButtonNotificaciones.setFont(new Font("Segoe UI", Font.BOLD, 12));
        jButtonNotificaciones.setForeground(Color.WHITE);
        jButtonNotificaciones.setFocusPainted(false);
        jButtonNotificaciones.setBorder(BorderFactory.createEmptyBorder(7, 12, 7, 12));
        jButtonNotificaciones.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jButtonNotificaciones.addActionListener(e -> PROVEEDORES.mostrarAlertaEntregasPendientes(this));
        jButtonMenu = new JButton(nombreUsuario());
        jButtonMenu.setIcon(new IconoMenu("usuarios"));
        jButtonMenu.setIconTextGap(7);
        jButtonMenu.setFont(new Font("Segoe UI", Font.BOLD, 12));
        jButtonMenu.setForeground(AZUL);
        jButtonMenu.setBackground(Color.WHITE);
        jButtonMenu.setFocusPainted(false);
        jButtonMenu.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jButtonMenu.addActionListener(e -> menuSesion.show(jButtonMenu, 0, jButtonMenu.getHeight()));
        JButton botonCerrar = new JButton("×");
        botonCerrar.setFont(new Font("Segoe UI", Font.BOLD, 20));
        botonCerrar.setForeground(Color.WHITE);
        botonCerrar.setContentAreaFilled(false);
        botonCerrar.setBorder(BorderFactory.createEmptyBorder(0, 8, 0, 2));
        botonCerrar.setFocusPainted(false);
        botonCerrar.setToolTipText("Cerrar");
        botonCerrar.addActionListener(e -> confirmarSalida());
        acciones.add(jButtonNotificaciones);
        acciones.add(jButtonMenu);
        acciones.add(botonCerrar);
        encabezado.add(acciones, BorderLayout.EAST);
        contenido.add(encabezado, BorderLayout.NORTH);

        cuerpoTablero = new JPanel();
        cuerpoTablero.setOpaque(false);
        cuerpoTablero.setBackground(Color.WHITE);
        cuerpoTablero.setLayout(new BoxLayout(cuerpoTablero, BoxLayout.Y_AXIS));
        reconstruirTablero();
        scrollTablero = new JScrollPane(cuerpoTablero);
        scrollTablero.setBorder(null);
        scrollTablero.setHorizontalScrollBarPolicy(javax.swing.JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        scrollTablero.setOpaque(false);
        scrollTablero.getViewport().setOpaque(false);
        panelVistas = new JPanel(new CardLayout());
        panelVistas.setOpaque(false);
        panelVistas.add(scrollTablero, "INICIO");
        panelVistas.add(crearPaginaConstruccion(), "CONSTRUCCION");
        contenido.add(panelVistas, BorderLayout.CENTER);
        return contenido;
    }

    private void reconstruirTablero() {
        cuerpoTablero.removeAll();
        cuerpoTablero.add(crearTarjetasIndicadores());
        cuerpoTablero.add(Box.createVerticalStrut(10));
        JPanel centro = new JPanel(new GridLayout(1, 2, 12, 0));
        centro.setOpaque(false);
        centro.setPreferredSize(new Dimension(0, 194));
        centro.setMaximumSize(new Dimension(Integer.MAX_VALUE, 194));
        centro.add(crearPanelTabla("Mis pendientes", cargarPendientes()));
        centro.add(crearPanelActividadReciente());
        cuerpoTablero.add(centro);
        cuerpoTablero.add(Box.createVerticalStrut(10));
        JPanel operaciones = new JPanel(new GridLayout(1, 2, 12, 0));
        operaciones.setOpaque(false);
        operaciones.setPreferredSize(new Dimension(0, 190));
        operaciones.setMaximumSize(new Dimension(Integer.MAX_VALUE, 190));
        operaciones.add(crearPanelTabla("Stock bajo", cargarStockBajo()));
        operaciones.add(crearPanelAccesos());
        cuerpoTablero.add(operaciones);
        cuerpoTablero.add(Box.createVerticalStrut(10));
        JPanel resumen = new JPanel(new GridLayout(1, 2, 12, 0));
        resumen.setOpaque(false);
        resumen.setPreferredSize(new Dimension(0, 138));
        resumen.setMaximumSize(new Dimension(Integer.MAX_VALUE, 138));
        resumen.add(crearPanelAlertas());
        resumen.add(crearPanelResumenGraficos());
        cuerpoTablero.add(resumen);
        cuerpoTablero.revalidate();
        cuerpoTablero.repaint();
    }

    private JPanel crearPaginaConstruccion() {
        JPanel pagina = new JPanel(new GridBagLayout());
        pagina.setOpaque(false);
        JLabel mensaje = new JLabel("Página en construcción", SwingConstants.CENTER);
        mensaje.setForeground(AZUL);
        mensaje.setFont(new Font("Segoe UI", Font.BOLD, 24));
        pagina.add(mensaje);
        return pagina;
    }

    private JPanel crearTarjetasIndicadores() {
        JPanel tarjetas = new JPanel(new GridLayout(1, 4, 10, 0));
        tarjetas.setOpaque(false);
        int ordenes = PROVEEDORES.contarOrdenesRegistradas();
        int pendientes = PROVEEDORES.contarEntregasPorRevisar();
        int alertas = PROVEEDORES.contarAlertasEntregasPendientes();
        int stockBajo = contarStockBajo();
        tarjetas.add(crearTarjeta(String.valueOf(ordenes), "Órdenes de compra", "Total registradas",
                new Color(224, 234, 255), "orden", new Color(61, 96, 190)));
        tarjetas.add(crearTarjeta(String.valueOf(pendientes), "Entregas por revisar", "Requieren seguimiento",
                new Color(255, 237, 239), "pendiente", new Color(210, 48, 65)));
        tarjetas.add(crearTarjeta(String.valueOf(alertas), "Alertas de entrega y pago", "Requieren atención",
                new Color(255, 220, 220), "alerta", new Color(225, 76, 85)));
        tarjetas.add(crearTarjeta(String.valueOf(stockBajo), "Productos con stock bajo", "Necesitan reposición",
                new Color(222, 246, 220), "inventario", new Color(35, 161, 91)));
        tarjetas.setPreferredSize(new Dimension(0, 86));
        tarjetas.setMaximumSize(new Dimension(Integer.MAX_VALUE, 86));
        return tarjetas;
    }

    private JPanel crearTarjeta(String valor, String titulo, String detalle, Color fondo,
            String tipoIcono, Color colorIcono) {
        JPanel tarjeta = new JPanel(new BorderLayout(10, 0));
        tarjeta.setBackground(fondo);
        tarjeta.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(210, 220, 235)),
                BorderFactory.createEmptyBorder(9, 10, 9, 8)));
        tarjeta.setToolTipText(titulo + ": " + valor + " · " + detalle);
        tarjeta.setPreferredSize(new Dimension(180, 86));

        JPanel iconoContenedor = new JPanel(new GridBagLayout());
        iconoContenedor.setPreferredSize(new Dimension(44, 44));
        iconoContenedor.setBackground(new Color(colorIcono.getRed(), colorIcono.getGreen(), colorIcono.getBlue(), 28));
        iconoContenedor.setBorder(BorderFactory.createLineBorder(new Color(colorIcono.getRed(), colorIcono.getGreen(), colorIcono.getBlue(), 40)));
        JLabel icono = new JLabel(new IconoMenu(tipoIcono));
        icono.setForeground(colorIcono);
        iconoContenedor.add(icono);
        tarjeta.add(iconoContenedor, BorderLayout.WEST);

        JPanel contenido = new JPanel();
        contenido.setOpaque(false);
        contenido.setLayout(new BoxLayout(contenido, BoxLayout.Y_AXIS));
        JLabel numero = new JLabel(valor);
        numero.setFont(new Font("Segoe UI", Font.BOLD, 23));
        numero.setForeground(AZUL);
        JLabel nombre = new JLabel(titulo);
        nombre.setFont(new Font("Segoe UI", Font.BOLD, 11));
        nombre.setForeground(TEXTO);
        JLabel descripcion = new JLabel("<html><span style='white-space:nowrap'>" + escaparHtml(detalle) + "</span></html>");
        descripcion.setFont(new Font("Segoe UI", Font.PLAIN, 10));
        descripcion.setForeground(new Color(90, 105, 125));
        contenido.add(numero);
        contenido.add(nombre);
        contenido.add(descripcion);
        tarjeta.add(contenido, BorderLayout.CENTER);

        JLabel flecha = new JLabel("›");
        flecha.setFont(new Font("Segoe UI", Font.BOLD, 20));
        flecha.setForeground(new Color(72, 92, 125));
        tarjeta.add(flecha, BorderLayout.EAST);
        return tarjeta;
    }

    private String escaparHtml(String texto) {
        return texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }

    private JPanel crearPanelTabla(String titulo, Object[][] datos) {
        JPanel panel = new JPanel(new BorderLayout(0, 6));
        panel.setBackground(new Color(255, 255, 255, 235));
        panel.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(215, 224, 238)),
                BorderFactory.createEmptyBorder(8, 8, 8, 8)));
        JPanel barraTitulo = new JPanel(new BorderLayout());
        barraTitulo.setOpaque(false);
        JLabel encabezado = new JLabel(titulo);
        encabezado.setIcon(new IconoMenu("Stock bajo".equals(titulo) ? "inventario" : "orden"));
        encabezado.setIconTextGap(7);
        encabezado.setFont(new Font("Segoe UI", Font.BOLD, 14));
        encabezado.setForeground(AZUL);
        barraTitulo.add(encabezado, BorderLayout.WEST);
        JButton verModulo = new JButton("Stock bajo".equals(titulo) ? "Ver inventario" : "Ver órdenes");
        prepararEnlacePanel(verModulo);
        verModulo.addActionListener(e -> {
            if ("Stock bajo".equals(titulo)) {
                abrirInventarios();
            } else {
                abrirOrdenes();
            }
        });
        barraTitulo.add(verModulo, BorderLayout.EAST);
        panel.add(barraTitulo, BorderLayout.NORTH);
        String[] columnas = "Stock bajo".equals(titulo)
                ? new String[]{"Código", "Producto", "Existencia", "Mínimo"}
                : new String[]{"Folio", "Estado", "Proveedor", "Fecha"};
        JTable tabla = new JTable(new DefaultTableModel(datos, columnas) {
            @Override public boolean isCellEditable(int fila, int columna) { return false; }
        });
        tabla.setRowHeight("Stock bajo".equals(titulo) ? 24 : 25);
        tabla.setFont(new Font("Segoe UI", Font.PLAIN, 10));
        tabla.setForeground(TEXTO);
        tabla.setGridColor(new Color(229, 235, 243));
        tabla.setShowVerticalLines(false);
        tabla.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 10));
        tabla.getTableHeader().setForeground(AZUL);
        tabla.getTableHeader().setBackground(new Color(239, 244, 250));
        tabla.getTableHeader().setReorderingAllowed(false);
        panel.add(new JScrollPane(tabla), BorderLayout.CENTER);
        return panel;
    }

    private JPanel crearPanelActividadReciente() {
        JPanel panel = new JPanel(new BorderLayout(0, 5));
        panel.setBackground(new Color(255, 255, 255, 235));
        panel.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(215, 224, 238)),
                BorderFactory.createEmptyBorder(7, 8, 7, 8)));

        JPanel encabezado = new JPanel(new BorderLayout(6, 0));
        encabezado.setOpaque(false);
        JLabel titulo = new JLabel("Actividad reciente");
        titulo.setIcon(new IconoMenu("tablero"));
        titulo.setIconTextGap(6);
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 14));
        titulo.setForeground(AZUL);
        encabezado.add(titulo, BorderLayout.WEST);

        JButton verTodas = new JButton("Ver órdenes");
        prepararEnlacePanel(verTodas);
        verTodas.addActionListener(e -> abrirOrdenes());
        encabezado.add(verTodas, BorderLayout.EAST);
        panel.add(encabezado, BorderLayout.NORTH);

        String[] columnas = {"Fecha", "Módulo", "Actividad", "Folio / Referencia", "Usuario"};
        Object[][] actividades = cargarActividad();
        DefaultTableModel modelo = new DefaultTableModel(actividades, columnas) {
            @Override public boolean isCellEditable(int fila, int columna) { return false; }
        };
        JTable tabla = new JTable(modelo);
        tabla.setRowHeight(24);
        tabla.setFont(new Font("Segoe UI", Font.PLAIN, 10));
        tabla.setForeground(TEXTO);
        tabla.setGridColor(new Color(232, 229, 221));
        tabla.setShowVerticalLines(false);
        tabla.setIntercellSpacing(new Dimension(1, 0));
        tabla.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 10));
        tabla.getTableHeader().setForeground(AZUL);
        tabla.getTableHeader().setBackground(new Color(241, 239, 247));
        tabla.getTableHeader().setReorderingAllowed(false);
        tabla.getColumnModel().getColumn(0).setPreferredWidth(86);
        tabla.getColumnModel().getColumn(1).setPreferredWidth(95);
        tabla.getColumnModel().getColumn(2).setPreferredWidth(112);
        tabla.getColumnModel().getColumn(3).setPreferredWidth(105);
        tabla.getColumnModel().getColumn(4).setPreferredWidth(92);
        panel.add(new JScrollPane(tabla), BorderLayout.CENTER);
        return panel;
    }

    private void prepararEnlacePanel(JButton boton) {
        boton.setFont(new Font("Segoe UI", Font.BOLD, 10));
        boton.setForeground(new Color(30, 87, 170));
        boton.setBorderPainted(false);
        boton.setContentAreaFilled(false);
        boton.setFocusPainted(false);
        boton.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
    }

    private JPanel crearPanelAccesos() {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setBackground(new Color(255, 255, 255, 235));
        panel.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(215, 224, 238)),
                BorderFactory.createEmptyBorder(8, 8, 8, 8)));
        JLabel titulo = new JLabel("Accesos rápidos", new IconoMenu("tablero"), SwingConstants.LEFT);
        titulo.setIconTextGap(7);
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 14));
        titulo.setForeground(AZUL);
        panel.add(titulo, BorderLayout.NORTH);
        JPanel botones = new JPanel(new GridLayout(2, 2, 8, 8));
        botones.setOpaque(false);
        botones.add(crearAcceso("Órdenes de compra", "orden", new Color(224, 239, 255),
                new Color(30, 93, 188), e -> abrirOrdenes()));
        botones.add(crearAcceso("Proveedores", "proveedor", new Color(235, 229, 255),
                new Color(90, 54, 180), e -> abrirProveedores()));
        botones.add(crearAcceso("Inventarios", "inventario", new Color(222, 246, 239),
                new Color(18, 126, 101), e -> abrirInventarios()));
        botones.add(crearAcceso("Usuarios", "usuarios", new Color(255, 239, 216),
                new Color(177, 103, 12), e -> abrirUsuarios()));
        panel.add(botones, BorderLayout.CENTER);
        return panel;
    }

    private JButton crearAcceso(String texto, String tipoIcono, Color fondo,
            Color color, ActionListener accion) {
        JButton boton = new JButton(texto, new IconoMenu(tipoIcono));
        boton.setHorizontalAlignment(SwingConstants.CENTER);
        boton.setHorizontalTextPosition(SwingConstants.CENTER);
        boton.setVerticalTextPosition(SwingConstants.BOTTOM);
        boton.setIconTextGap(5);
        boton.setForeground(color);
        boton.setBackground(fondo);
        boton.setFont(new Font("Segoe UI", Font.BOLD, 10));
        boton.setFocusPainted(false);
        boton.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(210, 220, 235)),
                BorderFactory.createEmptyBorder(5, 4, 5, 4)));
        boton.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        boton.addActionListener(accion);
        return boton;
    }

    private JPanel crearPanelAlertas() {
        JPanel panel = new JPanel(new BorderLayout(0, 6));
        panel.setBackground(new Color(255, 255, 255, 235));
        panel.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(215, 224, 238)),
                BorderFactory.createEmptyBorder(8, 8, 7, 8)));

        int entregasPorRevisar = PROVEEDORES.contarEntregasPorRevisar();
        int alertasComerciales = PROVEEDORES.contarAlertasEntregasPendientes();
        int productosBajos = contarStockBajo();
        boolean hayAlertas = entregasPorRevisar > 0 || alertasComerciales > 0 || productosBajos > 0;
        JPanel encabezado = new JPanel(new BorderLayout());
        encabezado.setOpaque(false);
        JLabel titulo = new JLabel("Alertas del sistema", new IconoMenu("alerta"), SwingConstants.LEFT);
        titulo.setIconTextGap(7);
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 14));
        titulo.setForeground(AZUL);
        encabezado.add(titulo, BorderLayout.WEST);
        JButton verAlertas = new JButton(alertasComerciales > 0 ? "Ver detalles"
                : productosBajos > 0 ? "Ver inventario" : "Ver órdenes");
        prepararEnlacePanel(verAlertas);
        verAlertas.setVisible(hayAlertas);
        verAlertas.addActionListener(e -> {
            if (alertasComerciales > 0) {
                PROVEEDORES.mostrarAlertaEntregasPendientes(this);
            } else if (productosBajos > 0) {
                abrirInventarios();
            } else {
                abrirOrdenes();
            }
        });
        encabezado.add(verAlertas, BorderLayout.EAST);
        panel.add(encabezado, BorderLayout.NORTH);

        JPanel filas = new JPanel();
        filas.setOpaque(false);
        filas.setLayout(new BoxLayout(filas, BoxLayout.Y_AXIS));
        if (!hayAlertas) {
            filas.add(crearFilaAlerta("Todo está al día", "No hay pendientes que requieran atención.",
                    "OK", new Color(35, 161, 91)));
        } else {
            if (entregasPorRevisar > 0) {
                filas.add(crearFilaAlerta("Entregas por revisar", "Órdenes pendientes de confirmar.",
                        String.valueOf(entregasPorRevisar), new Color(208, 68, 76)));
            }
            if (alertasComerciales > 0) {
                filas.add(crearFilaAlerta("Vencimientos y pagos", "Alertas de entrega o pago registradas.",
                        String.valueOf(alertasComerciales), new Color(218, 155, 0)));
            }
            if (productosBajos > 0) {
                filas.add(crearFilaAlerta("Stock bajo", "Productos que alcanzaron su mínimo.",
                        String.valueOf(productosBajos), new Color(30, 115, 190)));
            }
        }
        panel.add(filas, BorderLayout.CENTER);
        return panel;
    }

    private JPanel crearFilaAlerta(String titulo, String detalle, String cantidad, Color color) {
        JPanel fila = new JPanel(new BorderLayout(8, 0));
        fila.setOpaque(false);
        fila.setBorder(BorderFactory.createMatteBorder(0, 0, 1, 0, new Color(232, 237, 244)));
        JPanel texto = new JPanel();
        texto.setOpaque(false);
        texto.setLayout(new BoxLayout(texto, BoxLayout.Y_AXIS));
        JLabel nombre = new JLabel(titulo);
        nombre.setFont(new Font("Segoe UI", Font.BOLD, 10));
        nombre.setForeground(TEXTO);
        JLabel descripcion = new JLabel(detalle);
        descripcion.setFont(new Font("Segoe UI", Font.PLAIN, 9));
        descripcion.setForeground(new Color(95, 108, 128));
        texto.add(nombre);
        texto.add(descripcion);
        fila.add(texto, BorderLayout.CENTER);

        JLabel valor = new JLabel(cantidad, SwingConstants.CENTER);
        valor.setFont(new Font("Segoe UI", Font.BOLD, 10));
        valor.setForeground(color);
        valor.setOpaque(true);
        valor.setBackground(new Color(color.getRed(), color.getGreen(), color.getBlue(), 24));
        valor.setBorder(BorderFactory.createEmptyBorder(3, 8, 3, 8));
        fila.add(valor, BorderLayout.EAST);
        fila.setPreferredSize(new Dimension(0, 29));
        fila.setMaximumSize(new Dimension(Integer.MAX_VALUE, 29));
        return fila;
    }

    private JPanel crearPanelResumenGraficos() {
        JPanel panel = new JPanel(new BorderLayout(0, 6));
        panel.setBackground(new Color(255, 255, 255, 235));
        panel.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(215, 224, 238)),
                BorderFactory.createEmptyBorder(8, 8, 7, 8)));
        JLabel titulo = new JLabel("Gráficos y resumen", new IconoMenu("reportes"), SwingConstants.LEFT);
        titulo.setIconTextGap(7);
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 14));
        titulo.setForeground(AZUL);
        panel.add(titulo, BorderLayout.NORTH);

        int ordenes = PROVEEDORES.contarOrdenesRegistradas();
        int pendientes = Math.min(ordenes, PROVEEDORES.contarEntregasPorRevisar());
        int productos = contarInventarioActivo();
        int stockBajo = Math.min(productos, contarStockBajo());
        JPanel graficas = new JPanel(new GridLayout(1, 2, 8, 0));
        graficas.setOpaque(false);
        graficas.add(crearResumenDona("Órdenes de compra", pendientes, ordenes - pendientes,
                "Por revisar", "Revisadas", new Color(255, 190, 55), new Color(51, 142, 229)));
        graficas.add(crearResumenDona("Estado del inventario", stockBajo, productos - stockBajo,
                "Stock bajo", "En rango", new Color(239, 79, 79), new Color(44, 185, 132)));
        panel.add(graficas, BorderLayout.CENTER);
        return panel;
    }

    private JPanel crearResumenDona(String titulo, int primero, int segundo, String leyendaPrimera,
            String leyendaSegunda, Color colorPrimero, Color colorSegundo) {
        JPanel resumen = new JPanel(new BorderLayout(6, 2));
        resumen.setOpaque(false);
        JLabel encabezado = new JLabel(titulo);
        encabezado.setFont(new Font("Segoe UI", Font.BOLD, 10));
        encabezado.setForeground(TEXTO);
        resumen.add(encabezado, BorderLayout.NORTH);

        JPanel contenido = new JPanel(new BorderLayout(7, 0));
        contenido.setOpaque(false);
        GraficaDona dona = new GraficaDona(primero, segundo, colorPrimero, colorSegundo);
        dona.setToolTipText(titulo + ": " + (primero + segundo) + " registros");
        contenido.add(dona, BorderLayout.WEST);
        JPanel leyendas = new JPanel();
        leyendas.setOpaque(false);
        leyendas.setLayout(new BoxLayout(leyendas, BoxLayout.Y_AXIS));
        leyendas.add(crearLeyenda(colorPrimero, leyendaPrimera, primero));
        leyendas.add(Box.createVerticalStrut(5));
        leyendas.add(crearLeyenda(colorSegundo, leyendaSegunda, segundo));
        contenido.add(leyendas, BorderLayout.CENTER);
        resumen.add(contenido, BorderLayout.CENTER);
        return resumen;
    }

    private JPanel crearLeyenda(Color color, String texto, int cantidad) {
        JPanel leyenda = new JPanel(new BorderLayout(5, 0));
        leyenda.setOpaque(false);
        JPanel muestra = new JPanel();
        muestra.setBackground(color);
        muestra.setPreferredSize(new Dimension(9, 9));
        muestra.setMaximumSize(new Dimension(9, 9));
        leyenda.add(muestra, BorderLayout.WEST);
        JLabel etiqueta = new JLabel(texto + "  " + cantidad);
        etiqueta.setFont(new Font("Segoe UI", Font.PLAIN, 9));
        etiqueta.setForeground(TEXTO);
        leyenda.add(etiqueta, BorderLayout.CENTER);
        return leyenda;
    }

    private static final class GraficaDona extends JPanel {
        private final int primero;
        private final int segundo;
        private final Color colorPrimero;
        private final Color colorSegundo;

        GraficaDona(int primero, int segundo, Color colorPrimero, Color colorSegundo) {
            this.primero = Math.max(0, primero);
            this.segundo = Math.max(0, segundo);
            this.colorPrimero = colorPrimero;
            this.colorSegundo = colorSegundo;
            setOpaque(false);
            setPreferredSize(new Dimension(62, 62));
            setMinimumSize(new Dimension(50, 50));
        }

        @Override
        protected void paintComponent(java.awt.Graphics g) {
            super.paintComponent(g);
            java.awt.Graphics2D grafico = (java.awt.Graphics2D) g.create();
            grafico.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                    java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
            int diametro = Math.max(1, Math.min(getWidth(), getHeight()) - 12);
            int x = (getWidth() - diametro) / 2;
            int y = (getHeight() - diametro) / 2;
            int total = primero + segundo;
            grafico.setStroke(new java.awt.BasicStroke(8f, java.awt.BasicStroke.CAP_BUTT,
                    java.awt.BasicStroke.JOIN_ROUND));
            if (total == 0) {
                grafico.setColor(new Color(225, 232, 240));
                grafico.drawArc(x, y, diametro, diametro, 0, 360);
            } else {
                int recorridoPrimero = Math.round(360f * primero / total);
                grafico.setColor(colorPrimero);
                grafico.drawArc(x, y, diametro, diametro, 90, -recorridoPrimero);
                grafico.setColor(colorSegundo);
                grafico.drawArc(x, y, diametro, diametro, 90 - recorridoPrimero,
                        -(360 - recorridoPrimero));
            }
            String etiqueta = String.valueOf(total);
            grafico.setColor(AZUL);
            grafico.setFont(new Font("Segoe UI", Font.BOLD, 13));
            java.awt.FontMetrics metricas = grafico.getFontMetrics();
            grafico.drawString(etiqueta, (getWidth() - metricas.stringWidth(etiqueta)) / 2,
                    (getHeight() + metricas.getAscent() - metricas.getDescent()) / 2);
            grafico.dispose();
        }
    }

    private Object[][] cargarPendientes() {
        List<Object[]> filas = new ArrayList<>();
        leerFilas(AppConfig.registroCFile(), (row, formato) -> {
            if (filas.size() < 5 && !leer(row, 0).isBlank() && !"SI".equalsIgnoreCase(leer(row, 24))) {
                filas.add(new Object[]{leer(row, 0), "Entrega pendiente", proveedorRepository.obtenerNombrePorId(leer(row, 4)), leer(row, 1)});
            }
        });
        return filas.toArray(new Object[0][]);
    }

    private Object[][] cargarActividad() {
        List<Object[]> filas = new ArrayList<>();
        leerFilas(actividadRepository.obtenerArchivoActividades(), (row, formato) -> {
            if (filas.size() < 5 && !leer(row, 0).isBlank()) {
                filas.add(new Object[]{leer(row, 0), leer(row, 1), leer(row, 2),
                    leer(row, 3), leer(row, 4)});
            }
        });
        leerFilas(AppConfig.registroCFile(), (row, formato) -> {
            if (filas.size() < 5 && !leer(row, 0).isBlank()) {
                String usuario = leer(row, 15);
                if (usuario.isBlank()) {
                    usuario = nombreUsuario();
                }
                filas.add(new Object[]{leer(row, 1), "Órdenes de compra", "Orden registrada",
                    leer(row, 0), usuario});
            }
        });
        return filas.toArray(new Object[0][]);
    }

    private Object[][] cargarStockBajo() {
        List<Object[]> filas = new ArrayList<>();
        leerFilas(AppConfig.inventariosFile(), (row, formato) -> {
            String codigo = leer(row, 0);
            if (filas.size() < 6 && !codigo.isBlank() && esStockBajo(leer(row, 4), leer(row, 8))) {
                filas.add(new Object[]{codigo, leer(row, 1), leer(row, 4), leer(row, 8)});
            }
        });
        return filas.toArray(new Object[0][]);
    }

    private boolean esStockBajo(String cantidad, String minimo) {
        try { return Double.parseDouble(cantidad.replace(",", "")) <= Double.parseDouble(minimo.replace(",", "")); }
        catch (NumberFormatException ex) { return false; }
    }

    private int contarStockBajo() {
        int[] total = {0};
        leerFilas(AppConfig.inventariosFile(), (row, formato) -> {
            if (!leer(row, 0).isBlank() && esStockBajo(leer(row, 4), leer(row, 8))
                    && "1".equals(InventarioService.normalizarEstado(leer(row, 5)))) total[0]++;
        });
        return total[0];
    }

    private int contarInventarioActivo() {
        int[] total = {0};
        leerFilas(AppConfig.inventariosFile(), (row, formato) -> {
            if (!leer(row, 0).isBlank()
                    && "1".equals(InventarioService.normalizarEstado(leer(row, 5)))) {
                total[0]++;
            }
        });
        return total[0];
    }

    @FunctionalInterface private interface FilaExcel { void aceptar(Row fila, DataFormatter formato); }
    private void leerFilas(File archivo, FilaExcel accion) {
        if (!archivo.exists()) return;
        try (FileInputStream entrada = new FileInputStream(archivo); Workbook libro = new XSSFWorkbook(entrada)) {
            if (libro.getNumberOfSheets() == 0) return;
            for (Row fila : libro.getSheetAt(0)) if (fila.getRowNum() > 0) accion.aceptar(fila, FORMATEADOR);
        } catch (Exception ex) {
            java.util.logging.Logger.getLogger(PRINCIPAL.class.getName()).log(java.util.logging.Level.WARNING,
                    "No se pudo leer el archivo del tablero: " + archivo.getName(), ex);
        }
    }

    private static String leer(Row fila, int columna) {
        return fila == null || fila.getCell(columna) == null ? "" : FORMATEADOR.formatCellValue(fila.getCell(columna)).trim();
    }

    private void configurarMenuSesion() {
        menuSesion = new javax.swing.JPopupMenu();
        menuSesion.setBackground(Color.WHITE);
        JButton cerrar = new JButton("CERRAR SESIÓN");
        cerrar.setFocusPainted(true);
        cerrar.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        cerrar.addActionListener(e -> cerrarSesion());
        menuSesion.add(cerrar);
    }

    private String nombreUsuario() {
        return INICIARSESION.nombreUsuarioActual == null || INICIARSESION.nombreUsuarioActual.isBlank()
                ? INICIARSESION.usuarioActual : INICIARSESION.nombreUsuarioActual;
    }

    private void actualizarBotonNotificaciones() {
        totalNotificaciones = PROVEEDORES.contarAlertasEntregasPendientes();
        jButtonNotificaciones.setText("NOTIFICACIONES (" + totalNotificaciones + ")");
        jButtonNotificaciones.setBackground(totalNotificaciones > 0 ? new Color(190, 45, 55) : new Color(55, 105, 175));
    }

    private void iniciarReloj() {
        actualizarFechaHora();
        reloj = new Timer(1000, e -> actualizarFechaHora());
        reloj.start();
        actualizadorTablero = new Timer(5000, e -> actualizarTableroSilencioso());
        actualizadorTablero.setRepeats(true);
        actualizadorTablero.start();
        addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosed(java.awt.event.WindowEvent evento) {
                reloj.stop();
                actualizadorTablero.stop();
            }
        });
    }

    private void actualizarFechaHora() {
        etiquetaFechaHora.setText(java.time.LocalDateTime.now().format(
                DateTimeFormatter.ofPattern("EEEE, d 'de' MMMM 'de' uuuu | HH:mm",
                        Locale.forLanguageTag("es-MX"))));
    }

    private JPanel crearPanelConFondo() {
        Image fondo = IconoVentanaUtil.obtenerIconoSeguro("/com/mycompany/ocxrst/IMAGENES/FONDO PRINCIPAL.png").getImage();
        return new JPanel() {
            @Override protected void paintComponent(java.awt.Graphics g) {
                super.paintComponent(g);
                if (fondo != null) g.drawImage(fondo, 0, 0, getWidth(), getHeight(), this);
            }
        };
    }

    private void refrescarTablero() {
        actualizarTableroSilencioso();
        mostrarVista("INICIO", null);
    }

    private void actualizarTableroSilencioso() {
        if (cuerpoTablero == null || !isDisplayable()) return;
        reconstruirTablero();
    }

    private void mostrarPaginaConstruccion() {
        mostrarVista("CONSTRUCCION", null);
    }

    private void actualizarEncabezado(String vista) {
        if (etiquetaEncabezado == null) {
            return;
        }
        String titulo;
        String descripcion;
        switch (vista) {
            case "MESATRABAJO":
                titulo = "Mesa de trabajo";
                descripcion = "Tareas pendientes, actividades prioritarias y seguimiento operativo.";
                break;
            case "ORDENES":
                titulo = "Órdenes de compra";
                descripcion = "Creación, consulta y seguimiento de órdenes de compra.";
                break;
            case "PROVEEDORES":
                titulo = "Proveedores";
                descripcion = "Administración del directorio, condiciones comerciales e historial.";
                break;
            case "INVENTARIOS":
                titulo = "Inventarios";
                descripcion = "Existencias, entradas, salidas y movimientos de productos.";
                break;
            case "USUARIOS":
                titulo = "Usuarios";
                descripcion = "Gestión de usuarios, permisos y accesos.";
                break;
            case "REPORTES":
            case "CONSTRUCCION":
                titulo = "Reportes";
                descripcion = "Información operativa, estadísticas y resultados del sistema.";
                break;
            case "TERMINOS":
                titulo = "Términos y condiciones";
                descripcion = "Consulta de los términos y condiciones de uso del sistema.";
                break;
            case "INICIO":
            default:
                titulo = "Inicio";
                descripcion = "Resumen general de actividades, pendientes y alertas del sistema.";
                break;
        }
        etiquetaEncabezado.setText("<html><b>" + titulo + "</b><br>"
                + "<font size='3'>" + descripcion + "</font></html>");
    }

    private void abrirMesaTrabajo() {
        mostrarVentanaEnPanel("MESATRABAJO", new MESATRABAJO());
    }

    private void confirmarSalida() {
        Object[] opciones = {"Sí", "No"};
        javax.swing.JOptionPane panel = new javax.swing.JOptionPane(
                "¿Desea salir del programa?",
                javax.swing.JOptionPane.QUESTION_MESSAGE,
                javax.swing.JOptionPane.DEFAULT_OPTION,
                null,
                opciones,
                opciones[1]);
        aplicarFondoBlanco(panel);
        javax.swing.JDialog dialogo = panel.createDialog(this, "Confirmar salida");
        dialogo.getContentPane().setBackground(Color.WHITE);
        dialogo.setVisible(true);
        if (opciones[0].equals(panel.getValue())) {
            System.exit(0);
        }
    }

    private void aplicarFondoBlanco(java.awt.Component componente) {
        if (componente instanceof javax.swing.JPanel || componente instanceof javax.swing.JOptionPane) {
            componente.setBackground(Color.WHITE);
        }
        if (componente instanceof java.awt.Container contenedor) {
            for (java.awt.Component hijo : contenedor.getComponents()) {
                aplicarFondoBlanco(hijo);
            }
        }
    }

    private void abrirOrdenes() {
        mostrarVentanaEnPanel("ORDENES", new ORDENCOMPRA());
    }

    private void abrirProveedores() {
        mostrarVentanaEnPanel("PROVEEDORES", new PROVEEDORES());
    }

    private void abrirInventarios() {
        mostrarVentanaEnPanel("INVENTARIOS", new INVENTARIOS());
    }

    private void abrirUsuarios() {
        mostrarVentanaEnPanel("USUARIOS", new USUARIOS());
    }

    private void abrirTerminos() {
        mostrarVentanaEnPanel("TERMINOS", new TERMINOS());
    }

    private void mostrarVentanaEnPanel(String nombre, javax.swing.JFrame ventana) {
        java.awt.Container contenidoVentana = ventana.getContentPane();
        ventana.setContentPane(new JPanel());
        contenidoVentana.setPreferredSize(null);
        contenidoVentana.setMinimumSize(new Dimension(0, 0));
        contenidoVentana.setMaximumSize(new Dimension(Integer.MAX_VALUE, Integer.MAX_VALUE));
        panelVistas.add(contenidoVentana, nombre);
        mostrarVista(nombre, ventana);
    }

    private void mostrarVista(String nombre, javax.swing.JFrame ventana) {
        actualizarEncabezado(nombre);
        ((CardLayout) panelVistas.getLayout()).show(panelVistas, nombre);
        if (ventana != null) {
            ventana.setVisible(false);
            ventana.dispose();
        }
        panelVistas.revalidate();
        panelVistas.repaint();
    }

    private void cerrarSesion() {
        INICIARSESION.usuarioActual = "";
        INICIARSESION.nombreUsuarioActual = "";
        reiniciarAlertaEntregasSesion();
        dispose();
        new INICIARSESION().setVisible(true);
    }

    private void mostrarAlertaEntregasUnaVezPorSesion() {
        if (!alertaEntregasMostradaEnSesion) {
            alertaEntregasMostradaEnSesion = true;
            java.awt.EventQueue.invokeLater(() -> PROVEEDORES.mostrarAlertaEntregasPendientes(this));
        }
    }

    public static void reiniciarAlertaEntregasSesion() { alertaEntregasMostradaEnSesion = false; }

    public static void actualizarAlarmasEnTiempoReal() {
        java.awt.EventQueue.invokeLater(() -> {
            for (java.awt.Window ventana : java.awt.Window.getWindows()) {
                if (ventana instanceof PRINCIPAL principal && principal.isDisplayable()) {
                    principal.actualizarBotonNotificaciones();
                    principal.actualizarTableroSilencioso();
                }
            }
        });
    }

    public static void main(String[] args) {
        java.awt.EventQueue.invokeLater(() -> new PRINCIPAL().setVisible(true));
    }
}