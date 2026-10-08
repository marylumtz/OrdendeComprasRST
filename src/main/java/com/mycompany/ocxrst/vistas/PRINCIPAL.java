package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.EstiloAplicacion;
import com.mycompany.ocxrst.IconoVentanaUtil;
import com.mycompany.ocxrst.config.AppConfig;
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
import java.awt.Insets;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.time.LocalDate;
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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PRINCIPAL extends javax.swing.JFrame {

    private static boolean alertaEntregasMostradaEnSesion;
    private static final DataFormatter FORMATEADOR = new DataFormatter();
    private static final Color AZUL = new Color(30, 57, 130);
    private static final Color AZUL_MENU = new Color(0x01, 0x45, 0xAB);
    private static final Color TEXTO = new Color(38, 52, 88);
    private final ProveedorRepository proveedorRepository = new ProveedorRepository();
    private JButton jButtonNotificaciones;
    private JButton jButtonMenu;
    private JButton botonMenuActivo;
    private JLabel etiquetaFechaHora;
    private javax.swing.JPopupMenu menuSesion;
    private Timer reloj;
    private int totalNotificaciones;
    private JPanel panelVistas;

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
        JPanel capaFecha = new JPanel(new BorderLayout());
        capaFecha.setOpaque(false);
        capaFecha.add(etiquetaFechaHora, BorderLayout.SOUTH);
        setGlassPane(capaFecha);
        capaFecha.setVisible(true);
    }

    private JPanel crearMenuLateral() {
        JPanel lateral = new JPanel(new BorderLayout());
        lateral.setOpaque(false);
        lateral.setBorder(BorderFactory.createEmptyBorder(86, 0, 0, 0));
        JPanel menu = new JPanel();
        menu.setPreferredSize(new Dimension(186, 627));
        menu.setBackground(new Color(255, 255, 255, 150));
        menu.setOpaque(true);
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
        etiquetaFechaHora.setForeground(Color.WHITE);
        etiquetaFechaHora.setOpaque(false);
        etiquetaFechaHora.setFont(new Font("Segoe UI", Font.BOLD, 11));
        etiquetaFechaHora.setPreferredSize(new Dimension(360, 28));
        etiquetaFechaHora.setHorizontalAlignment(SwingConstants.RIGHT);
        etiquetaFechaHora.setBorder(BorderFactory.createEmptyBorder(0, 0, 13, 60));
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
        contenido.setBorder(BorderFactory.createEmptyBorder(14, 8, 44, 8));

        JPanel encabezado = new JPanel(new BorderLayout());
        encabezado.setOpaque(false);
        JLabel bienvenida = new JLabel();
        bienvenida.setForeground(AZUL);
        bienvenida.setFont(new Font("Segoe UI", Font.BOLD, 14));
        bienvenida.setBorder(BorderFactory.createEmptyBorder(60, 0, 0, 0));
        encabezado.add(bienvenida, BorderLayout.CENTER);
        JPanel acciones = new JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 8, 0));
        acciones.setOpaque(false);
        jButtonNotificaciones = new JButton();
        jButtonNotificaciones.setFont(new Font("Segoe UI", Font.BOLD, 12));
        jButtonNotificaciones.setForeground(Color.WHITE);
        jButtonNotificaciones.setFocusPainted(false);
        jButtonNotificaciones.addActionListener(e -> PROVEEDORES.mostrarAlertaEntregasPendientes(this));
        acciones.add(jButtonNotificaciones);
        jButtonMenu = new JButton(nombreUsuario());
        jButtonMenu.setFont(new Font("Segoe UI", Font.BOLD, 12));
        jButtonMenu.setForeground(AZUL);
        jButtonMenu.setBackground(Color.WHITE);
        jButtonMenu.setFocusPainted(false);
        jButtonMenu.addActionListener(e -> menuSesion.show(jButtonMenu, 0, jButtonMenu.getHeight()));
        acciones.add(jButtonMenu);
        JButton botonCerrar = new JButton("×");
        botonCerrar.setFont(new Font("Segoe UI", Font.BOLD, 20));
        botonCerrar.setForeground(Color.WHITE);
        botonCerrar.setContentAreaFilled(false);
        botonCerrar.setBorder(BorderFactory.createEmptyBorder(0, 8, 0, 2));
        botonCerrar.setFocusPainted(false);
        botonCerrar.setToolTipText("Cerrar");
        botonCerrar.addActionListener(e -> confirmarSalida());
        acciones.add(botonCerrar);
        encabezado.add(acciones, BorderLayout.EAST);
        contenido.add(encabezado, BorderLayout.NORTH);

        JPanel cuerpo = new JPanel();
        cuerpo.setOpaque(true);
        cuerpo.setBackground(Color.WHITE);
        cuerpo.setLayout(new BoxLayout(cuerpo, BoxLayout.Y_AXIS));
        cuerpo.add(crearTarjetasIndicadores());
        cuerpo.add(Box.createVerticalStrut(12));
        JPanel centro = new JPanel(new GridLayout(1, 2, 12, 0));
        centro.setOpaque(false);
        centro.setPreferredSize(new Dimension(0, 281));
        centro.setMaximumSize(new Dimension(Integer.MAX_VALUE, 281));
        centro.add(crearPanelTabla("Mis pendientes", cargarPendientes()));
        centro.add(crearPanelTabla("Actividad reciente", cargarActividad()));
        cuerpo.add(centro);
        cuerpo.add(Box.createVerticalStrut(12));
        JPanel inferior = new JPanel(new GridLayout(1, 2, 12, 0));
        inferior.setOpaque(false);
        inferior.add(crearPanelTabla("Stock bajo", cargarStockBajo()));
        inferior.add(crearPanelAccesos());
        cuerpo.add(inferior);
        JScrollPane scroll = new JScrollPane(cuerpo);
        scroll.setBorder(null);
        scroll.setHorizontalScrollBarPolicy(javax.swing.JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        scroll.setOpaque(false);
        scroll.getViewport().setOpaque(false);
        panelVistas = new JPanel(new CardLayout());
        panelVistas.setOpaque(false);
        panelVistas.add(scroll, "INICIO");
        panelVistas.add(crearPaginaConstruccion(), "CONSTRUCCION");
        contenido.add(panelVistas, BorderLayout.CENTER);
        return contenido;
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
        JPanel tarjetas = new JPanel(new GridLayout(1, 5, 10, 0));
        tarjetas.setOpaque(false);
        int ordenes = PROVEEDORES.contarOrdenesRegistradas();
        int pendientes = PROVEEDORES.contarEntregasPorRevisar();
        int proveedores = PROVEEDORES.contarProveedoresActivos();
        int stockBajo = contarStockBajo();
        int usuarios = contarFilas(AppConfig.usuariosFile());
        tarjetas.add(crearTarjeta(String.valueOf(ordenes), "Órdenes de compra", "Total registradas",
                new Color(224, 234, 255), "orden", new Color(61, 96, 190)));
        tarjetas.add(crearTarjeta(String.valueOf(pendientes), "Pendientes", "Requieren atención",
                new Color(255, 244, 190), "pendiente", new Color(218, 155, 0)));
        tarjetas.add(crearTarjeta(String.valueOf(PROVEEDORES.contarAlertasEntregasPendientes()), "Alertas", "Entregas o pagos",
                new Color(255, 220, 220), "alerta", new Color(225, 76, 85)));
        tarjetas.add(crearTarjeta(String.valueOf(stockBajo), "Stock bajo", "Productos por revisar",
                new Color(222, 246, 220), "inventario", new Color(35, 161, 91)));
        tarjetas.add(crearTarjeta(String.valueOf(proveedores), "Proveedores", usuarios + " usuarios activos",
                new Color(247, 224, 248), "proveedor", new Color(157, 24, 191)));
        Dimension tamanoTarjetas = tarjetas.getPreferredSize();
        tarjetas.setPreferredSize(new Dimension(tamanoTarjetas.width + 19, tamanoTarjetas.height));
        tarjetas.setMaximumSize(new Dimension(Integer.MAX_VALUE, tamanoTarjetas.height));
        return tarjetas;
    }

    private JPanel crearTarjeta(String valor, String titulo, String detalle, Color fondo,
            String tipoIcono, Color colorIcono) {
        JPanel tarjeta = new JPanel();
        tarjeta.setLayout(new BoxLayout(tarjeta, BoxLayout.Y_AXIS));
        tarjeta.setBackground(fondo);
        tarjeta.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(210, 220, 235)),
                BorderFactory.createEmptyBorder(8, 8, 8, 8)));
        JPanel cabecera = new JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.CENTER, 12, 0));
        cabecera.setOpaque(false);
        JLabel icono = new JLabel(new IconoMenu(tipoIcono));
        icono.setForeground(colorIcono);
        JLabel numero = new JLabel(valor, SwingConstants.CENTER);
        numero.setFont(new Font("Segoe UI", Font.BOLD, 25));
        numero.setForeground(AZUL);
        cabecera.add(icono);
        cabecera.add(numero);
        JLabel nombre = new JLabel(titulo, SwingConstants.CENTER);
        nombre.setFont(new Font("Segoe UI", Font.BOLD, 12));
        nombre.setForeground(TEXTO);
        JLabel descripcion = new JLabel(detalle, SwingConstants.CENTER);
        descripcion.setFont(new Font("Segoe UI", Font.PLAIN, 10));
        descripcion.setForeground(new Color(90, 105, 125));
        cabecera.setAlignmentX(Component.CENTER_ALIGNMENT);
        nombre.setAlignmentX(Component.CENTER_ALIGNMENT);
        descripcion.setAlignmentX(Component.CENTER_ALIGNMENT);
        tarjeta.add(cabecera);
        tarjeta.add(nombre);
        tarjeta.add(descripcion);
        return tarjeta;
    }

    private JPanel crearPanelTabla(String titulo, Object[][] datos) {
        JPanel panel = new JPanel(new BorderLayout(0, 6));
        panel.setBackground(new Color(255, 255, 255, 235));
        panel.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(215, 224, 238)),
                BorderFactory.createEmptyBorder(8, 8, 8, 8)));
        JLabel encabezado = new JLabel(titulo);
        encabezado.setFont(new Font("Segoe UI", Font.BOLD, 14));
        encabezado.setForeground(AZUL);
        panel.add(encabezado, BorderLayout.NORTH);
        String[] columnas = "Stock bajo".equals(titulo)
                ? new String[]{"Código", "Producto", "Existencia", "Mínimo"}
                : new String[]{"Orden", "Actividad", "Proveedor", "Fecha"};
        JTable tabla = new JTable(new DefaultTableModel(datos, columnas) {
            @Override public boolean isCellEditable(int fila, int columna) { return false; }
        });
        tabla.setRowHeight(26);
        tabla.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        tabla.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 11));
        tabla.setForeground(TEXTO);
        panel.add(new JScrollPane(tabla), BorderLayout.CENTER);
        return panel;
    }

    private JPanel crearPanelAccesos() {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setBackground(new Color(255, 255, 255, 235));
        panel.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(215, 224, 238)),
                BorderFactory.createEmptyBorder(8, 8, 8, 8)));
        JLabel titulo = new JLabel("Accesos rápidos");
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 14));
        titulo.setForeground(AZUL);
        panel.add(titulo, BorderLayout.NORTH);
        JPanel botones = new JPanel(new GridLayout(2, 2, 8, 8));
        botones.setOpaque(false);
        botones.add(crearAcceso("ÓRDENES DE COMPRA", e -> abrirOrdenes()));
        botones.add(crearAcceso("PROVEEDORES", e -> abrirProveedores()));
        botones.add(crearAcceso("INVENTARIOS", e -> abrirInventarios()));
        botones.add(crearAcceso("USUARIOS", e -> abrirUsuarios()));
        panel.add(botones, BorderLayout.CENTER);
        return panel;
    }

    private JButton crearAcceso(String texto, ActionListener accion) {
        JButton boton = new JButton("<html><b>" + texto + "</b><br><font size='2'>Abrir módulo</font></html>");
        boton.setHorizontalAlignment(SwingConstants.LEFT);
        boton.setForeground(AZUL);
        boton.setBackground(new Color(235, 242, 255));
        boton.setFocusPainted(false);
        boton.setBorder(BorderFactory.createLineBorder(new Color(190, 207, 235)));
        boton.addActionListener(accion);
        return boton;
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
        leerFilas(AppConfig.registroCFile(), (row, formato) -> {
            if (filas.size() < 5 && !leer(row, 0).isBlank()) {
                filas.add(new Object[]{leer(row, 0), "Orden registrada", proveedorRepository.obtenerNombrePorId(leer(row, 4)), leer(row, 1)});
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

    private int contarFilas(File archivo) {
        final int[] total = {0};
        leerFilas(archivo, (row, formato) -> { if (!leer(row, 0).isBlank()) total[0]++; });
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
    }

    private void actualizarFechaHora() {
        etiquetaFechaHora.setText(java.time.LocalDateTime.now().format(
                DateTimeFormatter.ofPattern("EEEE, d 'de' MMMM 'de' uuuu | HH:mm", new Locale("es", "MX"))));
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
        mostrarVista("INICIO", null);
    }

    private void mostrarPaginaConstruccion() {
        mostrarVista("CONSTRUCCION", null);
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
                if (ventana instanceof PRINCIPAL principal && principal.isDisplayable()) principal.actualizarBotonNotificaciones();
            }
        });
    }

    private void mostrarMensaje(String mensaje, int tipo) {
        javax.swing.JOptionPane.showMessageDialog(this, mensaje, "RST", tipo);
    }

    public static void main(String[] args) {
        java.awt.EventQueue.invokeLater(() -> new PRINCIPAL().setVisible(true));
    }
}
