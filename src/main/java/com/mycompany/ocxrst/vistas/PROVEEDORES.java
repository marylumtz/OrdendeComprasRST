/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.IconoVentanaUtil;
import com.mycompany.ocxrst.EstiloAplicacion;
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
        "ORDENCOMPRA", "PROVEEDOR", "FECHA_RECIBIDO", "HORA_RECIBIDO",
        "USUARIO", "NOMBRE_USUARIO", "MODELO", "CANTIDAD", "TIPO_RECEPCION"
    };
    private static final String[] ENCABEZADOS_INCIDENCIAS = {
        "ORDENCOMPRA", "FECHA", "CODIGO_PRODUCTO", "PRODUCTO", "CANTIDAD",
        "TIPO", "DESCRIPCION", "REGISTRADO_POR", "ESTATUS"
    };
    private static final String[] ENCABEZADOS_DEVOLUCIONES = {
        "ORDENCOMPRA", "FECHA", "CODIGO_PRODUCTO", "PRODUCTO", "CANTIDAD",
        "TIPO", "CONDICION", "GUIA", "ACCION_SOLICITADA", "MOTIVO",
        "REGISTRADO_POR", "ESTATUS"
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
        EstiloAplicacion.aplicarTipografiaGlobal();
        initComponents();
        configurarDisenoVisual();
        IconoVentanaUtil.aplicar(this);
        setSize(800, 520);
        setResizable(false);
        setLocationRelativeTo(null);
        asegurarArchivoProveedores();
        cargarCacheProveedores();
        configurarAutocompleteNombre();
        configurarEventos();
    }

    private static final java.awt.Color COLOR_AZUL_OSCURO = new java.awt.Color(28, 82, 132);
    private static final java.awt.Color COLOR_AZUL = new java.awt.Color(31, 112, 184);
    private static final java.awt.Color COLOR_FONDO = new java.awt.Color(237, 245, 253);
    private static final java.awt.Color COLOR_BORDE_CAMPO = new java.awt.Color(198, 210, 225);
    private static final java.awt.Color COLOR_VERDE = new java.awt.Color(0, 130, 60);
    private static final java.awt.Color COLOR_ROJO = new java.awt.Color(190, 0, 0);
    private static final java.awt.Color COLOR_GRIS = new java.awt.Color(90, 100, 110);
    private static final java.awt.Color COLOR_AZUL_CLARO = new java.awt.Color(219, 234, 248);
    private static final java.awt.Color COLOR_AZUL_PALIDO = new java.awt.Color(240, 246, 253);
    private static final java.awt.Color COLOR_TEXTO_GRIS = new java.awt.Color(110, 122, 138);
    private static final java.awt.Color COLOR_NARANJA = new java.awt.Color(196, 118, 0);
    private static final java.awt.Color COLOR_NARANJA_CLARO = new java.awt.Color(255, 240, 219);
    private static final java.awt.Color COLOR_VERDE_CLARO = new java.awt.Color(224, 246, 233);
    private static final java.awt.Color COLOR_ROJO_CLARO = new java.awt.Color(253, 226, 226);
    private static final java.awt.Color COLOR_GRIS_CLARO = new java.awt.Color(236, 239, 243);
    private static final java.awt.Color COLOR_VERDE_EXCEL = new java.awt.Color(33, 140, 87);
    private static final java.awt.Color COLOR_ROJO_PDF = new java.awt.Color(199, 47, 47);

    private void configurarDisenoVisual() {
        jPanel1.removeAll();
        jPanel1.setOpaque(true);
        jPanel1.setBackground(COLOR_FONDO);
        jPanel1.setBorder(javax.swing.BorderFactory.createEmptyBorder(18, 20, 18, 20));
        jPanel1.setLayout(new java.awt.BorderLayout(0, 14));

        // ----- Encabezado -----
        javax.swing.JPanel encabezado = new javax.swing.JPanel(new java.awt.BorderLayout(12, 0));
        encabezado.setOpaque(false);

        javax.swing.JPanel titulo = new javax.swing.JPanel(new java.awt.BorderLayout(10, 0));
        titulo.setOpaque(false);
        javax.swing.JLabel iconoGrupo = new javax.swing.JLabel(iconoGrupo(COLOR_AZUL_OSCURO, 34));
        titulo.add(iconoGrupo, java.awt.BorderLayout.WEST);

        javax.swing.JPanel textoTitulo = new javax.swing.JPanel();
        textoTitulo.setOpaque(false);
        textoTitulo.setLayout(new javax.swing.BoxLayout(textoTitulo, javax.swing.BoxLayout.Y_AXIS));
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        jLabel1.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 26));
        jLabel1.setForeground(COLOR_AZUL_OSCURO);
        jLabel1.setText("PROVEEDORES");
        javax.swing.JLabel subtitulo = new javax.swing.JLabel("Administra la información de tus proveedores");
        subtitulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        subtitulo.setForeground(COLOR_AZUL_OSCURO);
        textoTitulo.add(jLabel1);
        textoTitulo.add(subtitulo);
        titulo.add(textoTitulo, java.awt.BorderLayout.CENTER);
        encabezado.add(titulo, java.awt.BorderLayout.WEST);

        javax.swing.JLabel lema = new javax.swing.JLabel(
                "<html><div style='text-align:right'>SOLUCIONES HOY<br>CONEXIONES MAÑANA</div></html>");
        lema.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 9));
        lema.setForeground(new java.awt.Color(133, 157, 185));
        encabezado.add(lema, java.awt.BorderLayout.EAST);
        jPanel1.add(encabezado, java.awt.BorderLayout.NORTH);

        // ----- Tarjeta principal -----
        RoundedPanel tarjeta = new RoundedPanel(14, java.awt.Color.WHITE);
        tarjeta.setBorder(javax.swing.BorderFactory.createEmptyBorder(16, 22, 16, 22));
        tarjeta.setLayout(new java.awt.BorderLayout(0, 14));

        javax.swing.JPanel barra = new javax.swing.JPanel(new java.awt.BorderLayout());
        barra.setOpaque(false);
        prepararBoton(jButton6, "ATRÁS", iconoChevronIzquierda(java.awt.Color.WHITE), COLOR_AZUL, java.awt.Color.WHITE, null);
        barra.add(jButton6, java.awt.BorderLayout.WEST);

        javax.swing.JLabel registro = new javax.swing.JLabel("REGISTRO DE PROVEEDOR",
                iconoEdificio(COLOR_AZUL_OSCURO, 20), javax.swing.SwingConstants.LEFT);
        registro.setIconTextGap(8);
        registro.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 12));
        registro.setForeground(COLOR_AZUL_OSCURO);
        javax.swing.JPanel contenedorRegistro = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 0, 0));
        contenedorRegistro.setOpaque(false);
        contenedorRegistro.add(registro);
        barra.add(contenedorRegistro, java.awt.BorderLayout.EAST);
        tarjeta.add(barra, java.awt.BorderLayout.NORTH);

        javax.swing.JPanel formulario = new javax.swing.JPanel(new java.awt.GridBagLayout());
        formulario.setOpaque(false);
        agregarCampo(formulario, 0, iconoId(COLOR_AZUL_OSCURO), "ID:", jTextFieldId,
                "Ingrese el ID del proveedor", true);
        agregarCampo(formulario, 1, iconoEdificio(COLOR_AZUL_OSCURO, 18), "NOMBRE O RAZÓN SOCIAL:", jTextField1,
                "Nombre o razón social", false);
        agregarCampo(formulario, 2, iconoDocumento(COLOR_AZUL_OSCURO), "RFC:", jTextField2,
                "RFC (Ej. XAXX010101000)", false);
        agregarCampo(formulario, 3, iconoTelefono(COLOR_AZUL_OSCURO), "TELÉFONO:", jTextField3,
                "Número telefónico", false);
        agregarCampo(formulario, 4, iconoCorreo(COLOR_AZUL_OSCURO), "CORREO:", jTextField4,
                "correo@ejemplo.com", false);
        agregarCampo(formulario, 5, iconoUbicacion(COLOR_AZUL_OSCURO), "DIRECCIÓN:", jTextField5,
                "Dirección completa", false);
        agregarCampo(formulario, 6, iconoContacto(COLOR_AZUL_OSCURO), "CONTACTO:", jTextField6,
                "Nombre del contacto", false);
        tarjeta.add(formulario, java.awt.BorderLayout.CENTER);

        javax.swing.JPanel acciones = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 10, 0));
        acciones.setOpaque(false);
        prepararBoton(jButton4, "HISTORIAL DE COMPRAS", iconoHistorial(COLOR_AZUL), java.awt.Color.WHITE,
                COLOR_AZUL, COLOR_AZUL);
        prepararBoton(jButton1, "CREAR", iconoMas(java.awt.Color.WHITE), COLOR_AZUL, java.awt.Color.WHITE, null);
        prepararBoton(jButton2, "EDITAR", iconoEditar(java.awt.Color.WHITE), COLOR_AZUL, java.awt.Color.WHITE, null);
        acciones.add(jButton4);
        acciones.add(jButton1);
        acciones.add(jButton2);
        acciones.add(jButton3);
        tarjeta.add(acciones, java.awt.BorderLayout.SOUTH);
        jPanel1.add(tarjeta, java.awt.BorderLayout.CENTER);

        restablecerIndicadorEstado();

        getContentPane().setBackground(COLOR_FONDO);
        getContentPane().removeAll();
        getContentPane().setLayout(new java.awt.BorderLayout());
        getContentPane().add(jPanel1, java.awt.BorderLayout.CENTER);
        getContentPane().revalidate();
        getContentPane().repaint();
    }

    private void agregarCampo(javax.swing.JPanel formulario, int fila, javax.swing.Icon icono, String etiqueta,
            javax.swing.JTextField campo, String textoAyuda, boolean soloLectura) {
        javax.swing.JLabel iconoLabel = new javax.swing.JLabel(icono);

        javax.swing.JPanel contenedorEtiqueta = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 3, 0));
        contenedorEtiqueta.setOpaque(false);
        contenedorEtiqueta.setPreferredSize(new java.awt.Dimension(215, 20));
        javax.swing.JLabel etiquetaLabel = new javax.swing.JLabel(etiqueta);
        etiquetaLabel.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        etiquetaLabel.setForeground(new java.awt.Color(34, 76, 119));
        javax.swing.JLabel asteriscoLabel = new javax.swing.JLabel("*");
        asteriscoLabel.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        asteriscoLabel.setForeground(new java.awt.Color(217, 43, 43));
        contenedorEtiqueta.add(etiquetaLabel);
        contenedorEtiqueta.add(asteriscoLabel);

        campo.setToolTipText(textoAyuda);
        campo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        campo.setPreferredSize(new java.awt.Dimension(120, 30));
        campo.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO),
                javax.swing.BorderFactory.createEmptyBorder(4, 10, 4, 10)));
        campo.setBackground(soloLectura ? new java.awt.Color(247, 250, 253) : java.awt.Color.WHITE);
        campo.setEditable(!soloLectura);

        java.awt.GridBagConstraints c = new java.awt.GridBagConstraints();
        c.gridy = fila;
        c.insets = new java.awt.Insets(4, 0, 4, 10);
        c.anchor = java.awt.GridBagConstraints.CENTER;
        c.gridx = 0;
        c.weightx = 0;
        formulario.add(iconoLabel, c);
        c.gridx = 1;
        c.anchor = java.awt.GridBagConstraints.LINE_END;
        c.insets = new java.awt.Insets(4, 0, 4, 14);
        formulario.add(contenedorEtiqueta, c);
        c.gridx = 2;
        c.weightx = 1;
        c.fill = java.awt.GridBagConstraints.HORIZONTAL;
        c.insets = new java.awt.Insets(4, 0, 4, 0);
        formulario.add(campo, c);
    }

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
    }

    private static java.awt.Color mezclarColor(java.awt.Color base, java.awt.Color destino, float proporcion) {
        int r = base.getRed() + Math.round((destino.getRed() - base.getRed()) * proporcion);
        int g = base.getGreen() + Math.round((destino.getGreen() - base.getGreen()) * proporcion);
        int b = base.getBlue() + Math.round((destino.getBlue() - base.getBlue()) * proporcion);
        return new java.awt.Color(
                Math.max(0, Math.min(255, r)),
                Math.max(0, Math.min(255, g)),
                Math.max(0, Math.min(255, b)));
    }

    /**
     * ButtonUI plano que sustituye el bisel nativo de Windows (que generaba un
     * doble contorno al combinarse con nuestro borde personalizado) por un
     * relleno redondeado propio con retroalimentación sutil de hover/press.
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

    /**
     * Borde redondeado de una sola línea. Si el color es {@code null} no se
     * dibuja nada, pero conserva el mismo espaciado interno para que todos
     * los botones mantengan la misma altura.
     */
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

    // ===================== Iconos vectoriales (Java2D) =====================
    // Se dibujan con Graphics2D en vez de usar caracteres Unicode como "iconos",
    // ya que algunos glifos no están disponibles en todas las fuentes/instalaciones
    // de Windows y se mostraban como cuadros ("tofu").

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
                for (int col = 0; col < 2; col++) {
                    g2.drawRect(4 + col * 5, 3 + fila * 4, 2, 2);
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

    private static javax.swing.Icon iconoTelefono(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            java.awt.geom.AffineTransform original = g2.getTransform();
            g2.rotate(Math.toRadians(45), t / 2.0, t / 2.0);
            g2.drawRoundRect(t / 2 - 6, t / 2 - 2, 12, 4, 3, 3);
            g2.fillOval(t / 2 - 7, t / 2 - 2, 4, 4);
            g2.fillOval(t / 2 + 3, t / 2 - 2, 4, 4);
            g2.setTransform(original);
        });
    }

    private static javax.swing.Icon iconoCorreo(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawRoundRect(1, 3, t - 3, t - 7, 2, 2);
            g2.drawLine(1, 3, t / 2, t / 2);
            g2.drawLine(t - 2, 3, t / 2, t / 2);
        });
    }

    private static javax.swing.Icon iconoUbicacion(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawOval(t / 2 - 5, 1, 10, 10);
            g2.fillOval(t / 2 - 2, 4, 4, 4);
            java.awt.geom.GeneralPath punta = new java.awt.geom.GeneralPath();
            punta.moveTo(t / 2 - 4, 8);
            punta.lineTo(t / 2.0, t - 1);
            punta.lineTo(t / 2 + 4, 8);
            g2.draw(punta);
        });
    }

    private static javax.swing.Icon iconoContacto(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawOval(t / 2 - 3, 1, 6, 6);
            java.awt.geom.Arc2D cuerpo = new java.awt.geom.Arc2D.Double(2, 8, t - 4, 11, 0, 180, java.awt.geom.Arc2D.OPEN);
            g2.draw(cuerpo);
        });
    }

    private static javax.swing.Icon iconoGrupo(java.awt.Color color, int tamano) {
        return new VectorIcon(tamano, color, (g2, t) -> {
            int r = t / 3;
            g2.fillOval(1, t / 2 - r / 2, r, r);
            g2.fillOval(t / 2, t / 2 - r / 2, r, r);
            java.awt.geom.Arc2D cuerpoIzq = new java.awt.geom.Arc2D.Double(-t / 6.0, t / 2.0 + r / 2.0 - 2,
                    t * 0.7, t * 0.6, 0, 180, java.awt.geom.Arc2D.OPEN);
            java.awt.geom.Arc2D cuerpoDer = new java.awt.geom.Arc2D.Double(t / 3.0, t / 2.0 + r / 2.0 - 2,
                    t * 0.7, t * 0.6, 0, 180, java.awt.geom.Arc2D.OPEN);
            g2.draw(cuerpoIzq);
            g2.draw(cuerpoDer);
        });
    }

    private static javax.swing.Icon iconoChevronIzquierda(java.awt.Color color) {
        return new VectorIcon(10, color, (g2, t) -> {
            g2.drawLine(t - 3, 1, 2, t / 2);
            g2.drawLine(2, t / 2, t - 3, t - 1);
        });
    }

    private static javax.swing.Icon iconoHistorial(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawOval(1, 1, t - 2, t - 2);
            g2.drawLine(t / 2, t / 2, t / 2, 4);
            g2.drawLine(t / 2, t / 2, t - 4, t / 2);
        });
    }

    private static javax.swing.Icon iconoMas(java.awt.Color color) {
        return new VectorIcon(14, color, (g2, t) -> {
            g2.drawLine(t / 2, 1, t / 2, t - 1);
            g2.drawLine(1, t / 2, t - 1, t / 2);
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
    }

    private static javax.swing.Icon iconoPower(java.awt.Color color) {
        return new VectorIcon(15, color, (g2, t) -> {
            g2.drawArc(1, 3, t - 2, t - 4, -60, 300);
            g2.drawLine(t / 2, 1, t / 2, t / 2);
        });
    }

    private static javax.swing.Icon iconoCarrito(java.awt.Color color, int tamano) {
        return new VectorIcon(tamano, color, (g2, t) -> {
            java.awt.geom.GeneralPath canasta = new java.awt.geom.GeneralPath();
            canasta.moveTo(1, 2);
            canasta.lineTo(4, 2);
            canasta.lineTo(t * 0.32, t * 0.62);
            canasta.lineTo(t - 3, t * 0.62);
            canasta.lineTo(t - 1, t * 0.28);
            canasta.lineTo(5, t * 0.28);
            g2.draw(canasta);
            g2.fillOval((int) (t * 0.34), (int) (t * 0.74), 3, 3);
            g2.fillOval((int) (t * 0.72), (int) (t * 0.74), 3, 3);
        });
    }

    private static javax.swing.Icon iconoDolar(java.awt.Color color, int tamano) {
        return new VectorIcon(tamano, color, (g2, t) -> {
            g2.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, (int) Math.round(t * 0.85)));
            java.awt.font.FontRenderContext contexto = g2.getFontRenderContext();
            java.awt.geom.Rectangle2D limites = g2.getFont().getStringBounds("$", contexto);
            float x = (float) ((t - limites.getWidth()) / 2.0);
            float y = (float) ((t - limites.getHeight()) / 2.0 - limites.getY());
            g2.drawString("$", x, y);
        });
    }

    private static javax.swing.Icon iconoTarjeta(java.awt.Color color, int tamano) {
        return new VectorIcon(tamano, color, (g2, t) -> {
            g2.drawRoundRect(1, (int) (t * 0.2), t - 2, (int) (t * 0.6), 2, 2);
            g2.drawLine(1, (int) (t * 0.4), t - 1, (int) (t * 0.4));
            g2.drawLine(3, (int) (t * 0.62), (int) (t * 0.4), (int) (t * 0.62));
        });
    }

    private static javax.swing.Icon iconoBuscar(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawOval(1, 1, (int) (t * 0.65), (int) (t * 0.65));
            g2.drawLine((int) (t * 0.6), (int) (t * 0.6), t - 1, t - 1);
        });
    }

    private static javax.swing.Icon iconoCalendario(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawRoundRect(1, 3, t - 2, t - 4, 2, 2);
            g2.drawLine(1, (int) (t * 0.45), t - 1, (int) (t * 0.45));
            g2.drawLine((int) (t * 0.3), 1, (int) (t * 0.3), 5);
            g2.drawLine((int) (t * 0.7), 1, (int) (t * 0.7), 5);
        });
    }

    private static javax.swing.Icon iconoFiltro(java.awt.Color color) {
        return new VectorIcon(14, color, (g2, t) -> {
            java.awt.geom.GeneralPath embudo = new java.awt.geom.GeneralPath();
            embudo.moveTo(1, 1);
            embudo.lineTo(t - 1, 1);
            embudo.lineTo(t * 0.62f, t * 0.55f);
            embudo.lineTo(t * 0.62f, t - 1);
            embudo.lineTo(t * 0.38f, t * 0.8f);
            embudo.lineTo(t * 0.38f, t * 0.55f);
            embudo.closePath();
            g2.draw(embudo);
        });
    }

    private static javax.swing.Icon iconoDocumentoAdjunto(java.awt.Color color) {
        return new VectorIcon(15, color, (g2, t) -> {
            g2.drawRoundRect(3, 1, t - 7, t - 2, 2, 2);
            g2.drawLine(5, 5, t - 6, 5);
            g2.drawLine(5, 8, t - 6, 8);
            g2.drawLine(5, 11, (int) (t * 0.55), 11);
        });
    }

    private static javax.swing.Icon iconoCerrarX(java.awt.Color color) {
        return new VectorIcon(12, color, (g2, t) -> {
            g2.drawLine(1, 1, t - 1, t - 1);
            g2.drawLine(t - 1, 1, 1, t - 1);
        });
    }

    private static javax.swing.Icon iconoChevronDerecha(java.awt.Color color) {
        return new VectorIcon(10, color, (g2, t) -> {
            g2.drawLine(2, 1, t - 3, t / 2);
            g2.drawLine(t - 3, t / 2, 2, t - 1);
        });
    }

    private static javax.swing.Icon iconoDobleChevronIzquierda(java.awt.Color color) {
        return new VectorIcon(13, color, (g2, t) -> {
            g2.drawLine(t / 2, 1, 2, t / 2);
            g2.drawLine(2, t / 2, t / 2, t - 1);
            g2.drawLine(t - 2, 1, t / 2 - 2, t / 2);
            g2.drawLine(t / 2 - 2, t / 2, t - 2, t - 1);
        });
    }

    private static javax.swing.Icon iconoDobleChevronDerecha(java.awt.Color color) {
        return new VectorIcon(13, color, (g2, t) -> {
            g2.drawLine(t / 2 - 2, 1, t - 2, t / 2);
            g2.drawLine(t - 2, t / 2, t / 2 - 2, t - 1);
            g2.drawLine(2, 1, t / 2, t / 2);
            g2.drawLine(t / 2, t / 2, 2, t - 1);
        });
    }

    private static javax.swing.Icon iconoMonedas(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawOval(1, (int) (t * 0.28), (int) (t * 0.62), (int) (t * 0.62));
            g2.drawOval((int) (t * 0.36), 1, (int) (t * 0.62), (int) (t * 0.62));
        });
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

    private static final class ResultadoHistorial {
        final java.util.List<Object[]> filas;
        final double sumaTotal;
        final double sumaAbonado;
        final String simboloMoneda;

        ResultadoHistorial(java.util.List<Object[]> filas, double sumaTotal, double sumaAbonado, String simboloMoneda) {
            this.filas = filas;
            this.sumaTotal = sumaTotal;
            this.sumaAbonado = sumaAbonado;
            this.simboloMoneda = simboloMoneda;
        }
    }

    private static final class TarjetaResumen {
        final javax.swing.JPanel panel;
        final javax.swing.JLabel etiquetaValor;

        TarjetaResumen(javax.swing.JPanel panel, javax.swing.JLabel etiquetaValor) {
            this.panel = panel;
            this.etiquetaValor = etiquetaValor;
        }
    }

    /**
     * Lee el archivo de ordenes de compra y arma las filas del historial
     * (filtradas por proveedor si se indica un ID), junto con los totales
     * acumulados. Se separa de mostrarHistorialCompras() para poder recargar
     * los datos "en caliente" cada vez que se aplica un filtro o se registra
     * un cambio (abono, recepcion, etc.), evitando que la tabla muestre
     * informacion desactualizada.
     */
    private ResultadoHistorial cargarFilasHistorial(String idProveedor) {
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
                    java.util.Date fechaPago = leerFechaCelda(row, 12);
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
                    String fechaPagoTexto = formatearFechaCelda(fechaPago);
                    String fechaEntregaTexto = formatearFechaCelda(fechaEntrega);
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

                    // Si la mercancia ya fue recibida en su totalidad no debe quedar nada
                    // pendiente por recibir, sin importar lo que reporte el detalle (que
                    // puede quedar desactualizado despues de una recepcion parcial previa).
                    double cantidadPendiente = "SI".equalsIgnoreCase(estadoRecepcion)
                            ? 0.0
                            : productosPendientesPorOrden.getOrDefault(noOrden.toUpperCase(java.util.Locale.ROOT), 0.0);
                    String productosPendientes = formatearCantidadInventario(cantidadPendiente);
                    String nombreProveedor = obtenerNombreProveedorPorId(idFila);
                    filas.add(new Object[]{idFila, nombreProveedor, noOrden, fechaTexto, fechaPagoTexto,
                        fechaEntregaTexto, cotizacion, solicitante, total, abonado, pendiente, diasFaltantesTexto,
                        productosPendientes, estatus});
                }
            } catch (Exception ex) {
                mostrarError("No se pudo leer el historial de ordenes de compra.", ex);
                return null;
            }
        }
        return new ResultadoHistorial(filas, sumaTotal, sumaAbonado, simboloMoneda);
    }//cargarFilasHistorial

    private void mostrarHistorialCompras() {
        String idProveedor = jTextFieldId.getText().trim();
        ResultadoHistorial resultadoInicial = cargarFilasHistorial(idProveedor);
        if (resultadoInicial == null) {
            return;
        }

        String[] columnas = {
            "ID PROVEEDOR", "NOMBRE PROVEEDOR", "NO. ORDEN", "FECHA DE CREACI\u00d3N",
            "FECHA DE PAGO PROMETIDA", "FECHA DE RECEPCI\u00d3N PROMETIDA", "COTIZACI\u00d3N", "SOLICITANTE",
            "TOTAL", "ABONADO", "PEND. X PAGAR", "DIAS FAL_RECIBIR", "PROD. X RECIBIR", "ESTATUS"
        };
        javax.swing.table.DefaultTableModel modelo = new javax.swing.table.DefaultTableModel(columnas, 0) {
            @Override
            public boolean isCellEditable(int fila, int columna) {
                return false;
            }
        };

        javax.swing.JTable tabla = new javax.swing.JTable(modelo) {
            @Override
            public java.awt.Component prepareRenderer(javax.swing.table.TableCellRenderer renderer,
                    int fila, int columna) {
                java.awt.Component componente = super.prepareRenderer(renderer, fila, columna);
                int columnaModelo = convertColumnIndexToModel(columna);
                if (!isCellSelected(fila, columna) && columnaModelo != 13) {
                    int filaModelo = convertRowIndexToModel(fila);
                    String estatus = String.valueOf(getModel().getValueAt(filaModelo, 13)).trim();
                    componente.setBackground(getBackground());
                    componente.setForeground(getForeground());
                    try {
                        long dias = Long.parseLong(String.valueOf(getModel().getValueAt(filaModelo, 11)).trim());
                        boolean columnaPendiente = columnaModelo == 10 || columnaModelo == 12;
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
        tabla.setRowHeight(48);
        tabla.setShowGrid(false);
        tabla.setIntercellSpacing(new java.awt.Dimension(0, 0));
        tabla.setSelectionBackground(new java.awt.Color(224, 236, 250));
        tabla.setSelectionForeground(COLOR_AZUL_OSCURO);
        tabla.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        tabla.setDefaultRenderer(Object.class, new javax.swing.table.DefaultTableCellRenderer() {
            private final javax.swing.JTextArea texto = new javax.swing.JTextArea();

            {
                texto.setLineWrap(true);
                texto.setWrapStyleWord(true);
                texto.setOpaque(true);
                texto.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
                texto.setBorder(javax.swing.BorderFactory.createEmptyBorder(4, 6, 4, 6));
            }

            @Override
            public java.awt.Component getTableCellRendererComponent(javax.swing.JTable tabla, Object valor,
                    boolean seleccionado, boolean foco, int fila, int columna) {
                texto.setText(valor == null ? "" : String.valueOf(valor));
                texto.setBackground(seleccionado ? tabla.getSelectionBackground() : tabla.getBackground());
                texto.setForeground(seleccionado ? tabla.getSelectionForeground() : tabla.getForeground());
                return texto;
            }
        });
        tabla.setFillsViewportHeight(true);
        tabla.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        tabla.getTableHeader().setReorderingAllowed(false);
        configurarEncabezadoHistorial(tabla);
        configurarColumnaEstatusHistorial(tabla);
        tabla.removeColumn(tabla.getColumn("ID PROVEEDOR"));
        tabla.removeColumn(tabla.getColumn("DIAS FAL_RECIBIR"));
        java.awt.Rectangle areaUtil = java.awt.GraphicsEnvironment.getLocalGraphicsEnvironment()
            .getMaximumWindowBounds();

        // ----- Encabezado (icono + titulo + subtitulo) -----
        javax.swing.JPanel panelHeader = new javax.swing.JPanel(new java.awt.BorderLayout(14, 0));
        panelHeader.setOpaque(false);

        javax.swing.JPanel panelTituloIzquierda = new javax.swing.JPanel(new java.awt.BorderLayout(14, 0));
        panelTituloIzquierda.setOpaque(false);
        panelTituloIzquierda.add(crearInsigniaCuadrada(iconoDocumentoAdjunto(java.awt.Color.WHITE), COLOR_AZUL, 46, 12),
                java.awt.BorderLayout.WEST);

        javax.swing.JPanel textoTitulo = new javax.swing.JPanel();
        textoTitulo.setOpaque(false);
        textoTitulo.setLayout(new javax.swing.BoxLayout(textoTitulo, javax.swing.BoxLayout.Y_AXIS));
        javax.swing.JLabel tituloHistorial = new javax.swing.JLabel(idProveedor.isEmpty()
                ? "Historial de \u00d3rdenes de Compra"
                : "Historial de \u00d3rdenes de Compra - " + idProveedor);
        tituloHistorial.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 22));
        tituloHistorial.setForeground(COLOR_AZUL_OSCURO);
        javax.swing.JLabel subtituloHistorial = new javax.swing.JLabel(idProveedor.isEmpty()
                ? "Consulta el historial de \u00f3rdenes de compra de todos los proveedores"
                : "Consulta el historial de \u00f3rdenes de compra de este proveedor");
        subtituloHistorial.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        subtituloHistorial.setForeground(COLOR_TEXTO_GRIS);
        textoTitulo.add(tituloHistorial);
        textoTitulo.add(subtituloHistorial);
        panelTituloIzquierda.add(textoTitulo, java.awt.BorderLayout.CENTER);
        panelHeader.add(panelTituloIzquierda, java.awt.BorderLayout.WEST);

        // ----- Tarjetas de resumen -----
        TarjetaResumen tarjetaCompras = crearTarjetaResumen(
                idProveedor.isEmpty() ? "Total de compras" : "Total de compras del proveedor",
                String.valueOf(resultadoInicial.filas.size()), iconoCarrito(COLOR_AZUL, 20), COLOR_AZUL_CLARO);
        TarjetaResumen tarjetaTotalOrdenes = crearTarjetaResumen("Total de todas las \u00f3rdenes",
                String.format("%s %,.2f", resultadoInicial.simboloMoneda, resultadoInicial.sumaTotal),
                iconoDolar(COLOR_AZUL, 18), COLOR_AZUL_CLARO);
        TarjetaResumen tarjetaAbonado = crearTarjetaResumen("Total abonado",
                String.format("%s %,.2f", resultadoInicial.simboloMoneda, resultadoInicial.sumaAbonado),
                iconoTarjeta(COLOR_AZUL, 20), COLOR_AZUL_CLARO);

        javax.swing.JPanel panelTarjetas = new javax.swing.JPanel(new java.awt.GridLayout(1, 3, 14, 0));
        panelTarjetas.setOpaque(false);
        panelTarjetas.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 0, 14, 0));
        panelTarjetas.add(tarjetaCompras.panel);
        panelTarjetas.add(tarjetaTotalOrdenes.panel);
        panelTarjetas.add(tarjetaAbonado.panel);

        // ----- Buscador y filtros -----
        javax.swing.JTextField campoBuscar = new javax.swing.JTextField();
        campoBuscar.setBorder(javax.swing.BorderFactory.createEmptyBorder(8, 4, 8, 10));
        campoBuscar.setOpaque(false);
        campoBuscar.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        String placeholderBusqueda = "Buscar por ID/nombre proveedor, No. orden, cotizaci\u00f3n o solicitante...";
        configurarPlaceholder(campoBuscar, placeholderBusqueda);

        javax.swing.JPanel contenedorBuscar = new javax.swing.JPanel(new java.awt.BorderLayout(6, 0));
        contenedorBuscar.setOpaque(false);
        contenedorBuscar.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                new BordeRedondeado(9, COLOR_BORDE_CAMPO, 1.2f),
                javax.swing.BorderFactory.createEmptyBorder(2, 12, 2, 6)));
        contenedorBuscar.add(new javax.swing.JLabel(iconoBuscar(COLOR_TEXTO_GRIS)), java.awt.BorderLayout.WEST);
        contenedorBuscar.add(campoBuscar, java.awt.BorderLayout.CENTER);

        javax.swing.JComboBox<String> comboFecha = new javax.swing.JComboBox<>(new String[]{
            "Todas las fechas", "Hoy", "\u00daltimos 7 d\u00edas", "Este mes", "Este a\u00f1o"
        });
        comboFecha.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        comboFecha.setPreferredSize(new java.awt.Dimension(150, 34));

        javax.swing.JPanel contenedorFecha = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.LEFT, 6, 0));
        contenedorFecha.setOpaque(false);
        contenedorFecha.add(new javax.swing.JLabel(iconoCalendario(COLOR_TEXTO_GRIS)));
        contenedorFecha.add(comboFecha);

        javax.swing.JPanel panelFiltrosDerecha = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 10, 0));
        panelFiltrosDerecha.setOpaque(false);
        panelFiltrosDerecha.add(contenedorFecha);

        javax.swing.JPanel panelBusquedaFiltros = new javax.swing.JPanel(new java.awt.BorderLayout(12, 0));
        panelBusquedaFiltros.setOpaque(false);
        panelBusquedaFiltros.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 0, 12, 0));
        panelBusquedaFiltros.add(contenedorBuscar, java.awt.BorderLayout.CENTER);
        panelBusquedaFiltros.add(panelFiltrosDerecha, java.awt.BorderLayout.EAST);

        // ----- Tabla -----
        javax.swing.JScrollPane scroll = new javax.swing.JScrollPane(tabla);
        scroll.setBorder(javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO));
        scroll.getViewport().setBackground(java.awt.Color.WHITE);

        // ----- Paginacion -----
        final int filasPorPagina = 8;
        int[] paginaActual = {0};
        int[] columnaOrden = {-1};
        boolean[] ordenAscendente = {true};
        java.util.List<Object[]> filasFiltradas = new java.util.ArrayList<>(resultadoInicial.filas);

        javax.swing.JLabel etiquetaPaginacion = new javax.swing.JLabel(" ");
        etiquetaPaginacion.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        etiquetaPaginacion.setForeground(COLOR_TEXTO_GRIS);

        javax.swing.JPanel panelBotonesPagina = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 6, 0));
        panelBotonesPagina.setOpaque(false);

        Runnable[] refrescarTabla = new Runnable[1];

        javax.swing.JButton botonPrimeraPagina = new javax.swing.JButton();
        prepararBotonIcono(botonPrimeraPagina, iconoDobleChevronIzquierda(COLOR_AZUL_OSCURO));
        javax.swing.JButton botonPaginaAnterior = new javax.swing.JButton();
        prepararBotonIcono(botonPaginaAnterior, iconoChevronIzquierda(COLOR_AZUL_OSCURO));
        javax.swing.JButton botonPaginaSiguiente = new javax.swing.JButton();
        prepararBotonIcono(botonPaginaSiguiente, iconoChevronDerecha(COLOR_AZUL_OSCURO));
        javax.swing.JButton botonUltimaPagina = new javax.swing.JButton();
        prepararBotonIcono(botonUltimaPagina, iconoDobleChevronDerecha(COLOR_AZUL_OSCURO));

        botonPrimeraPagina.addActionListener(e -> {
            paginaActual[0] = 0;
            refrescarTabla[0].run();
        });
        botonPaginaAnterior.addActionListener(e -> {
            paginaActual[0]--;
            refrescarTabla[0].run();
        });
        botonPaginaSiguiente.addActionListener(e -> {
            paginaActual[0]++;
            refrescarTabla[0].run();
        });
        botonUltimaPagina.addActionListener(e -> {
            paginaActual[0] = Integer.MAX_VALUE;
            refrescarTabla[0].run();
        });

        refrescarTabla[0] = () -> {
            int total = filasFiltradas.size();
            int totalPaginas = Math.max(1, (int) Math.ceil(total / (double) filasPorPagina));
            if (paginaActual[0] >= totalPaginas) {
                paginaActual[0] = totalPaginas - 1;
            }
            if (paginaActual[0] < 0) {
                paginaActual[0] = 0;
            }
            int inicio = paginaActual[0] * filasPorPagina;
            int fin = Math.min(total, inicio + filasPorPagina);

            modelo.setRowCount(0);
            for (int i = inicio; i < fin; i++) {
                modelo.addRow(filasFiltradas.get(i));
            }
            ajustarAnchoColumnasHistorial(tabla);

            etiquetaPaginacion.setText(total == 0
                    ? "No hay \u00f3rdenes para mostrar"
                    : String.format("Mostrando %d - %d de %d \u00f3rdenes", inicio + 1, fin, total));

            panelBotonesPagina.removeAll();
            panelBotonesPagina.add(botonPrimeraPagina);
            panelBotonesPagina.add(botonPaginaAnterior);
            int ventana = 5;
            int inicioBotones = Math.max(0, paginaActual[0] - ventana / 2);
            int finBotones = Math.min(totalPaginas, inicioBotones + ventana);
            inicioBotones = Math.max(0, finBotones - ventana);
            for (int pagina = inicioBotones; pagina < finBotones; pagina++) {
                javax.swing.JButton botonPagina = new javax.swing.JButton();
                boolean actual = pagina == paginaActual[0];
                prepararBoton(botonPagina, String.valueOf(pagina + 1), null,
                        actual ? COLOR_AZUL : java.awt.Color.WHITE,
                        actual ? java.awt.Color.WHITE : COLOR_TEXTO_GRIS,
                        actual ? null : COLOR_BORDE_CAMPO);
                botonPagina.setMargin(new java.awt.Insets(2, 2, 2, 2));
                botonPagina.setPreferredSize(new java.awt.Dimension(30, 30));
                int paginaDestino = pagina;
                botonPagina.addActionListener(e -> {
                    paginaActual[0] = paginaDestino;
                    refrescarTabla[0].run();
                });
                panelBotonesPagina.add(botonPagina);
            }
            panelBotonesPagina.add(botonPaginaSiguiente);
            panelBotonesPagina.add(botonUltimaPagina);
            botonPrimeraPagina.setEnabled(paginaActual[0] > 0);
            botonPaginaAnterior.setEnabled(paginaActual[0] > 0);
            botonPaginaSiguiente.setEnabled(paginaActual[0] < totalPaginas - 1);
            botonUltimaPagina.setEnabled(paginaActual[0] < totalPaginas - 1);
            panelBotonesPagina.revalidate();
            panelBotonesPagina.repaint();
        };

        Runnable aplicarFiltros = () -> {
            boolean esPlaceholder = Boolean.TRUE.equals(campoBuscar.getClientProperty("placeholder-activo"));
            String filtroTexto = esPlaceholder || campoBuscar.getText() == null
                    ? "" : campoBuscar.getText().trim().toLowerCase(java.util.Locale.ROOT);
            String filtroFecha = String.valueOf(comboFecha.getSelectedItem());

            ResultadoHistorial resultadoActual = cargarFilasHistorial(idProveedor);
            java.util.List<Object[]> filasOrigen = resultadoActual == null ? resultadoInicial.filas : resultadoActual.filas;
            if (resultadoActual != null) {
                tarjetaCompras.etiquetaValor.setText(String.valueOf(filasOrigen.size()));
                tarjetaTotalOrdenes.etiquetaValor.setText(
                        String.format("%s %,.2f", resultadoActual.simboloMoneda, resultadoActual.sumaTotal));
                tarjetaAbonado.etiquetaValor.setText(
                        String.format("%s %,.2f", resultadoActual.simboloMoneda, resultadoActual.sumaAbonado));
            }

            filasFiltradas.clear();
            for (Object[] fila : filasOrigen) {
                if (!filtroTexto.isEmpty()) {
                    boolean coincide = String.valueOf(fila[0]).toLowerCase(java.util.Locale.ROOT).contains(filtroTexto)
                            || String.valueOf(fila[1]).toLowerCase(java.util.Locale.ROOT).contains(filtroTexto)
                            || String.valueOf(fila[2]).toLowerCase(java.util.Locale.ROOT).contains(filtroTexto)
                            || String.valueOf(fila[6]).toLowerCase(java.util.Locale.ROOT).contains(filtroTexto)
                            || String.valueOf(fila[7]).toLowerCase(java.util.Locale.ROOT).contains(filtroTexto);
                    if (!coincide) {
                        continue;
                    }
                }
                if (!coincideFiltroFecha(String.valueOf(fila[3]), filtroFecha)) {
                    continue;
                }
                filasFiltradas.add(fila);
            }
            if (columnaOrden[0] >= 0) {
                ordenarFilas(filasFiltradas, columnaOrden[0], ordenAscendente[0]);
            }
            paginaActual[0] = 0;
            refrescarTabla[0].run();
        };

        campoBuscar.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
            @Override
            public void insertUpdate(javax.swing.event.DocumentEvent e) {
                aplicarFiltros.run();
            }

            @Override
            public void removeUpdate(javax.swing.event.DocumentEvent e) {
                aplicarFiltros.run();
            }

            @Override
            public void changedUpdate(javax.swing.event.DocumentEvent e) {
                aplicarFiltros.run();
            }
        });
        comboFecha.addActionListener(e -> aplicarFiltros.run());
        tabla.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mousePressed(java.awt.event.MouseEvent evento) {
                if (!javax.swing.SwingUtilities.isLeftMouseButton(evento)) {
                    return;
                }
                int filaVista = tabla.rowAtPoint(evento.getPoint());
                int columnaVista = tabla.columnAtPoint(evento.getPoint());
                if (filaVista < 0 || columnaVista < 0) {
                    return;
                }
                int columnaModelo = tabla.convertColumnIndexToModel(columnaVista);
                if (columnaModelo != 13) {
                    return;
                }
                int filaModelo = tabla.convertRowIndexToModel(filaVista);
                String noOrden = String.valueOf(modelo.getValueAt(filaModelo, 2)).trim();
                String estatus = String.valueOf(modelo.getValueAt(filaModelo, 13)).trim();
                if ("FINALIZADO".equalsIgnoreCase(estatus)) {
                    tabla.setRowSelectionInterval(filaVista, filaVista);
                    mostrarDetalleOrden(noOrden, tabla);
                    return;
                }
                double pendientePago = parseNumeroOrdenable(modelo.getValueAt(filaModelo, 10));
                double pendienteMercancia = parseNumeroOrdenable(modelo.getValueAt(filaModelo, 12));
                tabla.setRowSelectionInterval(filaVista, filaVista);
                mostrarAccionesPendiente(noOrden, pendientePago, pendienteMercancia, tabla, aplicarFiltros);
            }

            @Override
            public void mouseClicked(java.awt.event.MouseEvent evento) {
                if (tabla.rowAtPoint(evento.getPoint()) < 0) {
                    tabla.clearSelection();
                }
            }
        });

        tabla.getTableHeader().addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mouseClicked(java.awt.event.MouseEvent e) {
                int columnaVista = tabla.columnAtPoint(e.getPoint());
                if (columnaVista < 0) {
                    return;
                }
                int columnaModelo = tabla.convertColumnIndexToModel(columnaVista);
                if (columnaOrden[0] == columnaModelo) {
                    ordenAscendente[0] = !ordenAscendente[0];
                } else {
                    columnaOrden[0] = columnaModelo;
                    ordenAscendente[0] = true;
                }
                ordenarFilas(filasFiltradas, columnaModelo, ordenAscendente[0]);
                paginaActual[0] = 0;
                refrescarTabla[0].run();
            }
        });

        // ----- Contenedor de tabla (tarjeta blanca con borde) -----
        RoundedPanel contenedorTabla = new RoundedPanel(12, java.awt.Color.WHITE);
        contenedorTabla.setLayout(new java.awt.BorderLayout(0, 10));
        contenedorTabla.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO),
                javax.swing.BorderFactory.createEmptyBorder(14, 16, 10, 16)));
        contenedorTabla.add(panelBusquedaFiltros, java.awt.BorderLayout.NORTH);
        contenedorTabla.add(scroll, java.awt.BorderLayout.CENTER);

        javax.swing.JPanel panelPaginacion = new javax.swing.JPanel(new java.awt.BorderLayout());
        panelPaginacion.setOpaque(false);
        panelPaginacion.setBorder(javax.swing.BorderFactory.createEmptyBorder(10, 2, 0, 2));
        panelPaginacion.add(etiquetaPaginacion, java.awt.BorderLayout.WEST);
        panelPaginacion.add(panelBotonesPagina, java.awt.BorderLayout.EAST);
        contenedorTabla.add(panelPaginacion, java.awt.BorderLayout.SOUTH);

        // ----- Barra inferior (historial de recepciones / cerrar) -----
        javax.swing.JDialog[] dialogoRef = new javax.swing.JDialog[1];

        javax.swing.JButton botonCerrar = new javax.swing.JButton();
        prepararBoton(botonCerrar, "Cerrar", iconoCerrarX(new java.awt.Color(60, 70, 85)),
                java.awt.Color.WHITE, new java.awt.Color(60, 70, 85), COLOR_BORDE_CAMPO);
        botonCerrar.addActionListener(e -> {
            if (dialogoRef[0] != null) {
                dialogoRef[0].dispose();
            }
        });

        javax.swing.JPanel panelExportar = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.LEFT, 10, 0));
        panelExportar.setOpaque(false);
        javax.swing.JButton botonRecepciones = new javax.swing.JButton();
        prepararBoton(botonRecepciones, "Historial de recepciones",
                iconoCarrito(java.awt.Color.WHITE, 18), COLOR_AZUL, java.awt.Color.WHITE, null);
        botonRecepciones.addActionListener(e -> {
            java.util.Set<String> ordenesSeleccionadas = new java.util.LinkedHashSet<>();
            for (int filaVista : tabla.getSelectedRows()) {
                int filaModelo = tabla.convertRowIndexToModel(filaVista);
                String orden = String.valueOf(modelo.getValueAt(filaModelo, 2)).trim();
                if (!orden.isEmpty()) {
                    ordenesSeleccionadas.add(orden);
                }
            }
            mostrarHistorialRecepciones(ordenesSeleccionadas);
        });
        panelExportar.add(botonRecepciones);

        javax.swing.JPanel panelCerrar = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 0, 0));
        panelCerrar.setOpaque(false);
        panelCerrar.add(botonCerrar);

        javax.swing.JPanel panelInferior = new javax.swing.JPanel(new java.awt.BorderLayout());
        panelInferior.setOpaque(false);
        panelInferior.setBorder(javax.swing.BorderFactory.createEmptyBorder(14, 0, 0, 0));
        panelInferior.add(panelExportar, java.awt.BorderLayout.WEST);
        panelInferior.add(panelCerrar, java.awt.BorderLayout.EAST);

        // ----- Ensamblado final -----
        javax.swing.JPanel panelCentro = new javax.swing.JPanel(new java.awt.BorderLayout(0, 14));
        panelCentro.setOpaque(false);
        panelCentro.add(panelTarjetas, java.awt.BorderLayout.NORTH);
        panelCentro.add(contenedorTabla, java.awt.BorderLayout.CENTER);

        javax.swing.JPanel contenedorPrincipal = new javax.swing.JPanel(new java.awt.BorderLayout(0, 14));
        contenedorPrincipal.setBackground(COLOR_FONDO);
        contenedorPrincipal.setOpaque(true);
        contenedorPrincipal.setBorder(javax.swing.BorderFactory.createEmptyBorder(18, 20, 18, 20));
        contenedorPrincipal.add(panelHeader, java.awt.BorderLayout.NORTH);
        contenedorPrincipal.add(panelCentro, java.awt.BorderLayout.CENTER);
        contenedorPrincipal.add(panelInferior, java.awt.BorderLayout.SOUTH);

        contenedorPrincipal.setPreferredSize(new java.awt.Dimension(
                Math.min(1300, areaUtil.width - 60), Math.min(760, areaUtil.height - 80)));

        javax.swing.JDialog dialogo = new javax.swing.JDialog(this,
                idProveedor.isEmpty() ? "Historial de \u00d3rdenes de Compra - Todos los proveedores"
                        : "Historial de \u00d3rdenes de Compra - " + idProveedor,
                true);
        dialogoRef[0] = dialogo;
        dialogo.getContentPane().add(contenedorPrincipal);
        dialogo.setResizable(true);
        dialogo.pack();
        dialogo.setLocationRelativeTo(this);

        aplicarFiltros.run();

        dialogo.setVisible(true);
    }//mostrarHistorialCompras

    private void mostrarHistorialRecepciones(String ordenFiltro) {
        java.util.Set<String> ordenes = new java.util.LinkedHashSet<>();
        if (ordenFiltro != null && !ordenFiltro.isBlank()) {
            ordenes.add(ordenFiltro);
        }
        mostrarHistorialRecepciones(ordenes);
    }//mostrarHistorialRecepciones

    private void mostrarHistorialRecepciones(java.util.Set<String> ordenesFiltro) {
        java.util.List<String[]> registros = cargarRegistrosRecepcion(ordenesFiltro);
        String[] columnas = {
            "ORDEN", "PROVEEDOR", "FECHA DE ENTRADA", "HORA", "USUARIO",
            "NOMBRE", "MODELO ID", "DESCRIPCIÓN", "CANTIDAD", "ESTATUS ENTRADA"
        };
        javax.swing.table.DefaultTableModel modelo = new javax.swing.table.DefaultTableModel(columnas, 0) {
            @Override
            public boolean isCellEditable(int fila, int columna) {
                return false;
            }
        };
        for (String[] registro : registros) {
            modelo.addRow(registro);
        }

        javax.swing.JTable tabla = new javax.swing.JTable(modelo);
        tabla.setAutoCreateRowSorter(true);
        tabla.setRowHeight(30);
        tabla.setFillsViewportHeight(true);
        tabla.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        tabla.getTableHeader().setReorderingAllowed(false);
        tabla.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12));
        configurarEncabezadoHistorial(tabla);
        configurarEstatusEntradaRecepciones(tabla);
        ajustarAnchoColumnasRecepciones(tabla);

        javax.swing.JTextField campoBuscar = new javax.swing.JTextField();
        campoBuscar.setToolTipText("Filtrar por proveedor, modelo ID o descripción");
        campoBuscar.setPreferredSize(new java.awt.Dimension(320, 32));
        javax.swing.JButton botonBuscar = new javax.swing.JButton("Buscar");
        prepararBoton(botonBuscar, "Buscar", iconoBuscar(java.awt.Color.WHITE),
                COLOR_AZUL, java.awt.Color.WHITE, null);

        Runnable[] aplicarFiltro = new Runnable[1];
        aplicarFiltro[0] = () -> {
            boolean esPlaceholder = Boolean.TRUE.equals(campoBuscar.getClientProperty("placeholder-activo"));
            String filtro = esPlaceholder ? ""
                    : campoBuscar.getText().trim().toLowerCase(java.util.Locale.ROOT);
            modelo.setRowCount(0);
            for (String[] registro : cargarRegistrosRecepcion(ordenesFiltro)) {
                String texto = String.join(" ", registro).toLowerCase(java.util.Locale.ROOT);
                if (filtro.isEmpty() || texto.contains(filtro)) {
                    modelo.addRow(registro);
                }
            }
            ajustarAnchoColumnasRecepciones(tabla);
        };
        campoBuscar.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
            private void actualizar() {
                aplicarFiltro[0].run();
            }

            @Override
            public void insertUpdate(javax.swing.event.DocumentEvent evento) {
                actualizar();
            }

            @Override
            public void removeUpdate(javax.swing.event.DocumentEvent evento) {
                actualizar();
            }

            @Override
            public void changedUpdate(javax.swing.event.DocumentEvent evento) {
                actualizar();
            }
        });
        botonBuscar.addActionListener(e -> aplicarFiltro[0].run());
        campoBuscar.addActionListener(e -> aplicarFiltro[0].run());

        javax.swing.JPanel panelFiltros = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.LEFT, 8, 0));
        panelFiltros.setOpaque(false);
        panelFiltros.add(new javax.swing.JLabel("Buscar:"));
        panelFiltros.add(campoBuscar);
        panelFiltros.add(botonBuscar);

        javax.swing.JScrollPane scroll = new javax.swing.JScrollPane(tabla);
        java.awt.Rectangle areaUtil = java.awt.GraphicsEnvironment
                .getLocalGraphicsEnvironment().getMaximumWindowBounds();
        int anchoVentana = Math.min(1500, areaUtil.width - 40);
        int altoVentana = Math.min(850, areaUtil.height - 40);
        scroll.setPreferredSize(new java.awt.Dimension(anchoVentana - 40, altoVentana - 130));
        javax.swing.JPanel contenido = new javax.swing.JPanel(new java.awt.BorderLayout(0, 12));
        contenido.setBorder(javax.swing.BorderFactory.createEmptyBorder(16, 16, 16, 16));
        contenido.add(panelFiltros, java.awt.BorderLayout.NORTH);
        contenido.add(scroll, java.awt.BorderLayout.CENTER);

        String titulo = ordenesFiltro.isEmpty()
                ? "Historial de recepciones"
                : "Recepciones de órdenes seleccionadas";
        javax.swing.JDialog dialogo = new javax.swing.JDialog(this, titulo, true);
        dialogo.getContentPane().add(contenido);
        dialogo.setSize(anchoVentana, altoVentana);
        dialogo.setResizable(true);
        dialogo.setLocationRelativeTo(this);
        dialogo.setVisible(true);
    }//mostrarHistorialRecepciones

    private void ajustarAnchoColumnasRecepciones(javax.swing.JTable tabla) {
        javax.swing.table.JTableHeader encabezado = tabla.getTableHeader();
        java.awt.FontMetrics metricas = encabezado.getFontMetrics(encabezado.getFont());
        for (int columna = 0; columna < tabla.getColumnCount(); columna++) {
            int ancho = metricas.stringWidth(tabla.getColumnName(columna));
            for (int fila = 0; fila < tabla.getRowCount(); fila++) {
                javax.swing.table.TableCellRenderer renderer = tabla.getCellRenderer(fila, columna);
                java.awt.Component componente = renderer.getTableCellRendererComponent(
                        tabla, tabla.getValueAt(fila, columna), false, false, fila, columna);
                ancho = Math.max(ancho, componente.getPreferredSize().width);
            }
            int minimo = columna == 2 ? 125 : 70;
            int maximo = columna == 5 || columna == 7 ? 260 : 180;
            int anchoAjustado = Math.min(maximo, Math.max(minimo, ancho + 18));
            tabla.getColumnModel().getColumn(columna).setPreferredWidth(anchoAjustado);
        }
    }//ajustarAnchoColumnasRecepciones

    private void configurarEstatusEntradaRecepciones(javax.swing.JTable tabla) {
        int columnaEstatus = tabla.getColumnModel().getColumnIndex("ESTATUS ENTRADA");
        tabla.getColumnModel().getColumn(columnaEstatus).setCellRenderer(
                new javax.swing.table.DefaultTableCellRenderer() {
                    @Override
                    public java.awt.Component getTableCellRendererComponent(javax.swing.JTable tabla,
                            Object valor, boolean seleccionado, boolean foco, int fila, int columna) {
                        java.awt.Component componente = super.getTableCellRendererComponent(
                                tabla, valor, seleccionado, foco, fila, columna);
                        String estatus = String.valueOf(valor).toUpperCase(java.util.Locale.ROOT);
                        if (!seleccionado) {
                            if (estatus.contains("RETRASO")) {
                                componente.setBackground(new java.awt.Color(255, 226, 226));
                                componente.setForeground(new java.awt.Color(170, 0, 0));
                            } else {
                                componente.setBackground(new java.awt.Color(226, 246, 234));
                                componente.setForeground(new java.awt.Color(0, 115, 55));
                            }
                        }
                        setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
                        return componente;
                    }
                });
    }//configurarEstatusEntradaRecepciones

    private java.util.List<String[]> cargarRegistrosRecepcion(String ordenFiltro) {
        java.util.Set<String> ordenes = new java.util.LinkedHashSet<>();
        if (ordenFiltro != null && !ordenFiltro.isBlank()) {
            ordenes.add(ordenFiltro);
        }
        return cargarRegistrosRecepcion(ordenes);
    }//cargarRegistrosRecepcion

    private java.util.List<String[]> cargarRegistrosRecepcion(java.util.Set<String> ordenesFiltro) {
        java.util.List<String[]> registros = new java.util.ArrayList<>();
        File archivo = AppConfig.resolveBaseFile("REGISTRORECIBIR.xlsx");
        if (!archivo.exists() || archivo.length() == 0) {
            return registros;
        }
        try (FileInputStream fis = new FileInputStream(archivo);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            java.util.Map<String, String> descripciones = cargarDescripcionesInventario();
            java.util.Map<String, java.time.LocalDate> fechasEntrega = cargarFechasEntregaOrdenes();
            java.util.Map<String, Integer> columnas = new java.util.HashMap<>();
            Row encabezado = hoja.getRow(0);
            if (encabezado != null) {
                for (int columna = 0; columna < 20; columna++) {
                    String nombre = leerCelda(encabezado, columna).toUpperCase(java.util.Locale.ROOT);
                    if (!nombre.isEmpty()) {
                        columnas.put(nombre, columna);
                    }
                }
            }
            int indiceOrden = columnas.getOrDefault("ORDENCOMPRA", 0);
            int indiceProveedor = columnas.getOrDefault("PROVEEDOR", 1);
            int indiceFecha = columnas.getOrDefault("FECHA_RECIBIDO", 2);
            int indiceHora = columnas.getOrDefault("HORA_RECIBIDO", -1);
            int indiceUsuario = columnas.getOrDefault("USUARIO", -1);
            int indiceNombreUsuario = columnas.getOrDefault("NOMBRE_USUARIO", -1);
            int indiceModelo = columnas.containsKey("MODELO")
                    ? columnas.get("MODELO") : 3;
            int indiceCantidad = columnas.containsKey("CANTIDAD")
                    ? columnas.get("CANTIDAD") : 4;
            int indiceTipo = columnas.getOrDefault("TIPO_RECEPCION", -1);
            for (Row fila : hoja) {
                if (fila.getRowNum() == 0) {
                    continue;
                }
                String orden = leerCelda(fila, indiceOrden);
                boolean coincideOrden = ordenesFiltro.isEmpty()
                        || ordenesFiltro.stream().anyMatch(ordenFiltro -> ordenFiltro.equalsIgnoreCase(orden));
                if (orden.isEmpty() || !coincideOrden) {
                    continue;
                }
                String idProveedor = leerCelda(fila, indiceProveedor);
                String codigoProducto = leerCelda(fila, indiceModelo);
                registros.add(new String[]{
                    orden,
                    obtenerNombreProveedorPorId(idProveedor),
                    leerCelda(fila, indiceFecha),
                    indiceHora < 0 ? "" : leerCelda(fila, indiceHora),
                    indiceUsuario < 0 ? "" : leerCelda(fila, indiceUsuario),
                    indiceNombreUsuario < 0 ? "" : leerCelda(fila, indiceNombreUsuario),
                    codigoProducto,
                    descripciones.getOrDefault(
                            codigoProducto.toUpperCase(java.util.Locale.ROOT), codigoProducto),
                    leerCelda(fila, indiceCantidad),
                    calcularEstatusEntrada(orden, leerCelda(fila, indiceFecha), fechasEntrega)
                });
            }
        } catch (IOException ex) {
            mostrarError("No se pudo leer el historial de recepciones.", ex);
        }
        return registros;
    }//cargarRegistrosRecepcion

    private java.util.Map<String, java.time.LocalDate> cargarFechasEntregaOrdenes() {
        java.util.Map<String, java.time.LocalDate> fechas = new java.util.HashMap<>();
        File archivoOrdenes = new File(ORDENES_FILE_ABSOLUTE);
        if (!archivoOrdenes.exists() || archivoOrdenes.length() == 0) {
            return fechas;
        }
        try (FileInputStream fis = new FileInputStream(archivoOrdenes);
             Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (Row fila : hoja) {
                if (fila.getRowNum() == 0) {
                    continue;
                }
                String orden = leerCelda(fila, 0);
                java.util.Date fechaEntrega = leerFechaCelda(fila, 11);
                if (!orden.isEmpty() && fechaEntrega != null) {
                    fechas.put(orden.toUpperCase(java.util.Locale.ROOT),
                            fechaEntrega.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate());
                }
            }
        } catch (IOException ex) {
            logger.log(java.util.logging.Level.WARNING,
                    "No se pudieron cargar las fechas esperadas de entrega.", ex);
        }
        return fechas;
    }//cargarFechasEntregaOrdenes

    private String calcularEstatusEntrada(String noOrden, String fechaEntrada,
            java.util.Map<String, java.time.LocalDate> fechasEntrega) {
        java.time.LocalDate fechaEsperada = fechasEntrega.get(noOrden.toUpperCase(java.util.Locale.ROOT));
        if (fechaEsperada == null || fechaEntrada == null || fechaEntrada.isBlank()) {
            return "SIN FECHA ESPERADA";
        }
        try {
            java.time.LocalDate fechaReal = java.time.LocalDate.parse(fechaEntrada.trim(),
                    java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
            if (fechaReal.isEqual(fechaEsperada)) {
                return "A TIEMPO (0 DÍAS)";
            }
            if (fechaReal.isBefore(fechaEsperada)) {
                long diasAnticipacion = java.time.temporal.ChronoUnit.DAYS.between(fechaReal, fechaEsperada);
                return diasAnticipacion == 1
                        ? "1 DÍA ADELANTADA"
                        : diasAnticipacion + " DÍAS ADELANTADA";
            }
            long diasRetraso = java.time.temporal.ChronoUnit.DAYS.between(fechaEsperada, fechaReal);
            return diasRetraso == 1
                    ? "1 DÍA DE RETRASO"
                    : diasRetraso + " DÍAS DE RETRASO";
        } catch (java.time.format.DateTimeParseException ex) {
            return "SIN FECHA ESPERADA";
        }
    }//calcularEstatusEntrada

    private java.util.Map<String, String> cargarDescripcionesInventario() {
        java.util.Map<String, String> descripciones = new java.util.HashMap<>();
        File archivo = INVENTARIO_REPOSITORY.obtenerArchivoInventarios();
        if (!archivo.exists() || archivo.length() == 0) {
            return descripciones;
        }
        try (Workbook libro = INVENTARIO_REPOSITORY.cargarLibro(archivo)) {
            Sheet hoja = INVENTARIO_REPOSITORY.obtenerHojaInventarios(libro);
            for (Row fila : hoja) {
                if (fila.getRowNum() == 0) {
                    continue;
                }
                String codigo = INVENTARIO_REPOSITORY.leerCelda(fila, 0);
                String descripcion = INVENTARIO_REPOSITORY.leerCelda(fila, 1);
                if (!codigo.isEmpty() && !descripcion.isEmpty()) {
                    descripciones.put(codigo.toUpperCase(java.util.Locale.ROOT), descripcion);
                }
            }
        } catch (IOException ex) {
            logger.log(java.util.logging.Level.WARNING,
                    "No se pudieron cargar las descripciones del inventario.", ex);
        }
        return descripciones;
    }//cargarDescripcionesInventario

    private void mostrarAccionesPendiente(String noOrden, double pendientePago,
            double pendienteMercancia, javax.swing.JTable tabla, Runnable refrescar) {
        java.util.List<String> opcionesDisponibles = new java.util.ArrayList<>();
        if (pendientePago > 0) {
            opcionesDisponibles.add("Registrar abono");
        }
        if (pendienteMercancia > 0) {
            opcionesDisponibles.add("Recibir mercancía");
        }
        if (opcionesDisponibles.isEmpty()) {
            return;
        }
        opcionesDisponibles.add("Cancelar");
        Object[] opciones = opcionesDisponibles.toArray();
        int seleccion = JOptionPane.showOptionDialog(this,
                "¿Qué deseas registrar para la orden " + noOrden + "?",
                "Orden pendiente",
                JOptionPane.DEFAULT_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                opciones,
                opciones[0]);
        if (seleccion < 0 || seleccion == opcionesDisponibles.size() - 1) {
            return;
        }
        if ("Registrar abono".equals(opcionesDisponibles.get(seleccion))) {
            registrarAbono(noOrden, tabla);
            refrescar.run();
        } else if ("Recibir mercancía".equals(opcionesDisponibles.get(seleccion))) {
            mostrarDetalleOrden(noOrden, tabla);
            refrescar.run();
        }
    }//mostrarAccionesPendiente

    private void ordenarFilas(java.util.List<Object[]> filas, int columna, boolean ascendente) {
        java.util.Comparator<Object[]> comparador;
        if (columna == 2 || columna == 8 || columna == 9 || columna == 10
                || columna == 11 || columna == 12) {
            comparador = java.util.Comparator.comparingDouble(f -> parseNumeroOrdenable(f[columna]));
        } else if (columna == 3 || columna == 4 || columna == 5) {
            comparador = java.util.Comparator.comparing(f -> parseFechaOrdenable(f[columna]));
        } else {
            comparador = java.util.Comparator.comparing(f -> String.valueOf(f[columna]), String.CASE_INSENSITIVE_ORDER);
        }
        if (!ascendente) {
            comparador = comparador.reversed();
        }
        filas.sort(comparador);
    }//ordenarFilas

    private boolean coincideFiltroFecha(String fechaTexto, String filtro) {
        if (filtro == null || "Todas las fechas".equals(filtro)) {
            return true;
        }
        if (fechaTexto == null || fechaTexto.isEmpty()) {
            return false;
        }
        java.time.LocalDate fecha;
        try {
            fecha = java.time.LocalDate.parse(fechaTexto, java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        } catch (java.time.format.DateTimeParseException ex) {
            return true;
        }
        java.time.LocalDate hoy = java.time.LocalDate.now();
        switch (filtro) {
            case "Hoy":
                return fecha.isEqual(hoy);
            case "\u00daltimos 7 d\u00edas":
                return !fecha.isBefore(hoy.minusDays(7)) && !fecha.isAfter(hoy);
            case "Este mes":
                return fecha.getMonthValue() == hoy.getMonthValue() && fecha.getYear() == hoy.getYear();
            case "Este a\u00f1o":
                return fecha.getYear() == hoy.getYear();
            default:
                return true;
        }
    }//coincideFiltroFecha

    private static void configurarPlaceholder(javax.swing.JTextField campo, String textoPlaceholder) {
        java.awt.Color colorTextoNormal = new java.awt.Color(40, 50, 65);
        java.awt.Color colorPlaceholder = COLOR_TEXTO_GRIS;
        campo.setText(textoPlaceholder);
        campo.setForeground(colorPlaceholder);
        campo.putClientProperty("placeholder-activo", Boolean.TRUE);
        campo.addFocusListener(new java.awt.event.FocusAdapter() {
            @Override
            public void focusGained(java.awt.event.FocusEvent e) {
                if (Boolean.TRUE.equals(campo.getClientProperty("placeholder-activo"))) {
                    campo.putClientProperty("placeholder-activo", Boolean.FALSE);
                    campo.setText("");
                    campo.setForeground(colorTextoNormal);
                }
            }

            @Override
            public void focusLost(java.awt.event.FocusEvent e) {
                if (campo.getText().isEmpty()) {
                    campo.putClientProperty("placeholder-activo", Boolean.TRUE);
                    campo.setForeground(colorPlaceholder);
                    campo.setText(textoPlaceholder);
                }
            }
        });
    }//configurarPlaceholder

    private void configurarEncabezadoHistorial(javax.swing.JTable tabla) {
        javax.swing.table.JTableHeader encabezado = tabla.getTableHeader();
        encabezado.setDefaultRenderer((t, valor, seleccionado, foco, fila, columna) -> {
            String textoEncabezado = formatearEncabezadoHistorial(String.valueOf(valor));
            javax.swing.JLabel etiqueta = new javax.swing.JLabel(
                    "<html><div style='width:88px;text-align:left;'>"
                    + textoEncabezado + "</div></html>");
            etiqueta.setOpaque(true);
            etiqueta.setBackground(COLOR_AZUL);
            etiqueta.setForeground(java.awt.Color.WHITE);
            etiqueta.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
            etiqueta.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
            etiqueta.setVerticalAlignment(javax.swing.SwingConstants.CENTER);
            etiqueta.setBorder(javax.swing.BorderFactory.createEmptyBorder(6, 4, 6, 4));
            return etiqueta;
        });
        encabezado.setPreferredSize(new java.awt.Dimension(encabezado.getPreferredSize().width, 64));
    }//configurarEncabezadoHistorial

    private String formatearEncabezadoHistorial(String encabezado) {
        String texto = encabezado
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;");
        switch (encabezado) {
            case "NOMBRE PROVEEDOR":
                return "NOMBRE<br>PROVEEDOR";
            case "NO. ORDEN":
                return "NO. ORDEN";
            case "FECHA DE CREACI\u00d3N":
                return "FECHA DE<br>CREACI\u00d3N";
            case "FECHA DE PAGO PROMETIDA":
                return "FECHA DE<br>PAGO<br>PROMETIDA";
            case "FECHA DE RECEPCI\u00d3N PROMETIDA":
                return "FECHA DE<br>RECEPCI\u00d3N<br>PROMETIDA";
            case "PEND. X PAGAR":
                return "PEND. X<br>PAGAR";
            case "PROD. X RECIBIR":
                return "PROD. X<br>RECIBIR";
            default:
                return texto;
        }
    }//formatearEncabezadoHistorial

    private void configurarColumnaEstatusHistorial(javax.swing.JTable tabla) {
        tabla.getColumn("ESTATUS").setCellRenderer((t, valor, seleccionado, foco, fila, columna) -> {
            String estatus = String.valueOf(valor).trim();
            java.awt.Color colorTexto;
            java.awt.Color colorFondo;
            switch (estatus) {
                case "FINALIZADO":
                    colorTexto = COLOR_VERDE;
                    colorFondo = COLOR_VERDE_CLARO;
                    break;
                case "PENDIENTE DE PAGO":
                    colorTexto = COLOR_NARANJA;
                    colorFondo = COLOR_NARANJA_CLARO;
                    break;
                case "PENDIENTE DE MERCANCIA":
                    colorTexto = COLOR_AZUL;
                    colorFondo = COLOR_AZUL_CLARO;
                    break;
                case "PENDIENTE PAGO Y MERCANCIA":
                    colorTexto = COLOR_ROJO;
                    colorFondo = COLOR_ROJO_CLARO;
                    break;
                default:
                    colorTexto = COLOR_GRIS;
                    colorFondo = COLOR_GRIS_CLARO;
            }
            javax.swing.JComponent pastilla = crearPastillaEstatus(estatus, colorTexto, colorFondo);
            boolean accionable = !"FINALIZADO".equalsIgnoreCase(estatus);
            pastilla.setCursor(java.awt.Cursor.getPredefinedCursor(
                    java.awt.Cursor.HAND_CURSOR));
            pastilla.setToolTipText(accionable
                    ? "Selecciona para registrar un abono o recibir mercancía"
                    : "Selecciona para consultar los detalles de la orden finalizada");
            return pastilla;
        });
    }//configurarColumnaEstatusHistorial

    private static javax.swing.JComponent crearPastillaEstatus(String texto, java.awt.Color colorTexto, java.awt.Color colorFondo) {
        javax.swing.JPanel pastilla = new javax.swing.JPanel(new java.awt.BorderLayout(4, 0)) {
            @Override
            protected void paintComponent(java.awt.Graphics g) {
                java.awt.Graphics2D g2 = (java.awt.Graphics2D) g.create();
                g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING, java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
                g2.setColor(colorFondo);
                g2.fillRoundRect(0, 0, getWidth() - 1, getHeight() - 1, getHeight(), getHeight());
                g2.dispose();
                super.paintComponent(g);
            }
        };
        pastilla.setOpaque(false);
        pastilla.setBorder(javax.swing.BorderFactory.createEmptyBorder(4, 8, 4, 8));

        javax.swing.JLabel punto = new javax.swing.JLabel(iconoPuntoRelleno(colorTexto));
        punto.setVerticalAlignment(javax.swing.SwingConstants.TOP);
        punto.setBorder(javax.swing.BorderFactory.createEmptyBorder(3, 0, 0, 0));

        javax.swing.JLabel etiqueta = new javax.swing.JLabel(texto);
        etiqueta.setOpaque(false);
        etiqueta.setForeground(colorTexto);
        etiqueta.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 10));
        pastilla.add(punto, java.awt.BorderLayout.WEST);
        pastilla.add(etiqueta, java.awt.BorderLayout.CENTER);
        return pastilla;
    }//crearPastillaEstatus

    private static javax.swing.Icon iconoPuntoRelleno(java.awt.Color color) {
        return new VectorIcon(8, color, (g2, t) -> g2.fillOval(0, 0, t, t));
    }

    private void exportarHistorialExcel(java.util.List<Object[]> filas, String idProveedor) {
        javax.swing.JFileChooser selector = new javax.swing.JFileChooser();
        selector.setDialogTitle("Exportar historial a Excel");
        String nombreSugerido = "HistorialCompras" + (idProveedor.isEmpty() ? "" : "_" + idProveedor) + ".xlsx";
        selector.setSelectedFile(new File(System.getProperty("user.home") + "\\Desktop\\" + nombreSugerido));
        selector.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Excel (*.xlsx)", "xlsx"));
        if (selector.showSaveDialog(this) != javax.swing.JFileChooser.APPROVE_OPTION) {
            return;
        }
        String ruta = selector.getSelectedFile().getAbsolutePath();
        if (!ruta.toLowerCase(java.util.Locale.ROOT).endsWith(".xlsx")) {
            ruta += ".xlsx";
        }

        String[] encabezados = {
            "NOMBRE PROVEEDOR", "NO. ORDEN", "FECHA DE CREACI\u00d3N", "FECHA DE PAGO PROMETIDA",
            "FECHA DE RECEPCI\u00d3N PROMETIDA", "COTIZACI\u00d3N", "SOLICITANTE", "TOTAL", "ABONADO",
            "PEND. X PAGAR", "PROD. X RECIBIR", "ESTATUS"
        };
        try (Workbook libro = new XSSFWorkbook()) {
            Sheet hoja = libro.createSheet("Historial");
            org.apache.poi.ss.usermodel.CellStyle estiloEncabezado = libro.createCellStyle();
            org.apache.poi.ss.usermodel.Font fuenteEncabezado = libro.createFont();
            fuenteEncabezado.setBold(true);
            fuenteEncabezado.setColor(org.apache.poi.ss.usermodel.IndexedColors.WHITE.getIndex());
            estiloEncabezado.setFont(fuenteEncabezado);
            estiloEncabezado.setFillForegroundColor(org.apache.poi.ss.usermodel.IndexedColors.BLUE.getIndex());
            estiloEncabezado.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);

            Row filaEncabezado = hoja.createRow(0);
            for (int columna = 0; columna < encabezados.length; columna++) {
                Cell celda = filaEncabezado.createCell(columna);
                celda.setCellValue(encabezados[columna]);
                celda.setCellStyle(estiloEncabezado);
            }

            int numeroFila = 1;
            for (Object[] fila : filas) {
                Row rowExcel = hoja.createRow(numeroFila++);
                int[] indicesOrigen = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13};
                for (int columna = 0; columna < indicesOrigen.length; columna++) {
                    rowExcel.createCell(columna).setCellValue(String.valueOf(fila[indicesOrigen[columna]]));
                }
            }
            for (int columna = 0; columna < encabezados.length; columna++) {
                hoja.autoSizeColumn(columna);
            }

            try (FileOutputStream fos = new FileOutputStream(ruta)) {
                libro.write(fos);
            }
            JOptionPane.showMessageDialog(this, "Historial exportado correctamente a Excel.");
        } catch (Exception ex) {
            mostrarError("No se pudo exportar el historial a Excel.", ex);
        }
    }//exportarHistorialExcel

    private void exportarHistorialPdf(java.util.List<Object[]> filas, String idProveedor) {
        javax.swing.JFileChooser selector = new javax.swing.JFileChooser();
        selector.setDialogTitle("Exportar historial a PDF");
        String nombreSugerido = "HistorialCompras" + (idProveedor.isEmpty() ? "" : "_" + idProveedor) + ".pdf";
        selector.setSelectedFile(new File(System.getProperty("user.home") + "\\Desktop\\" + nombreSugerido));
        selector.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("PDF (*.pdf)", "pdf"));
        if (selector.showSaveDialog(this) != javax.swing.JFileChooser.APPROVE_OPTION) {
            return;
        }
        String ruta = selector.getSelectedFile().getAbsolutePath();
        if (!ruta.toLowerCase(java.util.Locale.ROOT).endsWith(".pdf")) {
            ruta += ".pdf";
        }

        String[] encabezados = {
            "NOMBRE PROVEEDOR", "NO. ORDEN", "FECHA DE CREACI\u00d3N", "FECHA DE PAGO PROMETIDA",
            "FECHA DE RECEPCI\u00d3N PROMETIDA", "COTIZACI\u00d3N", "SOLICITANTE", "TOTAL", "ABONADO",
            "PEND. X PAGAR", "PROD. X RECIBIR", "ESTATUS"
        };
        try {
            com.itextpdf.kernel.pdf.PdfWriter escritor = new com.itextpdf.kernel.pdf.PdfWriter(ruta);
            com.itextpdf.kernel.pdf.PdfDocument documentoPdf = new com.itextpdf.kernel.pdf.PdfDocument(escritor);
            com.itextpdf.layout.Document documento = new com.itextpdf.layout.Document(documentoPdf,
                    com.itextpdf.kernel.geom.PageSize.A4.rotate());
            documento.setMargins(24, 24, 24, 24);

            com.itextpdf.kernel.font.PdfFont negrita = com.itextpdf.kernel.font.PdfFontFactory.createFont(
                    com.itextpdf.io.font.constants.StandardFonts.HELVETICA_BOLD);
            com.itextpdf.kernel.font.PdfFont normal = com.itextpdf.kernel.font.PdfFontFactory.createFont(
                    com.itextpdf.io.font.constants.StandardFonts.HELVETICA);
            com.itextpdf.kernel.colors.DeviceRgb azulPdf = new com.itextpdf.kernel.colors.DeviceRgb(31, 112, 184);

            documento.add(new com.itextpdf.layout.element.Paragraph(
                    "Historial de \u00d3rdenes de Compra" + (idProveedor.isEmpty() ? "" : " - " + idProveedor))
                    .setFont(negrita).setFontSize(16).setFontColor(azulPdf));

            com.itextpdf.layout.element.Table tablaPdf = new com.itextpdf.layout.element.Table(
                    com.itextpdf.layout.properties.UnitValue.createPercentArray(encabezados.length)).useAllAvailableWidth();
            for (String encabezado : encabezados) {
                tablaPdf.addHeaderCell(new com.itextpdf.layout.element.Cell()
                        .add(new com.itextpdf.layout.element.Paragraph(encabezado).setFont(negrita).setFontSize(8)
                                .setFontColor(com.itextpdf.kernel.colors.ColorConstants.WHITE))
                        .setBackgroundColor(azulPdf)
                        .setTextAlignment(com.itextpdf.layout.properties.TextAlignment.CENTER));
            }
            int[] indicesOrigen = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13};
            for (Object[] fila : filas) {
                for (int indice : indicesOrigen) {
                    tablaPdf.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new com.itextpdf.layout.element.Paragraph(String.valueOf(fila[indice])).setFont(normal).setFontSize(8))
                            .setTextAlignment(com.itextpdf.layout.properties.TextAlignment.CENTER));
                }
            }
            documento.add(tablaPdf);
            documento.close();
            JOptionPane.showMessageDialog(this, "Historial exportado correctamente a PDF.");
        } catch (Exception ex) {
            mostrarError("No se pudo exportar el historial a PDF.", ex);
        }
    }//exportarHistorialPdf

    private void prepararBotonIcono(javax.swing.JButton boton, javax.swing.Icon icono) {
        boton.setUI(new BotonPlanoUI());
        boton.setText(null);
        boton.setIcon(icono);
        boton.setBackground(java.awt.Color.WHITE);
        boton.setFocusPainted(false);
        boton.setContentAreaFilled(false);
        boton.setOpaque(false);
        boton.setRolloverEnabled(true);
        boton.setBorderPainted(true);
        boton.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                new BordeRedondeado(8, COLOR_BORDE_CAMPO, 1.2f),
                javax.swing.BorderFactory.createEmptyBorder(6, 6, 6, 6)));
        boton.setPreferredSize(new java.awt.Dimension(32, 32));
        boton.setCursor(java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR));
    }//prepararBotonIcono

    private static TarjetaResumen crearTarjetaResumen(String titulo, String valorInicial,
            javax.swing.Icon icono, java.awt.Color colorFondoIcono) {
        RoundedPanel tarjeta = new RoundedPanel(12, COLOR_AZUL_PALIDO);
        tarjeta.setLayout(new java.awt.BorderLayout(12, 0));
        tarjeta.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO),
                javax.swing.BorderFactory.createEmptyBorder(12, 16, 12, 16)));

        tarjeta.add(crearInsigniaCircular(icono, colorFondoIcono, 42), java.awt.BorderLayout.WEST);

        javax.swing.JPanel textos = new javax.swing.JPanel();
        textos.setOpaque(false);
        textos.setLayout(new javax.swing.BoxLayout(textos, javax.swing.BoxLayout.Y_AXIS));
        javax.swing.JLabel etiquetaTitulo = new javax.swing.JLabel(titulo);
        etiquetaTitulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 11));
        etiquetaTitulo.setForeground(COLOR_TEXTO_GRIS);
        javax.swing.JLabel etiquetaValor = new javax.swing.JLabel(valorInicial);
        etiquetaValor.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 18));
        etiquetaValor.setForeground(COLOR_AZUL_OSCURO);
        textos.add(etiquetaTitulo);
        textos.add(etiquetaValor);
        tarjeta.add(textos, java.awt.BorderLayout.CENTER);
        return new TarjetaResumen(tarjeta, etiquetaValor);
    }//crearTarjetaResumen

    private static javax.swing.JComponent crearInsigniaCircular(javax.swing.Icon icono, java.awt.Color colorFondo, int diametro) {
        javax.swing.JLabel insignia = new javax.swing.JLabel(icono) {
            @Override
            protected void paintComponent(java.awt.Graphics g) {
                java.awt.Graphics2D g2 = (java.awt.Graphics2D) g.create();
                g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING, java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
                g2.setColor(colorFondo);
                g2.fillOval(0, 0, getWidth() - 1, getHeight() - 1);
                g2.dispose();
                super.paintComponent(g);
            }
        };
        insignia.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        insignia.setVerticalAlignment(javax.swing.SwingConstants.CENTER);
        insignia.setPreferredSize(new java.awt.Dimension(diametro, diametro));
        insignia.setOpaque(false);
        return insignia;
    }//crearInsigniaCircular

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

    private void registrarAbono(String noOrden, javax.swing.JTable tabla) {
        int filaVista = tabla.getSelectedRow();
        if (filaVista >= 0) {
            int filaSeleccionadaModelo = tabla.convertRowIndexToModel(filaVista);
            String ordenSeleccionada = String.valueOf(tabla.getModel()
                    .getValueAt(filaSeleccionadaModelo, 2)).trim();
            if (!noOrden.equalsIgnoreCase(ordenSeleccionada)) {
                filaVista = -1;
            }
        }
        if (filaVista < 0) {
            for (int fila = 0; fila < tabla.getRowCount(); fila++) {
                int filaModelo = tabla.convertRowIndexToModel(fila);
                if (noOrden.equalsIgnoreCase(String.valueOf(
                        tabla.getModel().getValueAt(filaModelo, 2)).trim())) {
                    filaVista = fila;
                    tabla.setRowSelectionInterval(fila, fila);
                    break;
                }
            }
            if (filaVista < 0) {
                JOptionPane.showMessageDialog(this,
                        "No se encontró la orden seleccionada en la tabla.");
                return;
            }
        }

        int filaModelo = tabla.convertRowIndexToModel(filaVista);
        javax.swing.table.TableModel modelo = tabla.getModel();
        String totalTexto = String.valueOf(modelo.getValueAt(filaModelo, 8)).trim();
        String abonadoTexto = String.valueOf(modelo.getValueAt(filaModelo, 9)).trim();

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
        String nuevoAbonadoTexto = String.format("%s %,.2f", simbolo, nuevoAbonado);
        modelo.setValueAt(nuevoAbonadoTexto, filaModelo, 9);
        modelo.setValueAt(String.format("%s %,.2f", simbolo, Math.max(0, totalNumero - nuevoAbonado)), filaModelo, 10);
        modelo.setValueAt(calcularEstatusOrden(totalTexto, nuevoAbonadoTexto,
            obtenerEstadoRecepcionOrden(noOrden)), filaModelo, 13);
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

        String textoNormalizado = texto.toString()
                .replace('\u00A0', ' ')
                .replaceAll("\\s+", " ");
        java.util.regex.Pattern importeEtiquetado = java.util.regex.Pattern.compile(
                "(?i)(?:importe|monto|total|cantidad)(?:\\s+en\\s+texto)?"
                + "(?:\\s+(?:pagad[oa]|de\\s+pago|a\\s+pagar))?"
                + "\\s*[:=]?\\s*(?:MXN|M\\.?N\\.?)?\\s*\\$?\\s*"
                + "([0-9][0-9.,\\s]*[0-9]|[0-9])");
        java.util.regex.Matcher coincidencia = importeEtiquetado.matcher(textoNormalizado);
        while (coincidencia.find()) {
            double importe = parseMontoDocumento(coincidencia.group(1));
            if (importe > 0) {
                return importe;
            }
        }

        java.util.regex.Matcher importeMoneda = java.util.regex.Pattern
                .compile("(?i)(?:MXN|M\\.?N\\.?)?\\s*\\$\\s*([0-9][0-9.,\\s]*[0-9]|[0-9])")
                .matcher(textoNormalizado);
        double mejorImporte = 0;
        while (importeMoneda.find()) {
            double importe = parseMontoDocumento(importeMoneda.group(1));
            if (importe > mejorImporte && importe <= pendiente + 0.001) {
                mejorImporte = importe;
            }
        }
        if (mejorImporte <= 0) {
            java.util.regex.Matcher importeDecimal = java.util.regex.Pattern
                    .compile("(?<![\\d])([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2})|[0-9]+[.,][0-9]{2})(?![\\d])")
                    .matcher(textoNormalizado);
            while (importeDecimal.find()) {
                double importe = parseMontoDocumento(importeDecimal.group(1));
                if (importe > mejorImporte && importe <= pendiente + 0.001) {
                    mejorImporte = importe;
                }
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
                    if (noOrden.equalsIgnoreCase(String.valueOf(modelo.getValueAt(filaModelo, 2)).trim())) {
                        double cantidadPendiente = "SI".equalsIgnoreCase(estadoRecepcion)
                                ? 0.0 : obtenerProductosPendientesOrden(noOrden);
                        modelo.setValueAt(abonado, filaModelo, 9);
                        modelo.setValueAt(calcularPendientePago(total, abonado), filaModelo, 10);
                        modelo.setValueAt(diasFaltantesTexto, filaModelo, 11);
                        modelo.setValueAt(formatearCantidadInventario(cantidadPendiente), filaModelo, 12);
                        modelo.setValueAt(calcularEstatusOrden(total, abonado, estadoRecepcion), filaModelo, 13);
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
        final int columnaNoOrden = 2;
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
        final boolean ordenFinalizada = "FINALIZADO".equalsIgnoreCase(datos.get("ESTATUS"));

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
                return !ordenFinalizada && columna == 6
                        && parseNumeroInventario(String.valueOf(getValueAt(fila, 5))) > 0;
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
        scrollProductos.setPreferredSize(new java.awt.Dimension(900, 520));
        panelProductos.add(scrollProductos, java.awt.BorderLayout.CENTER);

        double totalProductos = 0;
        double productosRecibidos = 0;
        double productosPendientes = 0;
        for (int fila = 0; fila < modeloProductos.getRowCount(); fila++) {
            totalProductos += parseNumeroInventario(String.valueOf(modeloProductos.getValueAt(fila, 3)));
            productosRecibidos += parseNumeroInventario(String.valueOf(modeloProductos.getValueAt(fila, 4)));
            productosPendientes += parseNumeroInventario(String.valueOf(modeloProductos.getValueAt(fila, 5)));
        }
        javax.swing.JPanel resumenRecepcion = new javax.swing.JPanel(new java.awt.BorderLayout(10, 8));
        resumenRecepcion.setBorder(javax.swing.BorderFactory.createEmptyBorder(8, 0, 0, 0));

        javax.swing.JPanel tarjetasRecepcion = new javax.swing.JPanel(new java.awt.GridLayout(1, 4, 8, 0));
        tarjetasRecepcion.add(crearTarjetaRecepcion("Productos recibidos",
                formatearCantidadInventario(productosRecibidos) + " de "
                        + formatearCantidadInventario(totalProductos),
                COLOR_AZUL, iconoCarrito(COLOR_AZUL, 20)).panel);
        tarjetasRecepcion.add(crearTarjetaRecepcion("Pendientes",
                formatearCantidadInventario(productosPendientes) + " por recibir",
                COLOR_AZUL, iconoDocumentoAdjunto(COLOR_AZUL)).panel);
        TarjetaResumen tarjetaIncidencias = crearTarjetaRecepcion("Incidencias", "0",
                new java.awt.Color(230, 160, 0), iconoMonedas(new java.awt.Color(230, 160, 0)));
        tarjetasRecepcion.add(tarjetaIncidencias.panel);
        TarjetaResumen tarjetaDevoluciones = crearTarjetaRecepcion("Devoluciones", "0",
                COLOR_ROJO, iconoCerrarX(COLOR_ROJO));
        tarjetasRecepcion.add(tarjetaDevoluciones.panel);

        javax.swing.JPanel entrega = new javax.swing.JPanel(new java.awt.GridLayout(2, 1, 4, 2));
        entrega.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO),
                javax.swing.BorderFactory.createEmptyBorder(6, 10, 6, 10)));
        entrega.setBackground(new java.awt.Color(248, 251, 255));
        javax.swing.JLabel fechaPrometida = new javax.swing.JLabel(
                "Fecha prometida: " + datos.getOrDefault("FECHA DE ENTREGA", "-"));
        fechaPrometida.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 11));
        java.time.LocalDate fechaEsperada = obtenerFechaLocal(datos.get("FECHA DE ENTREGA"));
        long diasRetraso = fechaEsperada == null ? 0
                : Math.max(0, java.time.temporal.ChronoUnit.DAYS.between(
                        fechaEsperada, java.time.LocalDate.now()));
        javax.swing.JLabel diasRetrasoEtiqueta = new javax.swing.JLabel(
                "Días de retraso: " + diasRetraso);
        diasRetrasoEtiqueta.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
        diasRetrasoEtiqueta.setForeground(diasRetraso > 0 ? COLOR_ROJO : COLOR_VERDE);
        entrega.add(fechaPrometida);
        entrega.add(diasRetrasoEtiqueta);

        javax.swing.JPanel accionesRecepcion = new javax.swing.JPanel(
                new java.awt.GridLayout(2, 1, 0, 6));
        accionesRecepcion.setOpaque(false);
        javax.swing.JButton botonIncidencia = new javax.swing.JButton("REGISTRAR INCIDENCIA");
        prepararBoton(botonIncidencia, "REGISTRAR INCIDENCIA",
                iconoMonedas(java.awt.Color.WHITE), java.awt.Color.WHITE,
                COLOR_AZUL, COLOR_AZUL);
        javax.swing.JButton botonDevolucion = new javax.swing.JButton("REGISTRAR DEVOLUCIÓN");
        prepararBoton(botonDevolucion, "REGISTRAR DEVOLUCIÓN",
                iconoCerrarX(java.awt.Color.WHITE), java.awt.Color.WHITE,
                COLOR_AZUL, COLOR_AZUL);
        botonIncidencia.setEnabled(!ordenFinalizada);
        botonDevolucion.setEnabled(!ordenFinalizada);
        accionesRecepcion.add(botonIncidencia);
        accionesRecepcion.add(botonDevolucion);

        javax.swing.JPanel encabezadoResumen = new javax.swing.JPanel(new java.awt.BorderLayout(8, 0));
        encabezadoResumen.setOpaque(false);
        encabezadoResumen.add(tarjetasRecepcion, java.awt.BorderLayout.CENTER);
        encabezadoResumen.add(entrega, java.awt.BorderLayout.EAST);
        resumenRecepcion.add(encabezadoResumen, java.awt.BorderLayout.NORTH);

        String[] columnasIncidencias = {"TIPO", "PRODUCTO", "CANTIDAD", "MOTIVO", "ESTATUS", "ACCIONES"};
        javax.swing.table.DefaultTableModel modeloIncidencias =
                new javax.swing.table.DefaultTableModel(columnasIncidencias, 0) {
                    @Override
                    public boolean isCellEditable(int fila, int columna) {
                        return false;
                    }
                };
        cargarIncidenciasEnTabla(noOrden, modeloIncidencias, tarjetaIncidencias);
        cargarIncidenciasEnTabla(noOrden, modeloIncidencias, tarjetaIncidencias);
        cargarDevolucionesEnTabla(noOrden, modeloIncidencias, tarjetaDevoluciones);
        botonIncidencia.addActionListener(e -> mostrarFormularioIncidencia(
                    noOrden, datos, modeloProductos, modeloIncidencias, tarjetaIncidencias));
        botonDevolucion.addActionListener(e -> mostrarFormularioDevolucion(
                    noOrden, modeloProductos, modeloIncidencias, tarjetaIncidencias, tarjetaDevoluciones));
        javax.swing.JTable tablaIncidencias = new javax.swing.JTable(modeloIncidencias);
        tablaIncidencias.setRowHeight(24);
        tablaIncidencias.setFillsViewportHeight(true);
        tablaIncidencias.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        tablaIncidencias.getTableHeader().setFont(
                tablaIncidencias.getTableHeader().getFont().deriveFont(java.awt.Font.BOLD));
        tablaIncidencias.getColumnModel().getColumn(5).setCellRenderer(
                new javax.swing.table.DefaultTableCellRenderer() {
                    @Override
                    public java.awt.Component getTableCellRendererComponent(javax.swing.JTable tabla,
                            Object valor, boolean seleccionado, boolean foco, int fila, int columna) {
                        javax.swing.JButton boton = new javax.swing.JButton("ELIMINAR");
                        boton.setForeground(COLOR_ROJO);
                        boton.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 11));
                        boton.setFocusPainted(false);
                        boton.setOpaque(true);
                        boton.setBackground(java.awt.Color.WHITE);
                        return boton;
                    }
                });
        ajustarAnchoColumnasIncidencias(tablaIncidencias);
        tablaIncidencias.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mouseClicked(java.awt.event.MouseEvent evento) {
                int columna = tablaIncidencias.columnAtPoint(evento.getPoint());
                int fila = tablaIncidencias.rowAtPoint(evento.getPoint());
                if (columna != 5 || fila < 0) {
                    return;
                }
                int filaModelo = tablaIncidencias.convertRowIndexToModel(fila);
                int respuesta = JOptionPane.showConfirmDialog(tablaIncidencias,
                        "¿Deseas eliminar esta incidencia?",
                        "Confirmar eliminación", JOptionPane.YES_NO_OPTION,
                        JOptionPane.WARNING_MESSAGE);
                if (respuesta != JOptionPane.YES_OPTION) {
                    return;
                }
                try {
                    boolean esDevolucion = "DEVOLUCIÓN".equals(
                            String.valueOf(modeloIncidencias.getValueAt(filaModelo, 0)));
                    if (esDevolucion) {
                        eliminarDevolucion(noOrden,
                                String.valueOf(modeloIncidencias.getValueAt(filaModelo, 1)),
                                String.valueOf(modeloIncidencias.getValueAt(filaModelo, 2)),
                                String.valueOf(modeloIncidencias.getValueAt(filaModelo, 3)));
                    } else {
                        eliminarIncidencia(noOrden,
                                String.valueOf(modeloIncidencias.getValueAt(filaModelo, 1)),
                                String.valueOf(modeloIncidencias.getValueAt(filaModelo, 2)),
                                String.valueOf(modeloIncidencias.getValueAt(filaModelo, 0)),
                                String.valueOf(modeloIncidencias.getValueAt(filaModelo, 3)));
                    }
                    cargarIncidenciasEnTabla(noOrden, modeloIncidencias, tarjetaIncidencias);
                    cargarDevolucionesEnTabla(noOrden, modeloIncidencias, tarjetaDevoluciones);
                } catch (Exception ex) {
                    mostrarError("No se pudo eliminar la incidencia.", ex);
                }
            }
        });
        javax.swing.JPanel panelIncidencias = new javax.swing.JPanel(new java.awt.BorderLayout());
        panelIncidencias.setBorder(javax.swing.BorderFactory.createTitledBorder(
                "INCIDENCIAS Y DEVOLUCIONES DE ESTA ORDEN"));
        panelIncidencias.add(new javax.swing.JScrollPane(tablaIncidencias), java.awt.BorderLayout.CENTER);
        panelIncidencias.setPreferredSize(new java.awt.Dimension(860, 150));
        resumenRecepcion.add(panelIncidencias, java.awt.BorderLayout.CENTER);

        javax.swing.JPanel controlesRecepcion = new javax.swing.JPanel(new java.awt.BorderLayout(8, 0));
        controlesRecepcion.setOpaque(false);
        controlesRecepcion.add(accionesRecepcion, java.awt.BorderLayout.WEST);

        javax.swing.JButton botonRecibir = new javax.swing.JButton("RECIBIR MERCANCIA");
        botonRecibir.setBackground(new java.awt.Color(0, 95, 131));
        botonRecibir.setForeground(java.awt.Color.WHITE);
        botonRecibir.setEnabled(!ordenFinalizada
                && !ordenMercanciaRecibida(noOrden) && hayPendienteRecibir(modeloProductos));
        botonRecibir.addActionListener(e -> {
            if (recibirMercancia(noOrden, modeloProductos)) {
                botonRecibir.setEnabled(!ordenFinalizada
                        && !ordenMercanciaRecibida(noOrden) && hayPendienteRecibir(modeloProductos));
                actualizarFilaHistorial(tablaHistorial, noOrden);
            }
        });
        javax.swing.JPanel panelRecibir = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT));
        panelRecibir.add(botonRecibir);
        controlesRecepcion.add(panelRecibir, java.awt.BorderLayout.EAST);
        javax.swing.JPanel inferiorProductos = new javax.swing.JPanel(new java.awt.BorderLayout(0, 8));
        inferiorProductos.setOpaque(false);
        inferiorProductos.add(resumenRecepcion, java.awt.BorderLayout.CENTER);
        inferiorProductos.add(controlesRecepcion, java.awt.BorderLayout.SOUTH);
        panelProductos.add(inferiorProductos, java.awt.BorderLayout.SOUTH);

        javax.swing.JPanel panelContenido = new javax.swing.JPanel(new java.awt.BorderLayout(10, 10));
        panelContenido.add(panelDatos, java.awt.BorderLayout.NORTH);
        panelContenido.add(panelProductos, java.awt.BorderLayout.CENTER);
        panelContenido.setPreferredSize(new java.awt.Dimension(900, 850));

        javax.swing.JDialog dialogo = new javax.swing.JDialog(this, "Orden de Compra No. " + noOrden, true);
        dialogo.getContentPane().add(panelContenido);
        dialogo.setResizable(true);
        dialogo.pack();
        dialogo.setLocationRelativeTo(null);
        dialogo.setVisible(true);
    }//mostrarDetalleOrden

    private void mostrarFormularioIncidencia(String noOrden, java.util.Map<String, String> datos,
            javax.swing.table.DefaultTableModel modeloProductos,
            javax.swing.table.DefaultTableModel modeloIncidencias,
            TarjetaResumen tarjetaIncidencias) {
        javax.swing.JDialog dialogo = new javax.swing.JDialog(this,
                "Registrar incidencia - Orden No. " + noOrden, true);
        dialogo.setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);

        javax.swing.JPanel contenido = new javax.swing.JPanel(new java.awt.BorderLayout(0, 12));
        contenido.setBorder(javax.swing.BorderFactory.createEmptyBorder(14, 18, 14, 18));

        javax.swing.JLabel titulo = new javax.swing.JLabel("REGISTRAR INCIDENCIA");
        titulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 20));
        titulo.setForeground(COLOR_AZUL);
        contenido.add(titulo, java.awt.BorderLayout.NORTH);

        javax.swing.JPanel formulario = new javax.swing.JPanel(new java.awt.GridBagLayout());
        java.awt.GridBagConstraints restricciones = new java.awt.GridBagConstraints();
        restricciones.insets = new java.awt.Insets(5, 5, 5, 5);
        restricciones.fill = java.awt.GridBagConstraints.HORIZONTAL;
        restricciones.anchor = java.awt.GridBagConstraints.WEST;

        javax.swing.JComboBox<String> tipo = new javax.swing.JComboBox<>(
                new String[]{"Producto dañado", "Producto incorrecto", "Faltante", "Otro"});
        javax.swing.JComboBox<String> producto = new javax.swing.JComboBox<>();
        java.util.List<String[]> productos = new java.util.ArrayList<>();
        for (int fila = 0; fila < modeloProductos.getRowCount(); fila++) {
            String codigo = obtenerValorModelo(modeloProductos, fila, 0);
            String descripcion = obtenerValorModelo(modeloProductos, fila, 1);
            int pendiente = (int) parseNumeroInventario(obtenerValorModelo(modeloProductos, fila, 5));
            if (!codigo.isEmpty() && pendiente > 0) {
                productos.add(new String[]{codigo, descripcion, String.valueOf(pendiente)});
                producto.addItem(codigo + " - " + descripcion);
            }
        }
        javax.swing.JSpinner cantidad = new javax.swing.JSpinner(
                new javax.swing.SpinnerNumberModel(1, 1, 999999, 1));
        javax.swing.JTextField fecha = new javax.swing.JTextField(
                java.time.LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        javax.swing.JTextArea descripcion = new javax.swing.JTextArea(5, 45);
        descripcion.setLineWrap(true);
        descripcion.setWrapStyleWord(true);
        javax.swing.JScrollPane scrollDescripcion = new javax.swing.JScrollPane(descripcion);
        javax.swing.JButton guardar = new javax.swing.JButton("GUARDAR INCIDENCIA");
        guardar.setEnabled(!productos.isEmpty());

        agregarCampoIncidencia(formulario, restricciones, 0, "Tipo de incidencia:", tipo);
        agregarCampoIncidencia(formulario, restricciones, 1, "Fecha de incidencia:", fecha);
        agregarCampoIncidencia(formulario, restricciones, 2, "Producto:", producto);
        agregarCampoIncidencia(formulario, restricciones, 3, "Cantidad con incidencia:", cantidad);
        agregarCampoIncidencia(formulario, restricciones, 4, "Descripción de la incidencia:", scrollDescripcion);
        contenido.add(formulario, java.awt.BorderLayout.CENTER);

        javax.swing.JPanel botones = new javax.swing.JPanel(
                new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 8, 0));
        javax.swing.JButton cancelar = new javax.swing.JButton("CANCELAR");
        prepararBoton(guardar, "GUARDAR INCIDENCIA", null, java.awt.Color.WHITE, COLOR_AZUL, COLOR_AZUL);
        producto.addActionListener(e -> {
            if (producto.getSelectedIndex() >= 0) {
                int maximo = Integer.parseInt(productos.get(producto.getSelectedIndex())[2]);
                cantidad.setModel(new javax.swing.SpinnerNumberModel(1, 1, maximo, 1));
            }
        });
        if (productos.isEmpty()) {
            javax.swing.JLabel aviso = new javax.swing.JLabel(
                    "No hay productos pendientes para registrar una incidencia.");
            aviso.setForeground(COLOR_ROJO);
            formulario.add(aviso, new java.awt.GridBagConstraints());
        } else {
            producto.setSelectedIndex(0);
        }
        cancelar.addActionListener(e -> dialogo.dispose());
        guardar.addActionListener(e -> {
            if (producto.getSelectedIndex() < 0) {
                JOptionPane.showMessageDialog(dialogo, "Selecciona un producto.", "Datos incompletos",
                        JOptionPane.WARNING_MESSAGE);
                return;
            }
            if (descripcion.getText().trim().isEmpty()) {
                JOptionPane.showMessageDialog(dialogo, "Escribe la descripción de la incidencia.",
                        "Datos incompletos", JOptionPane.WARNING_MESSAGE);
                descripcion.requestFocusInWindow();
                return;
            }
            String[] productoSeleccionado = productos.get(producto.getSelectedIndex());
            int cantidadIncidencia = ((Number) cantidad.getValue()).intValue();
            int cantidadPendiente = Integer.parseInt(productoSeleccionado[2]);
            if (cantidadIncidencia > cantidadPendiente) {
                JOptionPane.showMessageDialog(dialogo,
                        "La cantidad no puede ser mayor a la cantidad pendiente (" + cantidadPendiente + ").",
                        "Cantidad inválida", JOptionPane.WARNING_MESSAGE);
                cantidad.requestFocusInWindow();
                return;
            }
            try {
                guardarIncidencia(noOrden, fecha.getText().trim(), productoSeleccionado[0],
                        productoSeleccionado[1], cantidadIncidencia, String.valueOf(tipo.getSelectedItem()),
                        descripcion.getText().trim());
                cargarIncidenciasEnTabla(noOrden, modeloIncidencias, tarjetaIncidencias);
                dialogo.dispose();
                JOptionPane.showMessageDialog(this, "La incidencia se guardó correctamente.",
                        "Incidencia registrada", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                mostrarError("No se pudo guardar la incidencia.", ex);
            }
        });
        botones.add(cancelar);
        botones.add(guardar);
        contenido.add(botones, java.awt.BorderLayout.SOUTH);

        dialogo.setContentPane(contenido);
        dialogo.setSize(720, 520);
        dialogo.setResizable(false);
        dialogo.setLocationRelativeTo(null);
        dialogo.setVisible(true);
    }//mostrarFormularioIncidencia

    private void agregarCampoIncidencia(javax.swing.JPanel formulario,
            java.awt.GridBagConstraints restricciones, int fila, String etiqueta,
            java.awt.Component campo) {
        restricciones.gridx = 0;
        restricciones.gridy = fila;
        restricciones.weightx = 0;
        formulario.add(new javax.swing.JLabel(etiqueta), (java.awt.GridBagConstraints) restricciones.clone());
        restricciones.gridx = 1;
        restricciones.weightx = 1;
        formulario.add(campo, (java.awt.GridBagConstraints) restricciones.clone());
    }//agregarCampoIncidencia

    private void guardarIncidencia(String noOrden, String fecha, String codigoProducto,
            String producto, int cantidad, String tipo, String descripcion) throws IOException {
        File archivo = AppConfig.resolveBaseFile("REGISTROINCIDENCIAS.xlsx");
        Workbook libro;
        if (archivo.exists() && archivo.length() > 0) {
            try (FileInputStream fis = new FileInputStream(archivo)) {
                libro = new XSSFWorkbook(fis);
            }
        } else {
            libro = new XSSFWorkbook();
        }
        try (libro) {
            Sheet hoja = libro.getNumberOfSheets() == 0
                    ? libro.createSheet("INCIDENCIAS") : libro.getSheetAt(0);
            Row encabezado = hoja.getRow(0);
            if (encabezado == null) {
                encabezado = hoja.createRow(0);
            }
            for (int columna = 0; columna < ENCABEZADOS_INCIDENCIAS.length; columna++) {
                encabezado.getCell(columna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        .setCellValue(ENCABEZADOS_INCIDENCIAS[columna]);
            }
            Row fila = hoja.createRow(hoja.getLastRowNum() + 1);
            String usuario = INICIARSESION.usuarioActual == null ? "" : INICIARSESION.usuarioActual.trim();
            String nombre = INICIARSESION.nombreUsuarioActual == null
                    ? "" : INICIARSESION.nombreUsuarioActual.trim();
            String registradoPor = nombre.isEmpty() ? usuario : nombre;
            String[] valores = {noOrden, fecha, codigoProducto, producto, String.valueOf(cantidad),
                tipo, descripcion, registradoPor, "ABIERTA"};
            for (int columna = 0; columna < valores.length; columna++) {
                fila.createCell(columna).setCellValue(valores[columna]);
            }
            try (FileOutputStream fos = new FileOutputStream(archivo)) {
                libro.write(fos);
            }
        }
    }//guardarIncidencia

    private void cargarIncidenciasEnTabla(String noOrden,
            javax.swing.table.DefaultTableModel modelo, TarjetaResumen tarjeta) {
        modelo.setRowCount(0);
        File archivo = AppConfig.resolveBaseFile("REGISTROINCIDENCIAS.xlsx");
        if (!archivo.exists()) {
            tarjeta.etiquetaValor.setText("0");
            return;
        }
        try (FileInputStream fis = new FileInputStream(archivo);
                Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
                Row registro = hoja.getRow(fila);
                if (registro == null || !noOrden.equalsIgnoreCase(leerCelda(registro, 0))) {
                    continue;
                }
                modelo.addRow(new Object[]{leerCelda(registro, 5), leerCelda(registro, 3),
                    leerCelda(registro, 4), leerCelda(registro, 6), leerCelda(registro, 8),
                    "ELIMINAR"});
            }
            tarjeta.etiquetaValor.setText(String.valueOf(modelo.getRowCount()));
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING,
                    "No se pudieron cargar las incidencias de la orden " + noOrden, ex);
            tarjeta.etiquetaValor.setText("0");
        }
    }//cargarIncidenciasEnTabla

    private void eliminarIncidencia(String noOrden, String producto, String cantidad,
            String tipo, String descripcion) throws IOException {
        File archivo = AppConfig.resolveBaseFile("REGISTROINCIDENCIAS.xlsx");
        if (!archivo.exists()) {
            return;
        }
        try (FileInputStream fis = new FileInputStream(archivo);
                Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                Row registro = hoja.getRow(indice);
                if (registro == null
                        || !noOrden.equalsIgnoreCase(leerCelda(registro, 0))
                        || !producto.equals(leerCelda(registro, 3))
                        || !cantidad.equals(leerCelda(registro, 4))
                        || !tipo.equals(leerCelda(registro, 5))
                        || !descripcion.equals(leerCelda(registro, 6))) {
                    continue;
                }
                int ultimaFila = hoja.getLastRowNum();
                hoja.removeRow(registro);
                if (indice < ultimaFila) {
                    hoja.shiftRows(indice + 1, ultimaFila, -1);
                }
                break;
            }
            try (FileOutputStream fos = new FileOutputStream(archivo)) {
                libro.write(fos);
            }
        }
    }//eliminarIncidencia

    private void mostrarFormularioDevolucion(String noOrden,
            javax.swing.table.DefaultTableModel modeloProductos,
            javax.swing.table.DefaultTableModel modeloIncidencias,
            TarjetaResumen tarjetaIncidencias,
            TarjetaResumen tarjetaDevoluciones) {
        javax.swing.JDialog dialogo = new javax.swing.JDialog(this,
                "Registrar devolución - Orden No. " + noOrden, true);
        javax.swing.JPanel formulario = new javax.swing.JPanel(new java.awt.GridBagLayout());
        formulario.setBorder(javax.swing.BorderFactory.createEmptyBorder(16, 18, 8, 18));
        java.awt.GridBagConstraints restricciones = new java.awt.GridBagConstraints();
        restricciones.insets = new java.awt.Insets(5, 5, 5, 5);
        restricciones.fill = java.awt.GridBagConstraints.HORIZONTAL;

        javax.swing.JComboBox<String> producto = new javax.swing.JComboBox<>();
        java.util.List<String[]> productos = new java.util.ArrayList<>();
        for (int fila = 0; fila < modeloProductos.getRowCount(); fila++) {
            String codigo = obtenerValorModelo(modeloProductos, fila, 0);
            String descripcionProducto = obtenerValorModelo(modeloProductos, fila, 1);
            int recibido = (int) parseNumeroInventario(obtenerValorModelo(modeloProductos, fila, 4));
            if (!codigo.isEmpty() && recibido > 0) {
                productos.add(new String[]{codigo, descripcionProducto, String.valueOf(recibido)});
                producto.addItem(codigo + " - " + descripcionProducto);
            }
        }
        javax.swing.JTextField fecha = new javax.swing.JTextField(
                java.time.LocalDate.now().format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        javax.swing.JSpinner cantidad = new javax.swing.JSpinner(
                new javax.swing.SpinnerNumberModel(1, 1, productos.isEmpty() ? 1
                        : Integer.parseInt(productos.get(0)[2]), 1));
        javax.swing.JComboBox<String> tipo = new javax.swing.JComboBox<>(
                new String[]{"Producto defectuoso", "Producto incorrecto", "Excedente", "Otro"});
        javax.swing.JComboBox<String> condicion = new javax.swing.JComboBox<>(
                new String[]{"Con empaque", "Sin empaque", "Dañado"});
        javax.swing.JComboBox<String> accion = new javax.swing.JComboBox<>(
                new String[]{"Reposición", "Reembolso", "Nota de crédito"});
        javax.swing.JTextField guia = new javax.swing.JTextField();
        javax.swing.JTextArea motivo = new javax.swing.JTextArea(4, 42);
        motivo.setLineWrap(true);
        motivo.setWrapStyleWord(true);

        agregarCampoIncidencia(formulario, restricciones, 0, "Producto:", producto);
        agregarCampoIncidencia(formulario, restricciones, 1, "Fecha de devolución:", fecha);
        agregarCampoIncidencia(formulario, restricciones, 2, "Cantidad a devolver:", cantidad);
        agregarCampoIncidencia(formulario, restricciones, 3, "Tipo de devolución:", tipo);
        agregarCampoIncidencia(formulario, restricciones, 4, "Condición del producto:", condicion);
        agregarCampoIncidencia(formulario, restricciones, 5, "Número de guía (opcional):", guia);
        agregarCampoIncidencia(formulario, restricciones, 6, "Acción solicitada:", accion);
        agregarCampoIncidencia(formulario, restricciones, 7, "Motivo detallado:", new javax.swing.JScrollPane(motivo));

        javax.swing.JButton guardar = new javax.swing.JButton("GUARDAR DEVOLUCIÓN");
        guardar.setEnabled(!productos.isEmpty());
        producto.addActionListener(e -> {
            if (producto.getSelectedIndex() >= 0) {
                int maximo = Integer.parseInt(productos.get(producto.getSelectedIndex())[2]);
                cantidad.setModel(new javax.swing.SpinnerNumberModel(1, 1, maximo, 1));
            }
        });
        javax.swing.JPanel botones = new javax.swing.JPanel(
                new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT));
        javax.swing.JButton cancelar = new javax.swing.JButton("CANCELAR");
        prepararBoton(guardar, "GUARDAR DEVOLUCIÓN", null, java.awt.Color.WHITE, COLOR_AZUL, COLOR_AZUL);
        cancelar.addActionListener(e -> dialogo.dispose());
        guardar.addActionListener(e -> {
            if (producto.getSelectedIndex() < 0) {
                JOptionPane.showMessageDialog(dialogo, "No hay productos recibidos disponibles para devolver.",
                        "Datos incompletos", JOptionPane.WARNING_MESSAGE);
                return;
            }
            if (motivo.getText().trim().isEmpty()) {
                JOptionPane.showMessageDialog(dialogo, "Escribe el motivo de la devolución.",
                        "Datos incompletos", JOptionPane.WARNING_MESSAGE);
                return;
            }
            String[] productoSeleccionado = productos.get(producto.getSelectedIndex());
            int cantidadDevolver = ((Number) cantidad.getValue()).intValue();
            int maximo = Integer.parseInt(productoSeleccionado[2]);
            if (cantidadDevolver > maximo) {
                JOptionPane.showMessageDialog(dialogo,
                        "La cantidad no puede ser mayor a la cantidad recibida (" + maximo + ").",
                        "Cantidad inválida", JOptionPane.WARNING_MESSAGE);
                return;
            }
            try {
                guardarDevolucion(noOrden, fecha.getText().trim(), productoSeleccionado[0],
                        productoSeleccionado[1], cantidadDevolver, String.valueOf(tipo.getSelectedItem()),
                        String.valueOf(condicion.getSelectedItem()), guia.getText().trim(),
                        String.valueOf(accion.getSelectedItem()), motivo.getText().trim());
                cargarDevolucionesEnTabla(noOrden, modeloIncidencias, tarjetaDevoluciones);
                dialogo.dispose();
                JOptionPane.showMessageDialog(this, "La devolución se guardó correctamente.",
                        "Devolución registrada", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                mostrarError("No se pudo guardar la devolución.", ex);
            }
        });
        botones.add(cancelar);
        botones.add(guardar);

        javax.swing.JPanel contenido = new javax.swing.JPanel(new java.awt.BorderLayout());
        javax.swing.JLabel titulo = new javax.swing.JLabel("REGISTRAR DEVOLUCIÓN");
        titulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 20));
        titulo.setForeground(COLOR_AZUL);
        contenido.add(titulo, java.awt.BorderLayout.NORTH);
        contenido.add(formulario, java.awt.BorderLayout.CENTER);
        contenido.add(botones, java.awt.BorderLayout.SOUTH);
        dialogo.setContentPane(contenido);
        dialogo.setSize(700, 590);
        dialogo.setResizable(false);
        dialogo.setLocationRelativeTo(null);
        dialogo.setVisible(true);
    }//mostrarFormularioDevolucion

    private void guardarDevolucion(String noOrden, String fecha, String codigoProducto,
            String producto, int cantidad, String tipo, String condicion, String guia,
            String accion, String motivo) throws IOException {
        File archivo = AppConfig.resolveBaseFile("REGISTRODEVOLUCIONES.xlsx");
        Workbook libro;
        if (archivo.exists() && archivo.length() > 0) {
            try (FileInputStream fis = new FileInputStream(archivo)) {
                libro = new XSSFWorkbook(fis);
            }
        } else {
            libro = new XSSFWorkbook();
        }
        try (libro) {
            Sheet hoja = libro.getNumberOfSheets() == 0
                    ? libro.createSheet("DEVOLUCIONES") : libro.getSheetAt(0);
            Row encabezado = hoja.getRow(0);
            if (encabezado == null) {
                encabezado = hoja.createRow(0);
            }
            for (int columna = 0; columna < ENCABEZADOS_DEVOLUCIONES.length; columna++) {
                encabezado.getCell(columna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        .setCellValue(ENCABEZADOS_DEVOLUCIONES[columna]);
            }
            Row fila = hoja.createRow(hoja.getLastRowNum() + 1);
            String usuario = INICIARSESION.usuarioActual == null ? "" : INICIARSESION.usuarioActual.trim();
            String nombre = INICIARSESION.nombreUsuarioActual == null
                    ? "" : INICIARSESION.nombreUsuarioActual.trim();
            String registradoPor = nombre.isEmpty() ? usuario : nombre;
            String[] valores = {noOrden, fecha, codigoProducto, producto, String.valueOf(cantidad),
                tipo, condicion, guia, accion, motivo, registradoPor, "PENDIENTE DE PROCESO"};
            for (int columna = 0; columna < valores.length; columna++) {
                fila.createCell(columna).setCellValue(valores[columna]);
            }
            try (FileOutputStream fos = new FileOutputStream(archivo)) {
                libro.write(fos);
            }
        }
    }//guardarDevolucion

    private void cargarDevolucionesEnTabla(String noOrden,
            javax.swing.table.DefaultTableModel modelo, TarjetaResumen tarjeta) {
        File archivo = AppConfig.resolveBaseFile("REGISTRODEVOLUCIONES.xlsx");
        if (!archivo.exists()) {
            tarjeta.etiquetaValor.setText("0");
            return;
        }
        int total = 0;
        try (FileInputStream fis = new FileInputStream(archivo);
                Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
                Row registro = hoja.getRow(fila);
                if (registro == null || !noOrden.equalsIgnoreCase(leerCelda(registro, 0))) {
                    continue;
                }
                modelo.addRow(new Object[]{"DEVOLUCIÓN", leerCelda(registro, 3),
                    leerCelda(registro, 4), leerCelda(registro, 9),
                    leerCelda(registro, 11), "ELIMINAR"});
                total++;
            }
            tarjeta.etiquetaValor.setText(String.valueOf(total));
        } catch (Exception ex) {
            logger.log(java.util.logging.Level.WARNING,
                    "No se pudieron cargar las devoluciones de la orden " + noOrden, ex);
            tarjeta.etiquetaValor.setText("0");
        }
    }//cargarDevolucionesEnTabla

    private void eliminarDevolucion(String noOrden, String producto, String cantidad,
            String motivo) throws IOException {
        File archivo = AppConfig.resolveBaseFile("REGISTRODEVOLUCIONES.xlsx");
        if (!archivo.exists()) {
            return;
        }
        try (FileInputStream fis = new FileInputStream(archivo);
                Workbook libro = new XSSFWorkbook(fis)) {
            Sheet hoja = libro.getSheetAt(0);
            for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                Row registro = hoja.getRow(indice);
                if (registro == null
                        || !noOrden.equalsIgnoreCase(leerCelda(registro, 0))
                        || !producto.equals(leerCelda(registro, 3))
                        || !cantidad.equals(leerCelda(registro, 4))
                        || !motivo.equals(leerCelda(registro, 9))) {
                    continue;
                }
                int ultimaFila = hoja.getLastRowNum();
                hoja.removeRow(registro);
                if (indice < ultimaFila) {
                    hoja.shiftRows(indice + 1, ultimaFila, -1);
                }
                break;
            }
            try (FileOutputStream fos = new FileOutputStream(archivo)) {
                libro.write(fos);
            }
        }
    }//eliminarDevolucion

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

    private TarjetaResumen crearTarjetaRecepcion(String titulo, String valor,
            java.awt.Color color, javax.swing.Icon icono) {
        javax.swing.JPanel tarjeta = new javax.swing.JPanel(new java.awt.BorderLayout(6, 0));
        tarjeta.setBackground(java.awt.Color.WHITE);
        tarjeta.setBorder(javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO));
        javax.swing.JLabel iconoEtiqueta = new javax.swing.JLabel(icono);
        iconoEtiqueta.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        iconoEtiqueta.setPreferredSize(new java.awt.Dimension(30, 30));
        javax.swing.JPanel textos = new javax.swing.JPanel();
        textos.setOpaque(false);
        textos.setLayout(new javax.swing.BoxLayout(textos, javax.swing.BoxLayout.Y_AXIS));
        javax.swing.JLabel etiquetaTitulo = new javax.swing.JLabel(titulo);
        etiquetaTitulo.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 10));
        etiquetaTitulo.setForeground(COLOR_TEXTO_GRIS);
        javax.swing.JLabel etiquetaValor = new javax.swing.JLabel(valor);
        etiquetaValor.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 16));
        etiquetaValor.setForeground(color);
        textos.add(etiquetaTitulo);
        textos.add(etiquetaValor);
        tarjeta.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO),
                javax.swing.BorderFactory.createEmptyBorder(5, 7, 5, 7)));
        tarjeta.add(iconoEtiqueta, java.awt.BorderLayout.WEST);
        tarjeta.add(textos, java.awt.BorderLayout.CENTER);
        return new TarjetaResumen(tarjeta, etiquetaValor);
    }//crearTarjetaRecepcion

    private java.time.LocalDate obtenerFechaLocal(String texto) {
        if (texto == null || texto.isBlank()) {
            return null;
        }
        try {
            return java.time.LocalDate.parse(texto.trim(),
                    java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        } catch (java.time.format.DateTimeParseException ex) {
            return null;
        }
    }//obtenerFechaLocal

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
            boolean formatoAnterior = "MODELO".equalsIgnoreCase(leerCelda(encabezado, 3))
                    && "CANTIDAD".equalsIgnoreCase(leerCelda(encabezado, 4));
            if (formatoAnterior) {
                for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
                    Row movimientoAnterior = hoja.getRow(fila);
                    if (movimientoAnterior == null || leerCelda(movimientoAnterior, 0).isEmpty()) {
                        continue;
                    }
                    String modeloAnterior = leerCelda(movimientoAnterior, 3);
                    String cantidadAnterior = leerCelda(movimientoAnterior, 4);
                    escribirCelda(movimientoAnterior, 3, "");
                    escribirCelda(movimientoAnterior, 4, "");
                    escribirCelda(movimientoAnterior, 5, "");
                    escribirCelda(movimientoAnterior, 6, modeloAnterior);
                    escribirCelda(movimientoAnterior, 7, cantidadAnterior);
                    escribirCelda(movimientoAnterior, 8, "");
                }
            }
            for (int columna = 0; columna < ENCABEZADOS_REGISTRO_RECEPCION.length; columna++) {
                encabezado.getCell(columna, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        .setCellValue(ENCABEZADOS_REGISTRO_RECEPCION[columna]);
            }

            String proveedor = obtenerProveedorOrden(noOrden);
            java.time.LocalDate fechaRecibido = java.time.LocalDate.now();
            String fechaTexto = fechaRecibido.format(java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy"));
            String horaTexto = java.time.LocalTime.now()
                    .format(java.time.format.DateTimeFormatter.ofPattern("HH:mm:ss"));
            String usuario = INICIARSESION.usuarioActual == null
                    ? "" : INICIARSESION.usuarioActual.trim();
            String nombreUsuario = INICIARSESION.nombreUsuarioActual == null
                    ? "" : INICIARSESION.nombreUsuarioActual.trim();
            boolean recepcionCompleta = true;
            for (int fila = 0; fila < modeloProductos.getRowCount(); fila++) {
                double pendiente = parseNumeroInventario(obtenerValorModelo(modeloProductos, fila, 5));
                double recibirAhora = parseNumeroInventario(obtenerValorModelo(modeloProductos, fila, 6));
                if (pendiente > recibirAhora) {
                    recepcionCompleta = false;
                    break;
                }
            }
            String tipoRecepcion = recepcionCompleta ? "COMPLETA" : "PARCIAL";
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
                escribirCelda(movimiento, 3, horaTexto);
                escribirCelda(movimiento, 4, usuario);
                escribirCelda(movimiento, 5, nombreUsuario);
                escribirCelda(movimiento, 6, obtenerValorModelo(modeloProductos, fila, 0));
                escribirCelda(movimiento, 7, formatearCantidadInventario(cantidadRecibida));
                escribirCelda(movimiento, 8, tipoRecepcion);
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

    private void ajustarAnchoColumnasHistorial(javax.swing.JTable tabla) {
        javax.swing.table.TableColumnModel modeloColumnas = tabla.getColumnModel();
        for (int columna = 0; columna < tabla.getColumnCount(); columna++) {
            javax.swing.table.TableColumn tableColumn = modeloColumnas.getColumn(columna);
            int ancho = "ESTATUS".equals(tabla.getColumnName(columna)) ? 230 : 100;
            tableColumn.setMinWidth(ancho);
            tableColumn.setPreferredWidth(ancho);
            tableColumn.setMaxWidth(ancho);
        }
    }//ajustarAnchoColumnasHistorial

    private void ajustarAnchoColumnasIncidencias(javax.swing.JTable tabla) {
        int[] minimos = {130, 220, 85, 220, 150, 130};
        int[] maximos = {190, 360, 110, 420, 240, 160};
        javax.swing.table.TableColumnModel modeloColumnas = tabla.getColumnModel();
        for (int columna = 0; columna < tabla.getColumnCount(); columna++) {
            int ancho = tabla.getTableHeader().getFontMetrics(
                    tabla.getTableHeader().getFont()).stringWidth(tabla.getColumnName(columna));
            for (int fila = 0; fila < tabla.getRowCount(); fila++) {
                java.awt.Component componente = tabla.prepareRenderer(
                        tabla.getCellRenderer(fila, columna), fila, columna);
                ancho = Math.max(ancho, componente.getPreferredSize().width);
            }
            int anchoAjustado = Math.max(minimos[columna],
                    Math.min(maximos[columna], ancho + 24));
            modeloColumnas.getColumn(columna).setPreferredWidth(anchoAjustado);
        }
    }//ajustarAnchoColumnasIncidencias

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
        if ("1".equals(estadoNormalizado)) {
            prepararBoton(jButton3, "DESACTIVAR (ACTIVO)", iconoPower(COLOR_VERDE),
                    java.awt.Color.WHITE, COLOR_VERDE, COLOR_VERDE);
            jButton3.setToolTipText("Estado actual: ACTIVO (1)");
            return;
        }

        prepararBoton(jButton3, "ACTIVAR (INACTIVO)", iconoPower(COLOR_ROJO),
                java.awt.Color.WHITE, COLOR_ROJO, COLOR_ROJO);
        jButton3.setToolTipText("Estado actual: INACTIVO (0)");
    }//actualizarIndicadorEstado

    private void restablecerIndicadorEstado() {
        prepararBoton(jButton3, "ACTIVAR/DESACTIVAR", iconoPower(COLOR_GRIS),
                java.awt.Color.WHITE, COLOR_GRIS, COLOR_GRIS);
        jButton3.setToolTipText("Estado actual: sin seleccionar");
    }//restablecerIndicadorEstado

    private void mostrarError(String mensajeUsuario, Exception ex) {
        logger.log(java.util.logging.Level.SEVERE, mensajeUsuario, ex);
        JOptionPane.showMessageDialog(this, mensajeUsuario + "\nDetalle tecnico: " + ex.getMessage());
    }//mostrarError
}
