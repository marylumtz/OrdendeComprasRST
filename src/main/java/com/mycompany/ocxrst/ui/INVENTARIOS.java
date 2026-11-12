/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.util.IconoVentanaUtil;
import com.mycompany.ocxrst.util.EstiloAplicacion;
import com.mycompany.ocxrst.config.AppConfig;
import com.mycompany.ocxrst.repository.InventarioRepository;
import com.mycompany.ocxrst.repository.ProductoProveedorRepository;
import com.mycompany.ocxrst.repository.ProveedorRepository;
import com.mycompany.ocxrst.service.InventarioService;
import com.mycompany.ocxrst.validation.InventarioValidator;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Maria Luisa Martinez
 */
public class INVENTARIOS extends javax.swing.JFrame {

    // =====================================================================
    // Estado, repositorios y configuracion del inventario
    // =====================================================================

    private static final java.util.logging.Logger logger = java.util.logging.Logger.getLogger(INVENTARIOS.class.getName());
    private final InventarioRepository inventarioRepository = new InventarioRepository();
    private final ProveedorRepository proveedorRepository = new ProveedorRepository();
    private final ProductoProveedorRepository productoProveedorRepository = new ProductoProveedorRepository();

    // Caché de productos existentes (código | descripción) para el autocompletado del campo CÓDIGO.
    



    // ===================== Paleta de colores (misma que PROVEEDORES) =====================
    private static final java.awt.Color COLOR_AZUL_OSCURO = new java.awt.Color(28, 82, 132);
    private static final java.awt.Color COLOR_AZUL = new java.awt.Color(31, 112, 184);
    private static final java.awt.Color COLOR_FONDO = EstiloAplicacion.FONDO_PANTALLA;
    private static final java.awt.Color COLOR_BORDE_CAMPO = new java.awt.Color(224, 232, 242);
    private static final java.awt.Color COLOR_BORDE_TABLA = new java.awt.Color(225, 232, 241);
    private static final java.awt.Color COLOR_TEXTO = new java.awt.Color(24, 58, 104);
    private static final java.awt.Color COLOR_TEXTO_SECUNDARIO = new java.awt.Color(93, 119, 153);
    private static final java.awt.Color COLOR_VERDE = new java.awt.Color(0, 130, 60);
    private static final java.awt.Color COLOR_ROJO = new java.awt.Color(190, 0, 0);
    private static final java.awt.Color COLOR_AZUL_CLARO = new java.awt.Color(225, 239, 255);
    private static final java.awt.Color COLOR_AZUL_PALIDO = new java.awt.Color(240, 246, 253);
    private static final java.awt.Color COLOR_AMARILLO_PALIDO = new java.awt.Color(255, 242, 211);
    private static final java.awt.Color COLOR_VERDE_PALIDO = new java.awt.Color(224, 246, 231);
    private static final String FAMILIA_TIPOGRAFICA = "Segoe UI";
    private static final java.awt.Font FUENTE_TITULO = new java.awt.Font(FAMILIA_TIPOGRAFICA,
            java.awt.Font.BOLD, 20);
    private static final java.awt.Font FUENTE_SUBTITULO = new java.awt.Font(FAMILIA_TIPOGRAFICA,
            java.awt.Font.PLAIN, 11);
    private static final java.awt.Font FUENTE_BOTON = new java.awt.Font(FAMILIA_TIPOGRAFICA,
            java.awt.Font.BOLD, 11);
    private static final java.awt.Font FUENTE_ETIQUETA = new java.awt.Font(FAMILIA_TIPOGRAFICA,
            java.awt.Font.BOLD, 10);

    // ===================== Campos de la interfaz =====================
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel panelDashboard;
    private javax.swing.table.DefaultTableModel modeloInventarioDashboard;
    private javax.swing.JTable tablaInventarioDashboard;
    private javax.swing.JTextField filtroInventario;
    private javax.swing.JComboBox<String> filtroCategoriaInventario;
    private javax.swing.JComboBox<String> filtroMarcaInventario;
    private javax.swing.JComboBox<String> filtroEstadoInventario;
    private javax.swing.JLabel totalProductosDashboard;
    private javax.swing.JLabel stockBajoDashboard;
    private javax.swing.JLabel sinExistenciaDashboard;
    private javax.swing.JLabel valorInventarioDashboard;
    private boolean actualizandoFiltrosInventario;

    // =====================================================================
    // Construccion y componentes del formulario
    // =====================================================================

    /** Inicializa la pantalla, archivos, eventos y autocompletado. */
    public INVENTARIOS() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        initComponents();
        configurarDisenoVisual();
        IconoVentanaUtil.aplicar(this);
        setSize(1280, 820);
        setResizable(false);
        setLocationRelativeTo(null);
        asegurarArchivoInventarios();
        EstiloAplicacion.aplicarEstiloBotones(jPanel1);
        EstiloAplicacion.aplicarTipografia(jPanel1);
    }

    private void initComponents() {
        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("INVENTARIOS");

        jPanel1 = new javax.swing.JPanel();

        getContentPane().setLayout(new java.awt.BorderLayout());
        getContentPane().add(jPanel1, java.awt.BorderLayout.CENTER);


        pack();
    }

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

        java.awt.EventQueue.invokeLater(() -> new INVENTARIOS().setVisible(true));
    }

    // =====================================================================
    // Diseño visual y eventos del inventario
    // =====================================================================

    // ===================== Diseño visual =====================
    private void configurarDisenoVisual() {
        jPanel1.setLayout(new java.awt.BorderLayout());
        jPanel1.setBackground(COLOR_FONDO);
        jPanel1.setBorder(javax.swing.BorderFactory.createEmptyBorder(18, 24, 18, 24));

        panelDashboard = crearDashboardInventario();
        EstiloAplicacion.aplicarEstiloBotones(panelDashboard);
        EstiloAplicacion.aplicarTipografia(panelDashboard);
        ajustarBotonesDashboard();
        jPanel1.add(panelDashboard, java.awt.BorderLayout.CENTER);
    }//configurarDisenoVisual

    private void ajustarBotonesDashboard() {
        javax.swing.JButton nuevo = buscarBotonDashboard("NUEVO PRODUCTO");
        javax.swing.JButton exportar = buscarBotonDashboard("EXPORTAR");
        javax.swing.JButton actualizar = buscarBotonDashboard("ACTUALIZAR");
        javax.swing.JButton limpiar = buscarBotonDashboard("LIMPIAR FILTROS");
        dimensionarBotonDashboard(nuevo, 175, 40);
        dimensionarBotonDashboard(exportar, 142, 40);
        dimensionarBotonDashboard(actualizar, 142, 40);
        dimensionarBotonDashboard(limpiar, 150, 34);
        dimensionarBotonDashboard(buscarBotonDashboard("ENTRADA DE INVENTARIO"), 240, 40);
        dimensionarBotonDashboard(buscarBotonDashboard("SALIDA DE INVENTARIO"), 240, 40);
    }

    private javax.swing.JButton buscarBotonDashboard(String texto) {
        return buscarBotonDashboard(panelDashboard, texto);
    }

    private javax.swing.JButton buscarBotonDashboard(java.awt.Container contenedor, String texto) {
        for (java.awt.Component componente : contenedor.getComponents()) {
            if (componente instanceof javax.swing.JButton boton
                    && boton.getText().replace(" ", "").equals(texto.replace(" ", ""))) {
                return boton;
            }
            if (componente instanceof java.awt.Container hijo) {
                javax.swing.JButton encontrado = buscarBotonDashboard(hijo, texto);
                if (encontrado != null) {
                    return encontrado;
                }
            }
        }
        return null;
    }

    private void dimensionarBotonDashboard(javax.swing.JButton boton, int ancho, int alto) {
        if (boton == null) {
            return;
        }
        java.awt.Dimension dimension = new java.awt.Dimension(ancho, alto);
        boton.setPreferredSize(dimension);
        boton.setMinimumSize(dimension);
        boton.setMaximumSize(dimension);
        boton.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        boton.setHorizontalTextPosition(javax.swing.SwingConstants.RIGHT);
        boton.setIconTextGap(7);
    }

    private javax.swing.JPanel crearDashboardInventario() {
        javax.swing.JPanel dashboard = new javax.swing.JPanel(new java.awt.BorderLayout(0, 12));
        dashboard.setOpaque(false);
        javax.swing.JPanel encabezado = new javax.swing.JPanel(new java.awt.BorderLayout(12, 0));
        encabezado.setOpaque(false);
        javax.swing.JLabel titulo = new javax.swing.JLabel("<html><b>Inventario de Productos</b><br>"
                + "<font size='3'>Administración, consulta y seguimiento de existencias.</font></html>");
        titulo.setForeground(COLOR_AZUL_OSCURO);
        titulo.setFont(FUENTE_TITULO);
        encabezado.add(titulo, java.awt.BorderLayout.CENTER);
        javax.swing.JPanel acciones = new javax.swing.JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 8, 0));
        acciones.setOpaque(false);
        javax.swing.JButton nuevo = new javax.swing.JButton("  NUEVO PRODUCTO", iconoMas(java.awt.Color.WHITE));
        nuevo.setBackground(COLOR_AZUL);
        nuevo.setForeground(java.awt.Color.WHITE);
        nuevo.setFont(FUENTE_BOTON);
        nuevo.addActionListener(e -> abrirAltaProducto());
        javax.swing.JButton exportar = new javax.swing.JButton("  EXPORTAR", iconoGuardar(COLOR_AZUL_OSCURO));
        exportar.setForeground(COLOR_AZUL_OSCURO);
        exportar.setFont(FUENTE_BOTON);
        exportar.setBackground(java.awt.Color.WHITE);
        exportar.addActionListener(e -> exportarInventarioDashboard());
        javax.swing.JButton actualizar = new javax.swing.JButton("  ACTUALIZAR", iconoLimpiar(COLOR_AZUL_OSCURO));
        actualizar.setForeground(COLOR_AZUL_OSCURO);
        actualizar.setFont(FUENTE_BOTON);
        actualizar.setBackground(java.awt.Color.WHITE);
        actualizar.addActionListener(e -> cargarTablaInventarioDashboard());
        acciones.add(nuevo);
        acciones.add(exportar);
        acciones.add(actualizar);
        encabezado.add(acciones, java.awt.BorderLayout.EAST);
        dashboard.add(encabezado, java.awt.BorderLayout.NORTH);

        javax.swing.JPanel tarjetas = new javax.swing.JPanel(new java.awt.GridLayout(1, 4, 10, 0));
        tarjetas.setOpaque(false);
        totalProductosDashboard = new javax.swing.JLabel("0");
        stockBajoDashboard = new javax.swing.JLabel("0");
        sinExistenciaDashboard = new javax.swing.JLabel("0");
        valorInventarioDashboard = new javax.swing.JLabel("$ 0.00");
        tarjetas.add(crearTarjetaResumenInventario("Total de productos", totalProductosDashboard,
                COLOR_AZUL, COLOR_AZUL_CLARO));
        tarjetas.add(crearTarjetaResumenInventario("Stock bajo", stockBajoDashboard,
                new java.awt.Color(190, 125, 0), COLOR_AMARILLO_PALIDO));
        tarjetas.add(crearTarjetaResumenInventario("Sin existencias", sinExistenciaDashboard,
                new java.awt.Color(190, 45, 55), new java.awt.Color(255, 232, 234)));
        tarjetas.add(crearTarjetaResumenInventario("Valor del inventario", valorInventarioDashboard,
                new java.awt.Color(0, 125, 70), COLOR_VERDE_PALIDO));

        javax.swing.JPanel accionesMovimiento = new javax.swing.JPanel(
                new java.awt.FlowLayout(java.awt.FlowLayout.LEFT, 8, 0));
        accionesMovimiento.setOpaque(false);
        javax.swing.JButton entrada = new javax.swing.JButton("  ENTRADA DE INVENTARIO",
                iconoCamion(COLOR_AZUL));
        entrada.setForeground(COLOR_AZUL_OSCURO);
        entrada.setFont(FUENTE_BOTON);
        entrada.setBackground(java.awt.Color.WHITE);
        entrada.setToolTipText("Consultar entradas de inventario");
        entrada.addActionListener(e -> mostrarFormularioEntradaInventario());
        javax.swing.JButton salida = new javax.swing.JButton("  SALIDA DE INVENTARIO",
                iconoBasura(COLOR_AZUL));
        salida.setForeground(COLOR_AZUL_OSCURO);
        salida.setFont(FUENTE_BOTON);
        salida.setBackground(java.awt.Color.WHITE);
        salida.setToolTipText("Registrar salida de inventario");
        salida.addActionListener(e -> mostrarFormularioSalidaInventario());
        accionesMovimiento.add(entrada);
        accionesMovimiento.add(salida);

        javax.swing.JPanel filtros = new javax.swing.JPanel(new java.awt.GridBagLayout());
        filtros.setOpaque(false);
        filtroInventario = new javax.swing.JTextField();
        filtroInventario.setBackground(java.awt.Color.WHITE);
        filtroInventario.setForeground(COLOR_TEXTO);
        filtroInventario.setFont(FUENTE_SUBTITULO);
        filtroInventario.setToolTipText("Buscar por código, descripción o marca");
        filtroCategoriaInventario = new javax.swing.JComboBox<>(new String[]{"Todas"});
        filtroMarcaInventario = new javax.swing.JComboBox<>(new String[]{"Todas"});
        filtroEstadoInventario = new javax.swing.JComboBox<>(new String[]{"Todos", "Disponible", "Stock bajo", "Agotado"});
        EstiloAplicacion.aplicarEstiloCamposTexto(filtroInventario);
        EstiloAplicacion.aplicarEstiloCamposTexto(filtroCategoriaInventario);
        EstiloAplicacion.aplicarEstiloCamposTexto(filtroMarcaInventario);
        EstiloAplicacion.aplicarEstiloCamposTexto(filtroEstadoInventario);
        agregarFiltroInventario(filtros, "Buscar por código, descripción o marca", filtroInventario, 0, 1.8);
        agregarFiltroInventario(filtros, "Categoría", filtroCategoriaInventario, 1, 0.8);
        agregarFiltroInventario(filtros, "Marca", filtroMarcaInventario, 2, 0.8);
        agregarFiltroInventario(filtros, "Estado de stock", filtroEstadoInventario, 3, 0.9);
        javax.swing.JButton limpiar = new javax.swing.JButton("  LIMPIAR FILTROS", iconoLimpiar(COLOR_AZUL));
        limpiar.addActionListener(e -> {
            filtroInventario.setText("");
            filtroCategoriaInventario.setSelectedIndex(0);
            filtroMarcaInventario.setSelectedIndex(0);
            filtroEstadoInventario.setSelectedIndex(0);
        });
        limpiar.setFont(FUENTE_BOTON);
        limpiar.setBackground(java.awt.Color.WHITE);
        limpiar.setPreferredSize(new java.awt.Dimension(0, 30));
        java.awt.GridBagConstraints limpiarGbc = new java.awt.GridBagConstraints();
        limpiarGbc.gridx = 4; limpiarGbc.gridy = 0; limpiarGbc.weightx = 0.7;
        limpiarGbc.fill = java.awt.GridBagConstraints.HORIZONTAL;
        limpiarGbc.insets = new java.awt.Insets(8, 5, 2, 0);
        filtros.add(limpiar, limpiarGbc);

        javax.swing.JPanel superior = new javax.swing.JPanel();
        superior.setLayout(new javax.swing.BoxLayout(superior, javax.swing.BoxLayout.Y_AXIS));
        superior.setOpaque(false);
        superior.add(encabezado);
        superior.add(javax.swing.Box.createVerticalStrut(10));
        superior.add(tarjetas);
        superior.add(javax.swing.Box.createVerticalStrut(8));
        superior.add(accionesMovimiento);
        superior.add(javax.swing.Box.createVerticalStrut(8));
        superior.add(filtros);
        dashboard.remove(encabezado);
        dashboard.add(superior, java.awt.BorderLayout.NORTH);

        modeloInventarioDashboard = new javax.swing.table.DefaultTableModel(
                new String[]{"CÓDIGO", "DESCRIPCIÓN", "CATEGORÍA", "MARCA",
                    "UNIDAD", "EXISTENCIA", "STOCK MÍN.", "PRECIO UNIT.", "ESTADO", "ACCIONES"}, 0) {
            @Override public boolean isCellEditable(int fila, int columna) { return false; }
        };
            tablaInventarioDashboard = new javax.swing.JTable(modeloInventarioDashboard) {
            @Override
            public String getToolTipText(java.awt.event.MouseEvent evento) {
                int fila = rowAtPoint(evento.getPoint());
                int columna = columnAtPoint(evento.getPoint());
                if (fila >= 0 && columna == 1) {
                    int filaModelo = convertRowIndexToModel(fila);
                    return String.valueOf(getModel().getValueAt(filaModelo, columna));
                }
                if (fila >= 0 && columna == 9) {
                    return "Visualizar, editar o desactivar producto";
                }
                return super.getToolTipText(evento);
            }
        };
            EstiloAplicacion.aplicarEstiloTabla(tablaInventarioDashboard);
        javax.swing.table.TableRowSorter<javax.swing.table.TableModel> ordenInventario =
                new javax.swing.table.TableRowSorter<>(modeloInventarioDashboard);
        java.util.Comparator<Object> comparadorNumerico = (primero, segundo) ->
                Double.compare(numero(String.valueOf(primero)), numero(String.valueOf(segundo)));
        ordenInventario.setComparator(5, comparadorNumerico);
        ordenInventario.setComparator(6, comparadorNumerico);
        ordenInventario.setComparator(7, comparadorNumerico);
        tablaInventarioDashboard.setRowSorter(ordenInventario);
        tablaInventarioDashboard.getColumnModel().getColumn(9).setMinWidth(132);
        tablaInventarioDashboard.getColumnModel().getColumn(9).setPreferredWidth(144);
        tablaInventarioDashboard.getColumnModel().getColumn(8).setCellRenderer(
                (tabla, valor, seleccionado, foco, fila, columna) -> {
                    String estado = String.valueOf(valor);
                    java.awt.Color textoEstado = "Disponible".equals(estado)
                            ? COLOR_VERDE : "Stock bajo".equals(estado)
                            ? new java.awt.Color(184, 116, 0) : COLOR_ROJO;
                    java.awt.Color fondoEstado = "Disponible".equals(estado)
                            ? COLOR_VERDE_PALIDO : "Stock bajo".equals(estado)
                            ? COLOR_AMARILLO_PALIDO : new java.awt.Color(255, 232, 234);
                    javax.swing.JPanel celdaEstado = new javax.swing.JPanel(
                            new java.awt.FlowLayout(java.awt.FlowLayout.CENTER, 0, 5));
                    celdaEstado.setOpaque(true);
                    celdaEstado.setBackground(seleccionado ? COLOR_AZUL_CLARO
                            : (fila % 2 == 0 ? java.awt.Color.WHITE : new java.awt.Color(242, 242, 242)));
                    javax.swing.JLabel etiquetaEstado = new javax.swing.JLabel(
                            ("Disponible".equals(estado) ? "\u2713 " : "\u26a0 ") + estado);
                    etiquetaEstado.setOpaque(true);
                    etiquetaEstado.setFont(FUENTE_ETIQUETA);
                    etiquetaEstado.setForeground(textoEstado);
                    etiquetaEstado.setBackground(fondoEstado);
                    etiquetaEstado.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                            new BordeRedondeado(10, fondoEstado, 1f),
                            javax.swing.BorderFactory.createEmptyBorder(3, 8, 3, 8)));
                    celdaEstado.add(etiquetaEstado);
                    return celdaEstado;
                });
        tablaInventarioDashboard.getColumnModel().getColumn(9).setCellRenderer(
                (tabla, valor, seleccionado, foco, fila, columna) -> {
                    javax.swing.JPanel botonesAccion = new javax.swing.JPanel(
                            new java.awt.GridLayout(1, 3, 8, 0));
                    botonesAccion.setOpaque(true);
                    botonesAccion.setBackground(seleccionado ? COLOR_AZUL_CLARO : java.awt.Color.WHITE);
                    botonesAccion.setBorder(javax.swing.BorderFactory.createEmptyBorder(5, 6, 5, 6));
                    javax.swing.JLabel ver = new javax.swing.JLabel(iconoOjo(COLOR_AZUL));
                    ver.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
                    ver.setToolTipText("Visualizar producto");
                    javax.swing.JLabel editar = new javax.swing.JLabel(iconoEditar(COLOR_AZUL));
                    editar.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
                    editar.setToolTipText("Editar producto");
                    javax.swing.JLabel eliminar = new javax.swing.JLabel(iconoBasura(COLOR_ROJO));
                    eliminar.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
                    eliminar.setToolTipText("Desactivar producto");
                    botonesAccion.add(ver);
                    botonesAccion.add(editar);
                    botonesAccion.add(eliminar);
                    return botonesAccion;
                });
        tablaInventarioDashboard.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override public void mouseClicked(java.awt.event.MouseEvent e) {
                int fila = tablaInventarioDashboard.rowAtPoint(e.getPoint());
                if (fila >= 0 && tablaInventarioDashboard.columnAtPoint(e.getPoint()) == 9) {
                    int filaModelo = tablaInventarioDashboard.convertRowIndexToModel(fila);
                    String codigo = String.valueOf(modeloInventarioDashboard.getValueAt(filaModelo, 0));
                    java.awt.Rectangle celdaAcciones = tablaInventarioDashboard.getCellRect(fila, 9, true);
                    int posicion = e.getX() - celdaAcciones.x;
                    int anchoBoton = Math.max(1, celdaAcciones.width / 3);
                    if (posicion < anchoBoton) {
                        verProductoDashboard(codigo);
                    } else if (posicion < anchoBoton * 2) {
                        mostrarFormularioProducto(codigo);
                    } else {
                        desactivarProducto(codigo);
                        mostrarDashboardInventario();
                    }
                }
            }
        });
        javax.swing.JScrollPane scrollTabla = new javax.swing.JScrollPane(tablaInventarioDashboard);
        scrollTabla.setBorder(javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_TABLA));
        dashboard.add(scrollTabla, java.awt.BorderLayout.CENTER);
        filtroInventario.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
            public void insertUpdate(javax.swing.event.DocumentEvent e) { cargarTablaInventarioDashboard(); }
            public void removeUpdate(javax.swing.event.DocumentEvent e) { cargarTablaInventarioDashboard(); }
            public void changedUpdate(javax.swing.event.DocumentEvent e) { cargarTablaInventarioDashboard(); }
        });
        filtroCategoriaInventario.addActionListener(e -> cargarTablaInventarioDashboard());
        filtroMarcaInventario.addActionListener(e -> cargarTablaInventarioDashboard());
        filtroEstadoInventario.addActionListener(e -> cargarTablaInventarioDashboard());
        cargarTablaInventarioDashboard();
        return dashboard;
    }

    private javax.swing.JPanel crearTarjetaResumenInventario(String titulo,
            javax.swing.JLabel valor, java.awt.Color color, java.awt.Color fondoIcono) {
        javax.swing.JPanel tarjeta = new javax.swing.JPanel(new java.awt.BorderLayout(0, 3));
        tarjeta.setBackground(java.awt.Color.WHITE);
        tarjeta.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(COLOR_BORDE_CAMPO),
                javax.swing.BorderFactory.createEmptyBorder(10, 14, 10, 14)));
        javax.swing.JLabel etiqueta = new javax.swing.JLabel(titulo);
        etiqueta.setForeground(COLOR_TEXTO_SECUNDARIO);
        etiqueta.setFont(FUENTE_SUBTITULO);
        valor.setForeground(color);
        valor.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 20));
        javax.swing.JPanel contenido = new javax.swing.JPanel(new java.awt.BorderLayout(10, 0));
        contenido.setOpaque(false);
        javax.swing.JLabel icono = new javax.swing.JLabel();
        icono.setIcon("Total de productos".equals(titulo) ? iconoCaja(color)
                : "Stock bajo".equals(titulo) ? iconoAdvertencia(color)
                : "Sin existencias".equals(titulo) ? iconoBasura(color) : iconoMonedas(color));
        icono.setOpaque(true);
        icono.setBackground(fondoIcono);
        icono.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        icono.setPreferredSize(new java.awt.Dimension(48, 48));
        contenido.add(icono, java.awt.BorderLayout.WEST);
        javax.swing.JPanel textos = new javax.swing.JPanel(new java.awt.BorderLayout(0, 2));
        textos.setOpaque(false);
        textos.add(etiqueta, java.awt.BorderLayout.NORTH);
        textos.add(valor, java.awt.BorderLayout.CENTER);
        contenido.add(textos, java.awt.BorderLayout.CENTER);
        tarjeta.add(contenido, java.awt.BorderLayout.CENTER);
        return tarjeta;
    }

    private void agregarFiltroInventario(javax.swing.JPanel panel, String texto,
            java.awt.Component campo, int columna, double peso) {
        javax.swing.JPanel envoltura = new javax.swing.JPanel(new java.awt.BorderLayout(0, 3));
        envoltura.setOpaque(false);
        javax.swing.JLabel etiqueta = new javax.swing.JLabel(texto);
        etiqueta.setForeground(COLOR_AZUL_OSCURO);
        etiqueta.setFont(new java.awt.Font(FAMILIA_TIPOGRAFICA, java.awt.Font.BOLD, 11));
        envoltura.add(etiqueta, java.awt.BorderLayout.NORTH);
        envoltura.add(campo, java.awt.BorderLayout.CENTER);
        java.awt.GridBagConstraints gbc = new java.awt.GridBagConstraints();
        gbc.gridx = columna; gbc.gridy = 0; gbc.weightx = peso;
        gbc.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gbc.insets = new java.awt.Insets(2, columna == 0 ? 0 : 5, 2, 0);
        panel.add(envoltura, gbc);
    }

    private void verProductoDashboard(String codigo) {
        try {
            File archivo = obtenerArchivoInventarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Row fila = obtenerHojaInventarios(libro).getRow(buscarFilaPorCodigo(
                        obtenerHojaInventarios(libro), codigo));
                if (fila != null) {
                    mostrarMovimientosProducto(codigo, leerCelda(fila, 1));
                }
            }
        } catch (IOException ex) {
            mostrarError("No se pudo consultar el producto.", ex);
        }
    }

    private void mostrarMovimientosProducto(String codigo, String descripcion) {
        javax.swing.table.DefaultTableModel modelo = new javax.swing.table.DefaultTableModel(
                new String[]{"Fecha", "Tipo", "Documento", "Proveedor / Cliente",
                    "Cantidad", "Usuario", "Observaciones"}, 0) {
            @Override public boolean isCellEditable(int fila, int columna) {
                return false;
            }
        };
        try {
            agregarMovimientosRecepcion(modelo, codigo);
            agregarMovimientosEntradaManual(modelo, codigo);
            agregarMovimientosArchivo(modelo, AppConfig.resolveBaseFile("REGISTRODEVOLUCIONES.xlsx"),
                    codigo, "Salida", new String[]{"FECHA", "ORDENCOMPRA", "CANTIDAD",
                        "REGISTRADO_POR", "MOTIVO"});
            agregarMovimientosArchivo(modelo, AppConfig.resolveBaseFile("REGISTROINCIDENCIAS.xlsx"),
                    codigo, "Ajuste", new String[]{"FECHA", "ORDENCOMPRA", "CANTIDAD",
                        "REGISTRADO_POR", "DESCRIPCION"});

            if (modelo.getRowCount() == 0) {
                JOptionPane.showMessageDialog(this,
                        "No hay movimientos registrados para " + codigo + ".",
                        "Historial del producto", JOptionPane.PLAIN_MESSAGE);
                return;
            }
            javax.swing.JTable tabla = new javax.swing.JTable(modelo);
            EstiloAplicacion.aplicarEstiloTabla(tabla);
            tabla.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_LAST_COLUMN);
            javax.swing.JScrollPane scroll = new javax.swing.JScrollPane(tabla);
            scroll.setPreferredSize(new java.awt.Dimension(850, 260));
            JOptionPane.showMessageDialog(this, scroll,
                    "Movimientos de " + codigo + " - " + descripcion,
                    JOptionPane.PLAIN_MESSAGE);
        } catch (IOException ex) {
            mostrarError("No se pudo cargar el historial del producto.", ex);
        }
    }

    private void mostrarMovimientosInventario(String tipoMovimiento) {
        javax.swing.table.DefaultTableModel modelo = new javax.swing.table.DefaultTableModel(
                new String[]{"Fecha", "Tipo", "Documento", "Proveedor / Cliente",
                    "Cantidad", "Usuario", "Observaciones"}, 0) {
            @Override public boolean isCellEditable(int fila, int columna) {
                return false;
            }
        };
        try {
            agregarMovimientosRecepcion(modelo, null);
            agregarMovimientosEntradaManual(modelo, null);
            agregarMovimientosArchivo(modelo, AppConfig.resolveBaseFile("REGISTRODEVOLUCIONES.xlsx"),
                    null, "Salida", new String[]{"FECHA", "ORDENCOMPRA", "CANTIDAD",
                        "REGISTRADO_POR", "MOTIVO"});
            agregarMovimientosArchivo(modelo, AppConfig.resolveBaseFile("REGISTROINCIDENCIAS.xlsx"),
                    null, "Ajuste", new String[]{"FECHA", "ORDENCOMPRA", "CANTIDAD",
                        "REGISTRADO_POR", "DESCRIPCION"});
            for (int indice = modelo.getRowCount() - 1; indice >= 0; indice--) {
                if (!tipoMovimiento.equals(modelo.getValueAt(indice, 1))) {
                    modelo.removeRow(indice);
                }
            }
            if (modelo.getRowCount() == 0) {
                JOptionPane.showMessageDialog(this,
                        "No hay movimientos de tipo " + tipoMovimiento + " registrados.",
                        "Historial del inventario", JOptionPane.PLAIN_MESSAGE);
                return;
            }
            javax.swing.JTable tabla = new javax.swing.JTable(modelo);
            EstiloAplicacion.aplicarEstiloTabla(tabla);
            javax.swing.JScrollPane scroll = new javax.swing.JScrollPane(tabla);
            scroll.setPreferredSize(new java.awt.Dimension(850, 260));
            JOptionPane.showMessageDialog(this, scroll,
                    tipoMovimiento + " de inventario", JOptionPane.PLAIN_MESSAGE);
        } catch (IOException ex) {
            mostrarError("No se pudo cargar el historial del inventario.", ex);
        }
    }

    private void mostrarFormularioEntradaInventario() {
        ENTRADAINVENTARIO formulario = new ENTRADAINVENTARIO(this::mostrarDashboardInventario);
        jPanel1.removeAll();
        jPanel1.add(formulario, java.awt.BorderLayout.CENTER);
        jPanel1.revalidate();
        jPanel1.repaint();
    }

    private void mostrarFormularioSalidaInventario() {
        SALIDAINVENTARIO formulario = new SALIDAINVENTARIO(this::mostrarDashboardInventario);
        jPanel1.removeAll();
        jPanel1.add(formulario, java.awt.BorderLayout.CENTER);
        jPanel1.revalidate();
        jPanel1.repaint();
    }

    private void agregarMovimientosRecepcion(javax.swing.table.DefaultTableModel modelo,
            String codigo) throws IOException {
        File archivo = AppConfig.resolveBaseFile("REGISTRORECIBIR.xlsx");
        if (!archivo.exists()) {
            return;
        }
        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaInventarios(libro);
            for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                Row fila = hoja.getRow(indice);
                if (codigo == null || codigo.equalsIgnoreCase(leerCelda(fila, 6))) {
                    modelo.addRow(new Object[]{leerCelda(fila, 2), "Entrada",
                        leerCelda(fila, 0), valorONa(leerCelda(fila, 1)), leerCelda(fila, 7),
                        valorONa(leerCelda(fila, 5), "Usuario no identificado"),
                        "Recepción de mercancía"});
                }
            }
        }
    }

    private void agregarMovimientosEntradaManual(javax.swing.table.DefaultTableModel modelo,
            String codigo) throws IOException {
        agregarMovimientosManuales(modelo, codigo, AppConfig.resolveBaseFile("MOVIMIENTOSINVENTARIO.xlsx"),
                inventarioRepository);
    }

    static void agregarMovimientosManuales(javax.swing.table.DefaultTableModel modelo,
            String codigo, File archivo, InventarioRepository repository) throws IOException {
        agregarMovimientosManuales(modelo, codigo, archivo, repository, AppConfig.actividadesFile());
    }

    static void agregarMovimientosManuales(javax.swing.table.DefaultTableModel modelo,
            String codigo, File archivo, InventarioRepository repository, File archivoActividades)
            throws IOException {
        if (!archivo.exists()) {
            return;
        }
        Map<String, String> usuariosPorFolio = cargarUsuariosMovimientos(archivoActividades, repository);
        try (Workbook libro = repository.cargarLibro(archivo)) {
            Sheet hoja = repository.obtenerHojaInventarios(libro);
            for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                Row fila = hoja.getRow(indice);
                if (fila == null || (codigo != null
                        && !codigo.equalsIgnoreCase(repository.leerCelda(fila, 2)))) {
                    continue;
                }
                String folio = repository.leerCelda(fila, 0);
                boolean salida = folio.startsWith("SI");
                String usuario = repository.leerCelda(fila, 6);
                if (usuario.isBlank()) {
                    usuario = usuariosPorFolio.getOrDefault(folio, "Usuario no identificado");
                }
                modelo.addRow(new Object[]{repository.leerCelda(fila, 1), salida ? "Salida" : "Entrada",
                    folio, "N/A", repository.leerCelda(fila, 4), usuario,
                    salida ? "Salida manual de inventario" : "Entrada manual de inventario"});
            }
        }
    }

    private static Map<String, String> cargarUsuariosMovimientos(File archivoActividades,
            InventarioRepository repository) throws IOException {
        Map<String, String> usuariosPorFolio = new HashMap<>();
        if (!archivoActividades.exists()) {
            return usuariosPorFolio;
        }
        try (Workbook libro = repository.cargarLibro(archivoActividades)) {
            Sheet hoja = repository.obtenerHojaInventarios(libro);
            for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                Row fila = hoja.getRow(indice);
                if (!"Inventarios".equalsIgnoreCase(repository.leerCelda(fila, 1))) {
                    continue;
                }
                String folio = repository.leerCelda(fila, 3);
                String usuario = repository.leerCelda(fila, 4);
                if (!folio.isBlank() && !usuario.isBlank()) {
                    usuariosPorFolio.put(folio, usuario);
                }
            }
        }
        return usuariosPorFolio;
    }

    private void agregarMovimientosArchivo(javax.swing.table.DefaultTableModel modelo,
            File archivo, String codigo, String tipo, String[] columnas) throws IOException {
        if (!archivo.exists()) {
            return;
        }
        try (Workbook libro = cargarLibro(archivo)) {
            Sheet hoja = obtenerHojaInventarios(libro);
            for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                Row fila = hoja.getRow(indice);
                if (codigo == null || codigo.equalsIgnoreCase(leerCelda(fila, 2))) {
                    modelo.addRow(new Object[]{leerCelda(fila, indiceColumna(hoja, columnas[0])),
                        tipo, leerCelda(fila, indiceColumna(hoja, columnas[1])), "N/A",
                        leerCelda(fila, indiceColumna(hoja, columnas[2])),
                        valorONa(leerCelda(fila, indiceColumna(hoja, columnas[3])),
                                "Usuario no identificado"),
                        leerCelda(fila, indiceColumna(hoja, columnas[4]))});
                }
            }
        }
    }

    private static String valorONa(String valor) {
        return valorONa(valor, "N/A");
    }

    private static String valorONa(String valor, String valorAlternativo) {
        return valor == null || valor.isBlank() ? valorAlternativo : valor;
    }

    private int indiceColumna(Sheet hoja, String encabezado) {
        Row fila = hoja.getRow(0);
        for (int indice = 0; indice < fila.getLastCellNum(); indice++) {
            if (encabezado.equalsIgnoreCase(leerCelda(fila, indice))) {
                return indice;
            }
        }
        return -1;
    }

    private void abrirAltaProducto() {
        mostrarFormularioProducto(null);
    }

    private void mostrarFormularioProducto() {
        mostrarFormularioProducto(null);
    }

    private void mostrarFormularioProducto(String codigoExistente) {
        ALTAPRODUCTO formulario = new ALTAPRODUCTO(this::mostrarDashboardInventario,
                codigoExistente);
        jPanel1.removeAll();
        jPanel1.add(formulario, java.awt.BorderLayout.CENTER);
        jPanel1.revalidate();
        jPanel1.repaint();
    }

    private void mostrarDashboardInventario() {
        cargarTablaInventarioDashboard();
        jPanel1.removeAll();
        jPanel1.add(panelDashboard, java.awt.BorderLayout.CENTER);
        jPanel1.revalidate();
        jPanel1.repaint();
    }

    private void cargarTablaInventarioDashboard() {
        if (modeloInventarioDashboard == null || actualizandoFiltrosInventario) {
            return;
        }
        java.util.List<String[]> productos = new java.util.ArrayList<>();
        java.util.Set<String> categorias = new java.util.TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        java.util.Set<String> marcas = new java.util.TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        int[] totales = {0, 0, 0};
        double[] valor = {0};
        try {
            File archivo = obtenerArchivoInventarios();
            if (archivo.exists()) {
                try (Workbook libro = cargarLibro(archivo)) {
                    Sheet hoja = obtenerHojaInventarios(libro);
                    for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                        Row fila = hoja.getRow(indice);
                        if (fila == null) {
                            continue;
                        }
                        String codigo = leerCelda(fila, 0);
                        if (codigo.isBlank()) {
                            continue;
                        }
                        String descripcion = leerCelda(fila, 1);
                        String categoria = leerCelda(fila, 6);
                        String marca = leerCelda(fila, 7);
                        String unidad = leerCelda(fila, 2);
                        String existencia = leerCelda(fila, 4);
                        String minimo = leerCelda(fila, 8);
                        String precio = leerCelda(fila, 3);
                        boolean activo = "1".equals(normalizarEstado(leerCelda(fila, 5)));
                        if (!activo) {
                            continue;
                        }
                        categorias.add(categoria);
                        marcas.add(marca);
                        double cantidad = numero(existencia);
                        double minimoNumero = numero(minimo);
                        double precioNumero = numero(precio);
                        String estado = cantidad <= 0 ? "Agotado"
                                : cantidad <= minimoNumero ? "Stock bajo" : "Disponible";
                        totales[0]++;
                        if (cantidad <= minimoNumero && cantidad > 0) {
                            totales[1]++;
                        }
                        if (cantidad <= 0) {
                            totales[2]++;
                        }
                        valor[0] += cantidad * precioNumero;
                        productos.add(new String[]{codigo, descripcion, categoria, marca, unidad,
                            existencia, minimo, precio, estado});
                    }
                }
            }
        } catch (IOException ex) {
            logger.log(java.util.logging.Level.WARNING, "No se pudo cargar el tablero de inventario.", ex);
        }
        actualizandoFiltrosInventario = true;
        actualizarOpcionesFiltro(filtroCategoriaInventario, categorias);
        actualizarOpcionesFiltro(filtroMarcaInventario, marcas);
        actualizandoFiltrosInventario = false;
        String texto = filtroInventario.getText().trim().toLowerCase(java.util.Locale.ROOT);
        String categoriaSeleccionada = String.valueOf(filtroCategoriaInventario.getSelectedItem());
        String marcaSeleccionada = String.valueOf(filtroMarcaInventario.getSelectedItem());
        String estadoSeleccionado = String.valueOf(filtroEstadoInventario.getSelectedItem());
        java.util.List<String[]> filtrados = new java.util.ArrayList<>();
        for (String[] producto : productos) {
            boolean coincideTexto = texto.isBlank()
                    || String.join(" ", producto).toLowerCase(java.util.Locale.ROOT).contains(texto);
            boolean coincideCategoria = "Todas".equals(categoriaSeleccionada)
                    || categoriaSeleccionada.equals(producto[2]);
            boolean coincideMarca = "Todas".equals(marcaSeleccionada)
                    || marcaSeleccionada.equals(producto[3]);
            boolean coincideEstado = "Todos".equals(estadoSeleccionado)
                    || estadoSeleccionado.equals(producto[8]);
            if (coincideTexto && coincideCategoria && coincideMarca && coincideEstado) {
                filtrados.add(producto);
            }
        }
        modeloInventarioDashboard.setRowCount(0);
        for (String[] producto : filtrados) {
            modeloInventarioDashboard.addRow(new Object[]{producto[0], producto[1], producto[2],
                producto[3], producto[4], producto[5], producto[6], "$ " + producto[7], producto[8], "ACCIONES"});
        }
        totalProductosDashboard.setText(String.valueOf(totales[0]));
        stockBajoDashboard.setText(String.valueOf(totales[1]));
        sinExistenciaDashboard.setText(String.valueOf(totales[2]));
        valorInventarioDashboard.setText(String.format(java.util.Locale.ROOT, "$ %,.2f", valor[0]));
    }

    private void actualizarOpcionesFiltro(javax.swing.JComboBox<String> combo,
            java.util.Set<String> valores) {
        String seleccionado = String.valueOf(combo.getSelectedItem());
        combo.removeAllItems();
        combo.addItem(combo == filtroEstadoInventario ? "Todos" : "Todas");
        for (String valor : valores) {
            if (!valor.isBlank()) {
                combo.addItem(valor);
            }
        }
        combo.setSelectedItem(seleccionado);
        if (combo.getSelectedIndex() < 0) {
            combo.setSelectedIndex(0);
        }
    }

    private double numero(String valor) {
        try {
            return Double.parseDouble(valor.replace(",", "").replace("$", "").trim());
        } catch (NumberFormatException ex) {
            return 0;
        }
    }

    private void exportarInventarioDashboard() {
        javax.swing.JFileChooser selector = new javax.swing.JFileChooser();
        selector.setSelectedFile(new File("inventario_productos.csv"));
        if (selector.showSaveDialog(this) != javax.swing.JFileChooser.APPROVE_OPTION) {
            return;
        }
        File destino = selector.getSelectedFile();
        try (java.io.BufferedWriter escritor = java.nio.file.Files.newBufferedWriter(
                destino.toPath(), java.nio.charset.StandardCharsets.UTF_8)) {
            for (int columna = 0; columna < modeloInventarioDashboard.getColumnCount() - 1; columna++) {
                if (columna > 0) {
                    escritor.write(",");
                }
                escritor.write("\"" + modeloInventarioDashboard.getColumnName(columna) + "\"");
            }
            escritor.newLine();
            for (int fila = 0; fila < modeloInventarioDashboard.getRowCount(); fila++) {
                for (int columna = 0; columna < modeloInventarioDashboard.getColumnCount() - 1; columna++) {
                    if (columna > 0) {
                        escritor.write(",");
                    }
                    String dato = String.valueOf(modeloInventarioDashboard.getValueAt(fila, columna))
                            .replace("\"", "\"\"");
                    escritor.write("\"" + dato + "\"");
                }
                escritor.newLine();
            }
            JOptionPane.showMessageDialog(this, "Inventario exportado correctamente.");
        } catch (IOException ex) {
            mostrarError("No se pudo exportar el inventario.", ex);
        }
    }

    private void desactivarProducto(String codigo) {
        int confirmacion = JOptionPane.showConfirmDialog(this,
                "¿Deseas desactivar el producto " + codigo + "?",
                "Confirmar cambio de estado", JOptionPane.YES_NO_OPTION);
        if (confirmacion != JOptionPane.YES_OPTION) {
            return;
        }
        try {
            File archivo = obtenerArchivoInventarios();
            try (Workbook libro = cargarLibro(archivo)) {
                Sheet hoja = obtenerHojaInventarios(libro);
                int filaIndice = buscarFilaPorCodigo(hoja, codigo);
                if (filaIndice < 0) {
                    throw new IOException("El producto ya no existe.");
                }
                escribirCelda(hoja.getRow(filaIndice), 5, "0");
                guardarLibro(libro, archivo);
            }
            JOptionPane.showMessageDialog(this, "Producto desactivado correctamente.");
        } catch (IOException ex) {
            mostrarError("No se pudo desactivar el producto.", ex);
        }
    }

    private void asegurarArchivoInventarios() {
        try {
            inventarioRepository.asegurarArchivoInventarios();
        } catch (IllegalStateException ex) {
            mostrarError("No se pudo preparar el archivo de inventarios.", ex);
        }
    }//asegurarArchivoInventarios

    private int buscarFilaPorCodigo(Sheet hoja, String codigo) {
        for (int fila = 1; fila <= hoja.getLastRowNum(); fila++) {
            Row row = hoja.getRow(fila);
            if (row == null) {
                continue;
            }
            if (codigo.equalsIgnoreCase(leerCelda(row, 0))) {
                return fila;
            }
        }
        return -1;
    }//buscarFilaPorCodigo

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
        return inventarioRepository.asegurarEncabezados(hoja);
    }//asegurarEncabezados

    private Sheet obtenerHojaInventarios(Workbook libro) {
        return inventarioRepository.obtenerHojaInventarios(libro);
    }//obtenerHojaInventarios

    private Workbook cargarLibro(File archivo) throws IOException {
        return inventarioRepository.cargarLibro(archivo);
    }//cargarLibro

    private void guardarLibro(Workbook libro, File archivo) throws IOException {
        inventarioRepository.guardarLibro(libro, archivo);
    }//guardarLibro

    private File obtenerArchivoInventarios() {
        return inventarioRepository.obtenerArchivoInventarios();
    }//obtenerArchivoInventarios

    private String leerCelda(Row fila, int columna) {
        return inventarioRepository.leerCelda(fila, columna);
    }//leerCelda

    private void escribirCelda(Row fila, int columna, String valor) {
        inventarioRepository.escribirCelda(fila, columna, valor);
    }//escribirCelda

    private String normalizarEstado(String valor) {
        return InventarioService.normalizarEstado(valor);
    }//normalizarEstado

    private void mostrarError(String mensajeUsuario, Exception ex) {
        logger.log(java.util.logging.Level.SEVERE, mensajeUsuario, ex);
        JOptionPane.showMessageDialog(this, mensajeUsuario + "\nDetalle t\u00e9cnico: " + ex.getMessage());
    }//mostrarError

    private void prepararBoton(javax.swing.JButton boton, String texto, javax.swing.Icon icono,
            java.awt.Color fondo, java.awt.Color textoColor, java.awt.Color colorBorde) {
        boton.setUI(new BotonPlanoUI());
        boton.setText(texto);
        boton.setIcon(icono);
        boton.setIconTextGap(8);
        boton.setBackground(fondo);
        boton.setForeground(textoColor);
        EstiloAplicacion.aplicarTipografiaBoton(boton);
        boton.setFocusPainted(false);
        boton.setContentAreaFilled(false);
        boton.setOpaque(false);
        boton.setRolloverEnabled(true);
        boton.setBorderPainted(colorBorde != null);
        boton.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                new BordeRedondeado(9, colorBorde, 1.4f),
                javax.swing.BorderFactory.createEmptyBorder(9, 16, 9, 16)));
        int ancho = texto != null && texto.contains("Agregar") ? 150 : 140;
        boton.setPreferredSize(new java.awt.Dimension(ancho, 38));
        boton.setMinimumSize(new java.awt.Dimension(ancho, 38));
        boton.setCursor(java.awt.Cursor.getPredefinedCursor(java.awt.Cursor.HAND_CURSOR));
    }//prepararBoton

    private static java.awt.Color mezclarColor(java.awt.Color base, java.awt.Color destino, float proporcion) {
        int r = base.getRed() + Math.round((destino.getRed() - base.getRed()) * proporcion);
        int g = base.getGreen() + Math.round((destino.getGreen() - base.getGreen()) * proporcion);
        int b = base.getBlue() + Math.round((destino.getBlue() - base.getBlue()) * proporcion);
        return new java.awt.Color(
                Math.max(0, Math.min(255, r)),
                Math.max(0, Math.min(255, g)),
                Math.max(0, Math.min(255, b)));
    }//mezclarColor

    /**
     * ButtonUI plano que sustituye el bisel nativo de Windows por un relleno
     * redondeado propio (mismo patr\u00f3n que PROVEEDORES.java, para mantener la
     * consistencia visual entre pantallas).
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

    // ===================== Iconos vectoriales (Java2D) =====================
    private static javax.swing.Icon iconoCaja(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawRect(2, 5, t - 4, t - 7);
            g2.drawLine(2, 5, t / 2, 1);
            g2.drawLine(t - 2, 5, t / 2, 1);
            g2.drawLine(t / 2, 1, t / 2, t - 2);
        });
    }//iconoCaja

    private static javax.swing.Icon iconoOjo(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawOval(1, 4, t - 2, t - 8);
            g2.fillOval(t / 2 - 2, t / 2 - 2, 4, 4);
        });
    }

    private static javax.swing.Icon iconoAdvertencia(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            java.awt.Polygon triangulo = new java.awt.Polygon(
                    new int[]{t / 2, 1, t - 2, t - 2, 2, t - 2},
                    new int[]{1, t - 2, t - 2}, 3);
            g2.drawPolygon(triangulo);
            g2.drawLine(t / 2, 5, t / 2, t - 6);
            g2.drawLine(t / 2, t - 4, t / 2, t - 3);
        });
    }

    private static javax.swing.Icon iconoMonedas(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawOval(2, 7, t - 6, 8);
            g2.drawOval(5, 3, t - 6, 8);
            g2.drawLine(2, 11, 2, 15);
            g2.drawLine(5, 7, 5, 11);
        });
    }

    private static javax.swing.Icon iconoCamion(java.awt.Color color) {
        return new VectorIcon(18, color, (g2, t) -> {
            g2.drawRect(1, 5, 9, 7);
            g2.drawRect(10, 8, 6, 4);
            g2.drawOval(3, t - 4, 3, 3);
            g2.drawOval(11, t - 4, 3, 3);
        });
    }//iconoCamion

    private static javax.swing.Icon iconoLimpiar(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawLine(2, t - 2, t - 6, 4);
            g2.drawRect(t - 8, 2, 6, 6);
            g2.drawLine(2, t - 2, t - 2, t - 2);
        });
    }//iconoLimpiar

    private static javax.swing.Icon iconoGuardar(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawRoundRect(1, 1, t - 2, t - 2, 2, 2);
            g2.drawRect(4, 1, t - 8, 5);
            g2.drawRect(4, t - 6, t - 8, 5);
        });
    }//iconoGuardar

    private static javax.swing.Icon iconoBasura(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawLine(3, 4, t - 3, 4);
            g2.drawRect(4, 4, t - 8, t - 5);
            g2.drawLine(6, 1, t - 6, 1);
            g2.drawLine(6, 7, 6, t - 4);
            g2.drawLine(t / 2, 7, t / 2, t - 4);
            g2.drawLine(t - 6, 7, t - 6, t - 4);
        });
    }//iconoBasura

    private static javax.swing.Icon iconoPower(java.awt.Color color) {
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawArc(2, 2, t - 4, t - 4, -45, 270);
            g2.drawLine(t / 2, 1, t / 2, t / 2);
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
    }//iconoEditar

    private static javax.swing.Icon iconoCampo(java.awt.Color color) {
        return new VectorIcon(15, color, (g2, t) -> {
            g2.drawRoundRect(2, 3, t - 5, t - 6, 3, 3);
            g2.drawLine(5, 7, t - 5, 7);
            g2.drawLine(5, 10, t - 7, 10);
        });
    }//iconoCampo

    private static javax.swing.Icon iconoRadio(boolean marcado) {
        java.awt.Color color = marcado ? new java.awt.Color(0, 130, 60) : new java.awt.Color(160, 170, 180);
        return new VectorIcon(16, color, (g2, t) -> {
            g2.drawOval(2, 2, t - 4, t - 4);
            if (marcado) {
                g2.fillOval(5, 5, t - 10, t - 10);
            }
        });
    }//iconoRadio
}
