package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.EstiloAplicacion;
import com.mycompany.ocxrst.model.Tarea;
import com.mycompany.ocxrst.repository.TareaRepository;
import com.toedter.calendar.JDateChooser;
import java.awt.*;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.List;
import java.util.Locale;
import javax.swing.*;
import javax.swing.table.*;

public class MESATRABAJO extends javax.swing.JFrame {
    private static final Color AZUL = new Color(28, 82, 132);
    private static final Color AZUL_CLARO = new Color(31, 112, 184);
    private static final DateTimeFormatter FECHA = DateTimeFormatter.ofPattern("dd/MM/yyyy");
    private final TareaRepository repository = new TareaRepository();
    private final DefaultTableModel modelo = new DefaultTableModel(
            new String[]{"FOLIO", "DESCRIPCIÓN", "TIPO", "ESTADO", "PRIORIDAD",
                "RESPONSABLE", "FECHA CREACIÓN", "FECHA VENCIMIENTO", "ACCIONES"}, 0) {
        @Override public boolean isCellEditable(int row, int column) { return false; }
    };
    private final JTable tabla = new JTable(modelo);
    private final JTextField buscar = new JTextField();
    private final JComboBox<String> estado = new JComboBox<>(new String[]{"Todos", "Pendiente", "En proceso", "Completada", "Cancelada"});
    private final JComboBox<String> prioridad = new JComboBox<>(new String[]{"Todas", "Alta", "Media", "Baja"});
    private final JComboBox<String> responsable = new JComboBox<>(new String[]{"Todos"});
    private final JDateChooser fechaInicio = crearSelectorFecha();
    private final JDateChooser fechaFin = crearSelectorFecha();
    private final JLabel totalRegistros = new JLabel();
    private List<Tarea> tareas = List.of();
    private JPanel contenedorVista;
    private JPanel vistaBusqueda;

    public MESATRABAJO() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        setTitle("MESA DE TRABAJO");
        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        vistaBusqueda = crearContenido();
        contenedorVista = new JPanel(new BorderLayout());
        contenedorVista.add(vistaBusqueda, BorderLayout.CENTER);
        setContentPane(contenedorVista);
        setSize(1000, 680);
        setMinimumSize(new Dimension(850, 540));
        setLocationRelativeTo(null);
        cargar();
    }

    private JPanel crearContenido() {
        JPanel fondo = new JPanel(new BorderLayout(0, 12));
        fondo.setBackground(new Color(240, 246, 253));
        fondo.setBorder(BorderFactory.createEmptyBorder(24, 24, 20, 24));
        JPanel encabezado = new JPanel(new BorderLayout(12, 0));
        encabezado.setOpaque(false);
        JButton nueva = boton("Nueva tarea", "crear", true, e -> nuevaTarea(null));
        JButton filtrosBoton = boton("Filtros", "filtros", false, e -> buscar.requestFocusInWindow());
        JButton exportar = boton("Exportar", "exportar", false, e -> exportar());
        JButton actualizar = boton("Actualizar", "actualizar", false, e -> cargar());
        JPanel acciones = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 5));
        acciones.setOpaque(false); acciones.add(nueva); acciones.add(filtrosBoton); acciones.add(exportar); acciones.add(actualizar);
        encabezado.add(acciones, BorderLayout.EAST); fondo.add(encabezado, BorderLayout.NORTH);

        JPanel tarjeta = new JPanel(new BorderLayout(0, 12));
        tarjeta.setBackground(Color.WHITE);
        tarjeta.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(205, 218, 235)),
                BorderFactory.createEmptyBorder(16, 16, 16, 16)));
        JPanel filtros = new JPanel(new GridBagLayout()); filtros.setOpaque(false);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(0, 0, 0, 10);
        agregarFiltro(filtros, campoFiltro("Buscar", buscar), gbc, 0, 0.95);
        agregarFiltro(filtros, campoFiltro("Estado", estado), gbc, 1, 0.65);
        agregarFiltro(filtros, campoFiltro("Prioridad", prioridad), gbc, 2, 0.65);
        agregarFiltro(filtros, campoFiltro("Responsable", responsable), gbc, 3, 0.75);
        agregarFiltro(filtros, campoFiltro("Fecha inicio", fechaInicio), gbc, 4, 0.7);
        agregarFiltro(filtros, campoFiltro("Fecha fin", fechaFin), gbc, 5, 0.7);
        buscar.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
            public void insertUpdate(javax.swing.event.DocumentEvent e) { refrescarTabla(); }
            public void removeUpdate(javax.swing.event.DocumentEvent e) { refrescarTabla(); }
            public void changedUpdate(javax.swing.event.DocumentEvent e) { refrescarTabla(); }
        });
        estado.addActionListener(e -> refrescarTabla());
        prioridad.addActionListener(e -> refrescarTabla());
        responsable.addActionListener(e -> refrescarTabla());
        fechaInicio.addPropertyChangeListener("date", e -> refrescarTabla());
        fechaFin.addPropertyChangeListener("date", e -> refrescarTabla());
        tarjeta.add(filtros, BorderLayout.NORTH);
        configurarTabla();
        tarjeta.add(new JScrollPane(tabla), BorderLayout.CENTER);
        JPanel pie = new JPanel(new BorderLayout());
        pie.setOpaque(false);
        totalRegistros.setForeground(AZUL);
        totalRegistros.setFont(new Font("Segoe UI", Font.BOLD, 11));
        pie.add(totalRegistros, BorderLayout.EAST);
        tarjeta.add(pie, BorderLayout.SOUTH);
        fondo.add(tarjeta, BorderLayout.CENTER);
        return fondo;
    }

    private void agregarFiltro(JPanel panel, JComponent componente, GridBagConstraints base, int columna, double peso) {
        GridBagConstraints gbc = (GridBagConstraints) base.clone();
        gbc.gridx = columna;
        gbc.gridy = 0;
        gbc.weightx = peso;
        if (columna == 5) gbc.insets = new Insets(0, 0, 0, 0);
        panel.add(componente, gbc);
    }

    private JDateChooser crearSelectorFecha() {
        JDateChooser selector = new JDateChooser();
        selector.setLocale(Locale.forLanguageTag("es-MX"));
        selector.setDateFormatString("d MMM yyyy");
        selector.setPreferredSize(new Dimension(0, 28));
        selector.setToolTipText("Selecciona una fecha");
        return selector;
    }

    private JPanel campoFiltro(String texto, JComponent campo) {
        JPanel panel = new JPanel(new BorderLayout(0, 4)); panel.setOpaque(false);
        JLabel label = new JLabel(texto); label.setForeground(AZUL); label.setFont(new Font("Segoe UI", Font.BOLD, 11));
        panel.add(label, BorderLayout.NORTH); panel.add(campo, BorderLayout.CENTER); return panel;
    }

    private void configurarTabla() {
        tabla.setRowHeight(42); tabla.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        tabla.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 11));
        tabla.getTableHeader().setPreferredSize(new Dimension(0, 38));
        tabla.getTableHeader().setBackground(new Color(232, 238, 247)); tabla.getTableHeader().setForeground(AZUL);
        tabla.setGridColor(new Color(225, 232, 241));
        tabla.setSelectionBackground(new Color(225, 238, 250));
        tabla.setSelectionForeground(AZUL);
        tabla.getColumnModel().getColumn(0).setPreferredWidth(72);
        tabla.getColumnModel().getColumn(1).setPreferredWidth(220);
        tabla.getColumnModel().getColumn(8).setPreferredWidth(150);
        tabla.getColumnModel().getColumn(2).setCellRenderer(new BadgeRenderer());
        tabla.getColumnModel().getColumn(3).setCellRenderer(new BadgeRenderer());
        tabla.getColumnModel().getColumn(4).setCellRenderer(new BadgeRenderer());
        tabla.getColumnModel().getColumn(8).setCellRenderer(new ActionRenderer());
        tabla.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override public void mouseClicked(java.awt.event.MouseEvent e) {
                if (e.getClickCount() != 1) return;
                int row = tabla.rowAtPoint(e.getPoint()), col = tabla.columnAtPoint(e.getPoint());
                if (row < 0 || col != 8) return;
                String folio = String.valueOf(modelo.getValueAt(row, 0));
                int x = e.getX() - tabla.getCellRect(row, col, true).x;
                if (x < 42) ver(folio); else if (x < 84) editar(folio); else desactivar(folio);
            }
        });
    }

    private void cargar() {
        try {
            tareas = repository.listar();
            actualizarResponsables();
            refrescarTabla();
        }
        catch (IOException ex) { error("No se pudieron cargar las tareas.", ex); }
    }

    private void actualizarResponsables() {
        String seleccionado = String.valueOf(responsable.getSelectedItem());
        responsable.removeAllItems();
        responsable.addItem("Todos");
        tareas.stream().map(Tarea::responsable).filter(r -> r != null && !r.isBlank()).distinct().sorted()
                .forEach(responsable::addItem);
        responsable.setSelectedItem(seleccionado);
        if (responsable.getSelectedIndex() < 0) responsable.setSelectedIndex(0);
    }

    private void refrescarTabla() {
        modelo.setRowCount(0); String texto = buscar.getText().trim().toLowerCase(Locale.ROOT);
        String est = String.valueOf(estado.getSelectedItem()), pri = String.valueOf(prioridad.getSelectedItem());
        String resp = String.valueOf(responsable.getSelectedItem());
        LocalDate inicio = leerFecha(fechaInicio), fin = leerFecha(fechaFin);
        tareas.stream().filter(t -> texto.isEmpty() || (t.folio() + " " + t.titulo() + " " + t.descripcion()).toLowerCase(Locale.ROOT).contains(texto))
                .filter(t -> "Todos".equals(est) || t.estado().equals(est))
                .filter(t -> "Todas".equals(pri) || t.prioridad().equals(pri))
                .filter(t -> "Todos".equals(resp) || t.responsable().equals(resp))
                .filter(t -> inicio == null || !t.inicio().isBefore(inicio))
                .filter(t -> fin == null || !t.vencimiento().isAfter(fin))
                .forEach(t -> modelo.addRow(new Object[]{t.folio(), t.titulo(), t.tipo(), t.estado(), t.prioridad(),
                    t.responsable(), t.creacion().format(FECHA), t.vencimiento().format(FECHA), "VER  EDITAR  ELIMINAR"}));
        totalRegistros.setText("Total: " + modelo.getRowCount() + " registros");
    }

    private LocalDate leerFecha(JDateChooser campo) {
        if (campo.getDate() == null) return null;
        return campo.getDate().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    private void nuevaTarea(Tarea tarea) {
        NUEVATAREA dialogo = new NUEVATAREA(this, tarea, repository);
        JComponent formulario = dialogo.obtenerContenidoIntegrado(() -> volverAMesaTrabajo(dialogo));
        dialogo.setContentPane(new JPanel());
        dialogo.dispose();
        mostrarVistaInterna(formulario);
    }

    private void volverAMesaTrabajo(NUEVATAREA formulario) {
        cargar();
        mostrarVistaPrincipal();
    }

    private void mostrarVistaInterna(JComponent vista) {
        contenedorVista.removeAll();
        contenedorVista.setLayout(new BorderLayout());
        contenedorVista.add(vista, BorderLayout.CENTER);
        contenedorVista.revalidate();
        contenedorVista.repaint();
        Container padre = contenedorVista.getParent();
        if (padre != null) {
            padre.revalidate();
            padre.repaint();
        }
    }

    private void mostrarVistaPrincipal() {
        mostrarVistaInterna(vistaBusqueda);
    }

    private Tarea buscar(String folio) { return tareas.stream().filter(t -> t.folio().equals(folio)).findFirst().orElse(null); }
    private void ver(String folio) { Tarea t = buscar(folio); if (t != null) mostrarMensaje(t.descripcion().isBlank() ? t.titulo() : t.descripcion(), folio, JOptionPane.INFORMATION_MESSAGE); }
    private void editar(String folio) { Tarea t = buscar(folio); if (t != null) nuevaTarea(t); }
    private void desactivar(String folio) {
        String usuario = obtenerUsuarioActual();
        if (mostrarConfirmacion("¿Desactivar la tarea " + folio + "?",
                "Confirmar desactivación") != JOptionPane.YES_OPTION) return;
        try {
            repository.desactivar(folio, usuario);
            cargar();
            mostrarMensaje("La tarea " + folio + " fue desactivada correctamente.",
                    "Tarea desactivada", JOptionPane.INFORMATION_MESSAGE);
        }
        catch (IOException ex) { error("No se pudo desactivar la tarea.", ex); }
    }

    private String obtenerUsuarioActual() {
        String nombre = INICIARSESION.nombreUsuarioActual == null ? "" : INICIARSESION.nombreUsuarioActual.trim();
        if (!nombre.isEmpty()) return nombre;
        String usuario = INICIARSESION.usuarioActual == null ? "" : INICIARSESION.usuarioActual.trim();
        return usuario.isEmpty() ? "Usuario no identificado" : usuario;
    }
    private void exportar() {
        JFileChooser chooser = EstiloAplicacion.crearSelectorArchivos();
        chooser.setDialogTitle("Guardar");
        chooser.setSelectedFile(new java.io.File("TAREAS_EXPORTADAS.xlsx"));
        if (chooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) try { repository.exportar(tareas, chooser.getSelectedFile().toPath()); }
        catch (IOException ex) { error("No se pudo exportar.", ex); }
    }
    private void error(String mensaje, Exception ex) { mostrarMensaje(mensaje + "\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE); }

    private void mostrarMensaje(String mensaje, String titulo, int tipo) {
        JOptionPane panel = new JOptionPane(crearBarraMensaje(mensaje), tipo);
        prepararMensaje(panel);
        JDialog dialogo = panel.createDialog(this, titulo);
        dialogo.setVisible(true);
    }

    private int mostrarConfirmacion(String mensaje, String titulo) {
        JOptionPane panel = new JOptionPane(crearBarraMensaje(mensaje),
                JOptionPane.QUESTION_MESSAGE, JOptionPane.YES_NO_OPTION);
        prepararMensaje(panel);
        JDialog dialogo = panel.createDialog(this, titulo);
        dialogo.setVisible(true);
        Object valor = panel.getValue();
        return valor instanceof Integer ? (Integer) valor : JOptionPane.NO_OPTION;
    }

    private JPanel crearBarraMensaje(String mensaje) {
        JPanel barra = new JPanel(new BorderLayout());
        barra.setBackground(new Color(232, 235, 239));
        barra.setBorder(BorderFactory.createEmptyBorder(10, 14, 10, 14));
        JLabel texto = new JLabel("<html>" + escaparHtml(mensaje).replace("\n", "<br>") + "</html>");
        texto.setForeground(new Color(35, 42, 50));
        barra.add(texto, BorderLayout.CENTER);
        return barra;
    }

    private String escaparHtml(String texto) {
        return texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }

    private void prepararMensaje(JOptionPane panel) {
        panel.setBackground(Color.WHITE);
        for (Component componente : panel.getComponents()) {
            componente.setBackground(Color.WHITE);
        }
    }
    private JButton boton(String texto, String tipoIcono, boolean principal, ActionListener accion) {
        JButton b = new JButton(texto);
        b.setIcon(new IconoBoton(tipoIcono));
        b.setIconTextGap(8);
        b.setFont(new Font("Segoe UI", Font.BOLD, 11));
        b.setPreferredSize(new Dimension(principal ? 128 : 140, 46));
        b.setFocusPainted(false);
        b.setUI(new BotonPlanoUI());
        b.setContentAreaFilled(false);
        b.setOpaque(false);
        b.setRolloverEnabled(true);
        b.setCursor(new Cursor(Cursor.HAND_CURSOR));
        if (principal) {
            b.setBackground(AZUL_CLARO);
            b.setForeground(Color.WHITE);
            b.setBorder(BorderFactory.createCompoundBorder(
                    new BordeRedondeado(9, null, 1.4f),
                    BorderFactory.createEmptyBorder(9, 16, 9, 16)));
        } else {
            b.setBackground(Color.WHITE);
            b.setForeground(AZUL);
            b.setBorder(BorderFactory.createCompoundBorder(
                    new BordeRedondeado(9, AZUL_CLARO, 1.4f),
                    BorderFactory.createEmptyBorder(9, 16, 9, 16)));
        }
        b.addActionListener(accion);
        return b;
    }
    private static final class BordeRedondeado extends javax.swing.border.AbstractBorder {
        private final Color color;
        private final int radio;
        private final float grosor;

        private BordeRedondeado(int radio, Color color, float grosor) {
            this.radio = radio;
            this.color = color;
            this.grosor = grosor;
        }

        @Override public void paintBorder(Component componente, Graphics graphics, int x, int y, int width, int height) {
            if (color == null) return;
            Graphics2D g = (Graphics2D) graphics.create();
            try {
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g.setColor(color);
                g.setStroke(new BasicStroke(grosor));
                g.drawRoundRect(x + 1, y + 1, width - 3, height - 3, radio, radio);
            } finally {
                g.dispose();
            }
        }

        @Override public Insets getBorderInsets(Component componente) {
            return new Insets(2, 2, 2, 2);
        }

        @Override public boolean isBorderOpaque() {
            return false;
        }
    }
    private static final class BotonPlanoUI extends javax.swing.plaf.basic.BasicButtonUI {
        @Override public void paint(Graphics graphics, JComponent componente) {
            AbstractButton boton = (AbstractButton) componente;
            ButtonModel modelo = boton.getModel();
            Color fondo = boton.getBackground();
            if (modelo.isPressed()) fondo = mezclarColor(fondo, Color.BLACK, 0.16f);
            else if (modelo.isRollover()) fondo = mezclarColor(fondo, Color.BLACK, 0.07f);
            Graphics2D g = (Graphics2D) graphics.create();
            try {
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g.setColor(fondo);
                g.fillRoundRect(0, 0, componente.getWidth() - 1, componente.getHeight() - 1, 9, 9);
            } finally {
                g.dispose();
            }
            super.paint(graphics, componente);
        }
    }
    private static Color mezclarColor(Color base, Color destino, float proporcion) {
        return new Color(
                base.getRed() + Math.round((destino.getRed() - base.getRed()) * proporcion),
                base.getGreen() + Math.round((destino.getGreen() - base.getGreen()) * proporcion),
                base.getBlue() + Math.round((destino.getBlue() - base.getBlue()) * proporcion));
    }
    private static final class IconoBoton implements Icon {
        private final String tipo;

        private IconoBoton(String tipo) {
            this.tipo = tipo;
        }

        @Override public int getIconWidth() { return 16; }
        @Override public int getIconHeight() { return 16; }

        @Override public void paintIcon(Component componente, Graphics graphics, int x, int y) {
            Graphics2D g = (Graphics2D) graphics.create();
            try {
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g.setColor(componente.isEnabled() ? componente.getForeground() : new Color(150, 160, 170));
                g.setStroke(new BasicStroke(1.6f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                if ("crear".equals(tipo)) {
                    g.drawLine(x + 8, y + 3, x + 8, y + 13);
                    g.drawLine(x + 3, y + 8, x + 13, y + 8);
                } else if ("exportar".equals(tipo)) {
                    g.drawRect(x + 4, y + 2, 9, 12);
                    g.drawLine(x + 6, y + 6, x + 11, y + 6);
                    g.drawLine(x + 6, y + 9, x + 11, y + 9);
                    g.drawLine(x + 6, y + 12, x + 9, y + 12);
                } else {
                    g.drawArc(x + 2, y + 2, 12, 12, 42, 285);
                    g.drawLine(x + 13, y + 2, x + 13, y + 6);
                    g.drawLine(x + 13, y + 2, x + 9, y + 2);
                }
            } finally {
                g.dispose();
            }
        }
    }
    private static class BadgeRenderer extends DefaultTableCellRenderer {
        @Override public Component getTableCellRendererComponent(JTable t, Object v, boolean s, boolean f, int r, int c) {
            super.getTableCellRendererComponent(t, v, s, f, r, c); setHorizontalAlignment(CENTER); setOpaque(true);
            String value = String.valueOf(v); setForeground(Color.WHITE);
            setBackground(value.equals("Completada") || value.equals("Baja") ? new Color(43, 150, 83) :
                    value.equals("Pendiente") || value.equals("Media") ? new Color(220, 164, 25) :
                    value.equals("Cancelada") || value.equals("Alta") ? new Color(205, 66, 66) : AZUL_CLARO);
            return this;
        }
    }
    private static class ActionRenderer extends JPanel implements TableCellRenderer {
        private ActionRenderer() {
            setOpaque(true);
            setLayout(new FlowLayout(FlowLayout.CENTER, 5, 5));
        }

        @Override public Component getTableCellRendererComponent(JTable tabla, Object valor, boolean seleccionado,
                boolean tieneFoco, int fila, int columna) {
            removeAll();
            setBackground(seleccionado ? new Color(225, 238, 250) : Color.WHITE);
            add(botonAccion(new IconoAccion("ver"), Color.WHITE, AZUL));
            add(botonAccion(new IconoAccion("editar"), AZUL_CLARO, Color.WHITE));
            add(botonAccion(new IconoAccion("eliminar"), new Color(247, 78, 78), Color.WHITE));
            return this;
        }

        private JButton botonAccion(Icon icono, Color fondo, Color texto) {
            JButton boton = new JButton(icono);
            boton.setPreferredSize(new Dimension(38, 30));
            boton.setBackground(fondo);
            boton.setForeground(texto);
            boton.setBorder(BorderFactory.createLineBorder(
                    fondo.equals(Color.WHITE) ? new Color(198, 210, 225) : fondo));
            boton.setFocusPainted(false);
            boton.setContentAreaFilled(true);
            boton.setOpaque(true);
            return boton;
        }
    }

    private static final class IconoAccion implements Icon {
        private final String tipo;

        private IconoAccion(String tipo) {
            this.tipo = tipo;
        }

        @Override public int getIconWidth() { return 16; }
        @Override public int getIconHeight() { return 16; }

        @Override public void paintIcon(Component componente, Graphics graphics, int x, int y) {
            Graphics2D g = (Graphics2D) graphics.create();
            try {
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g.setColor(componente.getForeground());
                g.setStroke(new BasicStroke(1.7f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
                if ("ver".equals(tipo)) {
                    g.drawOval(x + 1, y + 4, 14, 8);
                    g.fillOval(x + 6, y + 6, 4, 4);
                } else if ("editar".equals(tipo)) {
                    g.rotate(-Math.PI / 4, x + 8, y + 8);
                    g.fillRoundRect(x + 6, y + 1, 4, 13, 2, 2);
                    g.fillRect(x + 5, y + 1, 6, 3);
                    g.rotate(Math.PI / 4, x + 8, y + 8);
                } else {
                    g.drawRoundRect(x + 4, y + 4, 8, 11, 1, 1);
                    g.drawLine(x + 3, y + 3, x + 13, y + 3);
                    g.drawLine(x + 6, y + 1, x + 10, y + 1);
                    g.drawLine(x + 7, y + 7, x + 7, y + 12);
                    g.drawLine(x + 10, y + 7, x + 10, y + 12);
                }
            } finally {
                g.dispose();
            }
        }
    }
}