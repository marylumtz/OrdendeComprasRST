package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.EstiloAplicacion;
import com.mycompany.ocxrst.model.Tarea;
import com.mycompany.ocxrst.repository.TareaRepository;
import java.awt.*;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
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
    private List<Tarea> tareas = List.of();

    public MESATRABAJO() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        setTitle("MESA DE TRABAJO");
        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        setContentPane(crearContenido());
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
        JLabel icono = new JLabel(new IconoMesaTrabajo(), SwingConstants.CENTER);
        icono.setOpaque(true); icono.setBackground(AZUL_CLARO); icono.setForeground(Color.WHITE);
        icono.setPreferredSize(new Dimension(52, 52));
        encabezado.add(icono, BorderLayout.WEST);
        JPanel titulos = new JPanel();
        titulos.setOpaque(false); titulos.setLayout(new BoxLayout(titulos, BoxLayout.Y_AXIS));
        JLabel titulo = new JLabel("MESA DE TRABAJO");
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 24)); titulo.setForeground(AZUL);
        JLabel subtitulo = new JLabel("Gestión y seguimiento de actividades, solicitudes y pendientes");
        subtitulo.setForeground(new Color(100, 120, 145));
        titulos.add(titulo); titulos.add(subtitulo); encabezado.add(titulos, BorderLayout.CENTER);
        JButton nueva = boton("CREAR", "crear", true, e -> nuevaTarea(null));
        JButton exportar = boton("EXPORTAR", "exportar", false, e -> exportar());
        JButton actualizar = boton("ACTUALIZAR", "actualizar", false, e -> cargar());
        JPanel acciones = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 5));
        acciones.setOpaque(false); acciones.add(nueva); acciones.add(exportar); acciones.add(actualizar);
        encabezado.add(acciones, BorderLayout.EAST); fondo.add(encabezado, BorderLayout.NORTH);

        JPanel tarjeta = new JPanel(new BorderLayout(0, 12));
        tarjeta.setBackground(Color.WHITE);
        tarjeta.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(205, 218, 235)),
                BorderFactory.createEmptyBorder(16, 16, 16, 16)));
        JPanel filtros = new JPanel(new GridBagLayout()); filtros.setOpaque(false);
        GridBagConstraints gbc = new GridBagConstraints(); gbc.insets = new Insets(0, 0, 0, 10);
        gbc.fill = GridBagConstraints.HORIZONTAL; gbc.weightx = 1;
        filtros.add(campoFiltro("Buscar", buscar), gbc);
        gbc.weightx = 0.6; filtros.add(campoFiltro("Estado", estado), gbc);
        filtros.add(campoFiltro("Prioridad", prioridad), gbc);
        buscar.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
            public void insertUpdate(javax.swing.event.DocumentEvent e) { refrescarTabla(); }
            public void removeUpdate(javax.swing.event.DocumentEvent e) { refrescarTabla(); }
            public void changedUpdate(javax.swing.event.DocumentEvent e) { refrescarTabla(); }
        });
        estado.addActionListener(e -> refrescarTabla()); prioridad.addActionListener(e -> refrescarTabla());
        tarjeta.add(filtros, BorderLayout.NORTH);
        configurarTabla();
        tarjeta.add(new JScrollPane(tabla), BorderLayout.CENTER);
        fondo.add(tarjeta, BorderLayout.CENTER);
        return fondo;
    }

    private JPanel campoFiltro(String texto, JComponent campo) {
        JPanel panel = new JPanel(new BorderLayout(0, 4)); panel.setOpaque(false);
        JLabel label = new JLabel(texto); label.setForeground(AZUL); label.setFont(new Font("Segoe UI", Font.BOLD, 11));
        panel.add(label, BorderLayout.NORTH); panel.add(campo, BorderLayout.CENTER); return panel;
    }

    private void configurarTabla() {
        tabla.setRowHeight(36); tabla.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        tabla.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 11));
        tabla.getTableHeader().setBackground(new Color(232, 238, 247)); tabla.getTableHeader().setForeground(AZUL);
        tabla.setGridColor(new Color(225, 232, 241));
        tabla.setSelectionBackground(new Color(225, 238, 250));
        tabla.setSelectionForeground(AZUL);
        tabla.getColumnModel().getColumn(0).setPreferredWidth(72);
        tabla.getColumnModel().getColumn(1).setPreferredWidth(220);
        tabla.getColumnModel().getColumn(8).setPreferredWidth(190);
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
        try { tareas = repository.listar(); refrescarTabla(); }
        catch (IOException ex) { error("No se pudieron cargar las tareas.", ex); }
    }

    private void refrescarTabla() {
        modelo.setRowCount(0); String texto = buscar.getText().trim().toLowerCase(Locale.ROOT);
        String est = String.valueOf(estado.getSelectedItem()), pri = String.valueOf(prioridad.getSelectedItem());
        tareas.stream().filter(t -> texto.isEmpty() || (t.folio() + " " + t.titulo() + " " + t.descripcion()).toLowerCase(Locale.ROOT).contains(texto))
                .filter(t -> "Todos".equals(est) || t.estado().equals(est))
                .filter(t -> "Todas".equals(pri) || t.prioridad().equals(pri))
                .forEach(t -> modelo.addRow(new Object[]{t.folio(), t.titulo(), t.tipo(), t.estado(), t.prioridad(),
                    t.responsable(), t.creacion().format(FECHA), t.vencimiento().format(FECHA), "VER  EDITAR  DESACTIVAR"}));
    }

    private void nuevaTarea(Tarea tarea) {
        NUEVATAREA dialogo = new NUEVATAREA(this, tarea, repository);
        dialogo.setVisible(true); if (dialogo.guardado()) cargar();
    }

    private Tarea buscar(String folio) { return tareas.stream().filter(t -> t.folio().equals(folio)).findFirst().orElse(null); }
    private void ver(String folio) { Tarea t = buscar(folio); if (t != null) JOptionPane.showMessageDialog(this, t.descripcion().isBlank() ? t.titulo() : t.descripcion(), folio, JOptionPane.INFORMATION_MESSAGE); }
    private void editar(String folio) { Tarea t = buscar(folio); if (t != null) nuevaTarea(t); }
    private void desactivar(String folio) {
        if (JOptionPane.showConfirmDialog(this, "¿Desactivar la tarea " + folio + "? El registro se conservará.",
                "Confirmar desactivación", JOptionPane.YES_NO_OPTION) != JOptionPane.YES_OPTION) return;
        try { repository.desactivar(folio); cargar(); }
        catch (IOException ex) { error("No se pudo desactivar la tarea.", ex); }
    }
    private void exportar() {
        JFileChooser chooser = EstiloAplicacion.crearSelectorArchivos();
        chooser.setDialogTitle("Guardar");
        chooser.setSelectedFile(new java.io.File("TAREAS_EXPORTADAS.xlsx"));
        if (chooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) try { repository.exportar(tareas, chooser.getSelectedFile().toPath()); }
        catch (IOException ex) { error("No se pudo exportar.", ex); }
    }
    private void error(String mensaje, Exception ex) { JOptionPane.showMessageDialog(this, mensaje + "\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE); }
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
    private static final class IconoMesaTrabajo implements Icon {
        @Override public int getIconWidth() { return 20; }
        @Override public int getIconHeight() { return 20; }

        @Override public void paintIcon(Component componente, Graphics graphics, int x, int y) {
            Graphics2D g = (Graphics2D) graphics.create();
            try {
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g.setColor(componente.getForeground());
                for (int fila = 0; fila < 2; fila++) {
                    for (int columna = 0; columna < 2; columna++) {
                        g.fillRoundRect(x + 2 + columna * 8, y + 2 + fila * 8, 6, 6, 2, 2);
                    }
                }
            } finally {
                g.dispose();
            }
        }
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
    private static class ActionRenderer extends DefaultTableCellRenderer {
        @Override public Component getTableCellRendererComponent(JTable t, Object v, boolean s, boolean f, int r, int c) {
            super.getTableCellRendererComponent(t, v, s, f, r, c); setHorizontalAlignment(CENTER); setForeground(AZUL); setFont(new Font("Segoe UI", Font.BOLD, 10)); return this;
        }
    }
}
