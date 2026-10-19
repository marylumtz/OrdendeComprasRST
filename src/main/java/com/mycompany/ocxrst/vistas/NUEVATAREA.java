package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.model.Tarea;
import com.mycompany.ocxrst.repository.TareaRepository;
import java.awt.*;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import javax.swing.*;
import com.toedter.calendar.JDateChooser;

public class NUEVATAREA extends JDialog {
    private static final Color AZUL = new Color(7, 57, 127);
    private static final Color AZUL_ACCION = new Color(10, 116, 211);
    private static final Color AZUL_SECCION = new Color(232, 243, 253);
    private static final Color BORDE = new Color(190, 208, 229);
    private static final Color TEXTO = new Color(24, 58, 104);
    private final TareaRepository repository;
    private final Tarea original;
    private boolean guardado;
    private Runnable cerrarAccion = this::dispose;
    private final JTextField titulo = new JTextField();
    private final JTextArea descripcion = new JTextArea();
    private final JComboBox<String> tipo = new JComboBox<>(Tarea.TIPOS.toArray(String[]::new));
    private final JComboBox<String> estado = new JComboBox<>(Tarea.ESTADOS.toArray(String[]::new));
    private final JComboBox<String> prioridad = new JComboBox<>(Tarea.PRIORIDADES.toArray(String[]::new));
    private final JTextField responsable = new JTextField();
    private final JDateChooser inicio = crearSelectorFecha();
    private final JDateChooser vencimiento = crearSelectorFecha();
    private final JTextField relacionado = new JTextField();
    private final JTextField creadoPor = new JTextField();
    private final JTextArea observaciones = new JTextArea();
    private final java.util.List<String> adjuntos = new ArrayList<>();
    private final JLabel adjuntosTexto = new JLabel("Arrastra archivos aquí o haz clic para seleccionar");
    private String estadoInicial;

    public NUEVATAREA(Window owner, Tarea original, TareaRepository repository) {
        super(owner, "NUEVA TAREA", ModalityType.APPLICATION_MODAL);
        this.original = original; this.repository = repository;
        setUndecorated(true);
        setResizable(false);
        setContentPane(crearContenido()); setSize(900, 700); setLocationRelativeTo(owner);
        if (original != null) cargarOriginal();
        configurarCamposProtegidos();
        estadoInicial = estadoActual();
    }

    private JPanel crearContenido() {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setBackground(new Color(248, 251, 255));
        panel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(190, 202, 218)),
                BorderFactory.createEmptyBorder(0, 12, 10, 12)));

        JPanel encabezado = new JPanel(new BorderLayout(12, 0));
        encabezado.setBackground(AZUL);
        encabezado.setBorder(BorderFactory.createEmptyBorder(9, 16, 9, 12));
        JLabel icono = new JLabel("▣");
        icono.setForeground(Color.WHITE);
        icono.setFont(new Font("Segoe UI Symbol", Font.BOLD, 25));
        JLabel tituloEncabezado = new JLabel(original == null ? "Nueva tarea" : "Editar tarea");
        tituloEncabezado.setForeground(Color.WHITE);
        tituloEncabezado.setFont(new Font("Segoe UI", Font.BOLD, 17));
        JLabel descripcionEncabezado = new JLabel("Registra y asigna una nueva actividad en la mesa de trabajo.");
        descripcionEncabezado.setForeground(new Color(225, 237, 255));
        descripcionEncabezado.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        JPanel textos = new JPanel();
        textos.setOpaque(false);
        textos.setLayout(new BoxLayout(textos, BoxLayout.Y_AXIS));
        textos.add(tituloEncabezado);
        textos.add(descripcionEncabezado);
        encabezado.add(icono, BorderLayout.WEST);
        encabezado.add(textos, BorderLayout.CENTER);
        JButton cerrar = new JButton("ATRÁS");
        cerrar.setBorderPainted(false); cerrar.setContentAreaFilled(false); cerrar.setFocusPainted(false);
        cerrar.setForeground(Color.WHITE); cerrar.setFont(new Font("Segoe UI", Font.BOLD, 11));
        cerrar.addActionListener(e -> cerrarFormulario());
        encabezado.add(cerrar, BorderLayout.EAST);
        panel.add(encabezado, BorderLayout.NORTH);

        JPanel contenido = new JPanel();
        contenido.setOpaque(false);
        contenido.setLayout(new BoxLayout(contenido, BoxLayout.Y_AXIS));
        contenido.add(seccion("▣", "Información general", crearInformacionGeneral()));
        contenido.add(Box.createVerticalStrut(7));
        contenido.add(seccion("♟", "Asignación y seguimiento", crearAsignacion()));
        contenido.add(Box.createVerticalStrut(7));
        JPanel inferiores = new JPanel(new GridLayout(1, 2, 8, 0));
        inferiores.setOpaque(false);
        inferiores.add(seccion("▤", "Subtareas (opcional)", crearSubtareas()));
        inferiores.add(seccion("⌕", "Archivos adjuntos (opcional)", crearAdjuntos()));
        contenido.add(inferiores);
        contenido.add(Box.createVerticalStrut(7));
        contenido.add(seccion("▤", "Observaciones (opcional)", crearObservaciones()));
        panel.add(contenido, BorderLayout.CENTER);

        JPanel botones = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
        botones.setOpaque(false);
        JButton cancelar = boton("×  Cancelar", false);
        cancelar.addActionListener(e -> cerrarFormulario());
        JButton guardar = boton("▣  Guardar tarea", true);
        guardar.addActionListener(e -> guardar());
        botones.add(cancelar); botones.add(guardar);
        panel.add(botones, BorderLayout.SOUTH);
        estilizarCampos();
        return panel;
    }

    private JPanel crearInformacionGeneral() {
        JPanel panel = rejilla(2);
        agregar(panel, 0, 0, "Título / Descripción breve *", titulo, 1);
        agregar(panel, 1, 0, "Tipo de tarea *", tipo, 1);
        agregarArea(panel, 0, 1, "Descripción detallada", descripcion, 1);
        JPanel folio = new JPanel(new BorderLayout(5, 0)); folio.setOpaque(false);
        folio.add(relacionado, BorderLayout.CENTER);
        JButton buscarFolio = boton("⌕", true); buscarFolio.setPreferredSize(new Dimension(38, 28));
        folio.add(buscarFolio, BorderLayout.EAST);
        agregar(panel, 1, 1, "Folio / Documento relacionado", folio, 1);
        return panel;
    }

    private JPanel crearAsignacion() {
        JPanel panel = rejilla(4);
        agregar(panel, 0, 0, "Prioridad *", prioridad, 1);
        agregar(panel, 1, 0, "Estado *", estado, 1);
        agregar(panel, 2, 0, "Responsable *", responsable, 1);
        agregar(panel, 3, 0, "Creado por", creadoPor, 1);
        agregar(panel, 0, 1, "Fecha de inicio *", inicio, 1);
        agregar(panel, 1, 1, "Fecha de vencimiento *", vencimiento, 1);
        return panel;
    }

    private JPanel crearSubtareas() {
        JPanel panel = new JPanel(new BorderLayout(0, 5)); panel.setOpaque(false);
        JPanel fila = new JPanel(new BorderLayout(5, 0)); fila.setOpaque(false);
        JTextField subtarea = new JTextField(); subtarea.setToolTipText("Escribe una subtarea");
        fila.add(new JLabel("+"), BorderLayout.WEST); fila.add(subtarea, BorderLayout.CENTER);
        JButton agregar = boton("Agregar", false);
        agregar.addActionListener(e -> { if (!subtarea.getText().isBlank()) subtarea.setText(""); });
        fila.add(agregar, BorderLayout.EAST);
        panel.add(fila, BorderLayout.NORTH);
        JLabel nota = new JLabel("Las subtareas se pueden agregar al editar la tarea.");
        nota.setForeground(new Color(100, 121, 147)); nota.setFont(new Font("Segoe UI", Font.PLAIN, 10));
        panel.add(nota, BorderLayout.SOUTH);
        return panel;
    }

    private JPanel crearAdjuntos() {
        JPanel panel = new JPanel(new BorderLayout(0, 5)); panel.setOpaque(false);
        adjuntosTexto.setHorizontalAlignment(SwingConstants.CENTER);
        adjuntosTexto.setForeground(new Color(52, 92, 140));
        adjuntosTexto.setFont(new Font("Segoe UI", Font.BOLD, 11));
        JPanel drop = new JPanel(new BorderLayout());
        drop.setBackground(new Color(249, 252, 255));
        drop.setBorder(BorderFactory.createDashedBorder(BORDE, 1, 4, 2, true));
        drop.add(adjuntosTexto, BorderLayout.CENTER);
        JButton seleccionar = boton("Seleccionar archivos", false);
        seleccionar.addActionListener(e -> seleccionarAdjuntos());
        panel.add(drop, BorderLayout.CENTER);
        panel.add(seleccionar, BorderLayout.SOUTH);
        return panel;
    }

    private JPanel crearObservaciones() {
        JPanel panel = new JPanel(new BorderLayout()); panel.setOpaque(false);
        agregarArea(panel, "Información adicional, comentarios o notas...", observaciones);
        return panel;
    }

    private JPanel rejilla(int columnas) {
        JPanel panel = new JPanel(new GridBagLayout()); panel.setOpaque(false);
        panel.putClientProperty("columnas", columnas);
        return panel;
    }

    private JPanel seccion(String icono, String tituloSeccion, JComponent contenido) {
        JPanel panel = new JPanel(new BorderLayout(0, 6)); panel.setOpaque(false);
        JPanel encabezado = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 4));
        encabezado.setBackground(AZUL_SECCION);
        encabezado.setBorder(BorderFactory.createEmptyBorder(0, 5, 0, 5));
        JLabel iconoLabel = new JLabel(icono); iconoLabel.setForeground(AZUL); iconoLabel.setFont(new Font("Segoe UI Symbol", Font.BOLD, 16));
        JLabel texto = new JLabel(tituloSeccion); texto.setForeground(AZUL); texto.setFont(new Font("Segoe UI", Font.BOLD, 12));
        encabezado.add(iconoLabel); encabezado.add(texto);
        panel.add(encabezado, BorderLayout.NORTH);
        panel.add(contenido, BorderLayout.CENTER);
        return panel;
    }

    private void agregar(JPanel p, int x, int y, String label, Component input, int span) {
        int columnas = (int) p.getClientProperty("columnas");
        GridBagConstraints c = new GridBagConstraints();
        c.gridx = x; c.gridy = y; c.gridwidth = span; c.weightx = span;
        c.fill = GridBagConstraints.HORIZONTAL; c.anchor = GridBagConstraints.NORTH;
        c.insets = new Insets(0, x == 0 ? 0 : 6, 6, 0);
        JPanel box = new JPanel(new BorderLayout(0, 4)); box.setOpaque(false);
        JLabel l = new JLabel(label); l.setFont(new Font("Segoe UI", Font.BOLD, 10)); l.setForeground(TEXTO);
        box.add(l, BorderLayout.NORTH); box.add(input, BorderLayout.CENTER); p.add(box, c);
        if (x + span == columnas) c.insets.right = 0;
    }

    private void agregarArea(JPanel p, int x, int y, String label, JTextArea area, int span) {
        area.setLineWrap(true); area.setWrapStyleWord(true);
        JScrollPane scroll = new JScrollPane(area); scroll.setPreferredSize(new Dimension(0, 52));
        agregar(p, x, y, label, scroll, span);
    }

    private void agregarArea(JPanel p, String textoInicial, JTextArea area) {
        area.setLineWrap(true); area.setWrapStyleWord(true); area.setToolTipText(textoInicial);
        JScrollPane scroll = new JScrollPane(area); scroll.setPreferredSize(new Dimension(0, 45));
        p.add(scroll, BorderLayout.CENTER);
    }

    private JButton boton(String texto, boolean primario) {
        JButton boton = new JButton(texto);
        boton.setFont(new Font("Segoe UI", Font.BOLD, 11));
        boton.setFocusPainted(false);
        boton.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(primario ? AZUL_ACCION : BORDE),
                BorderFactory.createEmptyBorder(6, 12, 6, 12)));
        boton.setBackground(primario ? AZUL_ACCION : Color.WHITE);
        boton.setForeground(primario ? Color.WHITE : TEXTO);
        return boton;
    }

    private void estilizarCampos() {
        for (JComponent campo : new JComponent[]{titulo, responsable, relacionado, tipo, estado, prioridad, inicio, vencimiento}) {
            campo.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        }
        titulo.setToolTipText("Escribe un título claro y conciso");
        descripcion.setToolTipText("Describe la tarea, los objetivos y cualquier información relevante");
        relacionado.setToolTipText("Busca un folio, orden o proveedor");
        estiloArea(descripcion); estiloArea(observaciones);
    }

    private void configurarCamposProtegidos() {
        boolean administrador = esAdministrador();
        String nombre = INICIARSESION.nombreUsuarioActual == null
                ? "" : INICIARSESION.nombreUsuarioActual.trim();
        if (nombre.isBlank()) {
            nombre = INICIARSESION.usuarioActual == null ? "" : INICIARSESION.usuarioActual.trim();
        }
        creadoPor.setText(nombre);
        creadoPor.setToolTipText(administrador
                ? "El administrador puede modificar este valor."
                : "Este valor solo puede modificarlo el administrador.");
        creadoPor.setEditable(administrador);
        creadoPor.setEnabled(administrador);
        creadoPor.setForeground(administrador ? TEXTO : new Color(95, 108, 125));

        inicio.setToolTipText(administrador
                ? "El administrador puede capturar la fecha manualmente."
                : "La fecha de inicio se asigna automáticamente y solo puede modificarla el administrador.");
        inicio.setEnabled(administrador);
        inicio.getDateEditor().getUiComponent().setEnabled(administrador);
        if (inicio.getDateEditor().getUiComponent() instanceof JTextField editor) {
            editor.setEditable(administrador);
        }
        if (!administrador && inicio.getDate() == null) {
            inicio.setDate(Date.from(LocalDate.now().atStartOfDay(ZoneId.systemDefault()).toInstant()));
        }
    }

    private boolean esAdministrador() {
        return "admin".equalsIgnoreCase(INICIARSESION.usuarioActual == null
                ? "" : INICIARSESION.usuarioActual.trim());
    }

    private void estiloArea(JTextArea area) {
        area.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        area.setBorder(BorderFactory.createEmptyBorder(6, 8, 6, 8));
    }

    private void seleccionarAdjuntos() {
        JFileChooser selector = new JFileChooser();
        selector.setMultiSelectionEnabled(true);
        if (selector.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;
        for (java.io.File archivo : selector.getSelectedFiles()) adjuntos.add(archivo.getAbsolutePath());
        adjuntosTexto.setText(adjuntos.isEmpty() ? "Arrastra archivos aquí o haz clic para seleccionar"
                : adjuntos.size() + " archivo(s) seleccionado(s)");
    }
    private void cargarOriginal() {
        titulo.setText(original.titulo()); descripcion.setText(original.descripcion()); tipo.setSelectedItem(original.tipo()); estado.setSelectedItem(original.estado()); prioridad.setSelectedItem(original.prioridad()); responsable.setText(original.responsable()); inicio.setDate(Date.from(original.inicio().atStartOfDay(ZoneId.systemDefault()).toInstant())); vencimiento.setDate(Date.from(original.vencimiento().atStartOfDay(ZoneId.systemDefault()).toInstant())); relacionado.setText(original.folioRelacionado()); observaciones.setText(original.observaciones()); adjuntos.addAll(original.adjuntos());
    }
    private void guardar() {
        try {
            LocalDate desde = fecha(inicio), hasta = fecha(vencimiento);
            Tarea tarea = new Tarea(original == null ? "" : original.folio(), titulo.getText(), (String)tipo.getSelectedItem(), (String)estado.getSelectedItem(), (String)prioridad.getSelectedItem(), responsable.getText(), original == null ? LocalDate.now() : original.creacion(), hasta, descripcion.getText(), desde, relacionado.getText(), observaciones.getText(), adjuntos);
            repository.guardar(tarea); guardado = true; cerrarFormulario();
        } catch (IllegalArgumentException | IOException ex) { JOptionPane.showMessageDialog(this, ex.getMessage(), "No se pudo guardar", JOptionPane.WARNING_MESSAGE); }
    }

    public JComponent obtenerContenidoIntegrado(Runnable alCerrar) {
        cerrarAccion = alCerrar == null ? this::dispose : alCerrar;
        return (JComponent) getContentPane();
    }

    private void cerrarFormulario() {
        if (guardado || !formularioModificado()) {
            cerrarAccion.run();
            return;
        }

        Object[] opciones = {"GUARDAR", "NO GUARDAR", "CANCELAR"};
        int respuesta = JOptionPane.showOptionDialog(this,
                "El formulario tiene cambios sin guardar.\n¿Deseas guardar los cambios?",
                "Cambios sin guardar", JOptionPane.YES_NO_CANCEL_OPTION,
                JOptionPane.QUESTION_MESSAGE, null, opciones, opciones[0]);
        if (respuesta == JOptionPane.YES_OPTION) {
            guardar();
        } else if (respuesta == JOptionPane.NO_OPTION) {
            cerrarAccion.run();
        }
    }

    private boolean formularioModificado() {
        return estadoInicial == null || !estadoInicial.equals(estadoActual());
    }

    private String estadoActual() {
        String tipoSeleccionado = String.valueOf(tipo.getSelectedItem());
        String estadoSeleccionado = String.valueOf(estado.getSelectedItem());
        String prioridadSeleccionada = String.valueOf(prioridad.getSelectedItem());
        return String.join("\u001F",
                titulo.getText(), descripcion.getText(), tipoSeleccionado,
                estadoSeleccionado, prioridadSeleccionada, responsable.getText(),
                String.valueOf(inicio.getDate()), String.valueOf(vencimiento.getDate()),
                relacionado.getText(), observaciones.getText(),
                String.join("\u001E", adjuntos));
    }
    private JDateChooser crearSelectorFecha() {
        JDateChooser selector = new JDateChooser();
        selector.setLocale(new java.util.Locale("es", "MX"));
        selector.setDateFormatString("dd/MM/yyyy");
        selector.setPreferredSize(new Dimension(0, 28));
        return selector;
    }
    private LocalDate fecha(JDateChooser selector) {
        if (selector.getDate() == null) {
            throw new IllegalArgumentException("Completa las fechas de inicio y vencimiento.");
        }
        return selector.getDate().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }
    public boolean guardado() { return guardado; }
}