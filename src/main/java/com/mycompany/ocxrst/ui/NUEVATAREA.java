package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.util.EstiloAplicacion;
import com.mycompany.ocxrst.model.Tarea;
import com.mycompany.ocxrst.repository.TareaRepository;
import com.mycompany.ocxrst.service.SesionActual;
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
    private static final Color TEXTO_ETIQUETA = new Color(34, 76, 119);
    private static final Color BORDE_CAMPO = new Color(198, 210, 225);
    private static final Color FONDO_SOLO_LECTURA = new Color(247, 250, 253);
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
        EstiloAplicacion.aplicarTipografiaGlobal();
        this.original = original; this.repository = repository;
        setUndecorated(true);
        setResizable(false);
        setContentPane(crearContenido()); setSize(980, 720); setLocationRelativeTo(owner);
        if (original != null) cargarOriginal();
        configurarCamposProtegidos();
        estadoInicial = estadoActual();
        EstiloAplicacion.aplicarEstiloBotones((Container) getContentPane());
        EstiloAplicacion.aplicarTipografia(getContentPane());
        EstiloAplicacion.aplicarEstiloCamposTexto(getContentPane());
    }

    private JPanel crearContenido() {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setBackground(EstiloAplicacion.FONDO_PANTALLA);
        panel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(new Color(190, 205, 224)),
                BorderFactory.createEmptyBorder(0, 14, 12, 14)));

        JPanel encabezado = new JPanel(new BorderLayout(12, 0));
        encabezado.setBackground(AZUL);
        encabezado.setBorder(BorderFactory.createEmptyBorder(10, 16, 10, 12));
        JLabel icono = new JLabel(new IconoTarea("editar", 24));
        icono.setForeground(Color.WHITE);
        icono.setFont(new Font("Segoe UI Symbol", Font.BOLD, 25));
        JLabel tituloEncabezado = new JLabel(original == null ? "Nueva tarea" : "Editar tarea");
        tituloEncabezado.setForeground(Color.WHITE);
        tituloEncabezado.setFont(new Font("Segoe UI", Font.BOLD, 18));
        JLabel descripcionEncabezado = new JLabel(original == null
                ? "Registra una nueva actividad y asigna su seguimiento."
                : "Actualiza la información y el seguimiento de la actividad.");
        descripcionEncabezado.setForeground(new Color(225, 237, 255));
        descripcionEncabezado.setFont(new Font("Segoe UI", Font.PLAIN, 11));
        JPanel textos = new JPanel();
        textos.setOpaque(false);
        textos.setLayout(new BoxLayout(textos, BoxLayout.Y_AXIS));
        textos.add(tituloEncabezado);
        textos.add(descripcionEncabezado);
        encabezado.add(icono, BorderLayout.WEST);
        encabezado.add(textos, BorderLayout.CENTER);
        JButton cerrar = new JButton("<html><u>ATRÁS</u></html>");
        cerrar.setIcon(new IconoTarea("atras", 16));
        cerrar.setIconTextGap(6);
        cerrar.setBorderPainted(false); cerrar.setContentAreaFilled(false); cerrar.setFocusPainted(false);
        cerrar.setForeground(Color.WHITE); cerrar.setFont(new Font("Segoe UI", Font.BOLD, 11));
        cerrar.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        cerrar.setToolTipText("Volver a la mesa de trabajo");
        cerrar.addActionListener(e -> cerrarFormulario());
        encabezado.add(cerrar, BorderLayout.EAST);
        panel.add(encabezado, BorderLayout.NORTH);

        JPanel contenido = new JPanel();
        contenido.setOpaque(false);
        contenido.setLayout(new BoxLayout(contenido, BoxLayout.Y_AXIS));
        contenido.add(seccion("informacion", "Información general", crearInformacionGeneral()));
        contenido.add(Box.createVerticalStrut(7));
        contenido.add(seccion("asignacion", "Asignación y seguimiento", crearAsignacion()));
        contenido.add(Box.createVerticalStrut(7));
        JPanel inferiores = new JPanel(new GridLayout(1, 2, 8, 0));
        inferiores.setOpaque(false);
        inferiores.add(seccion("subtareas", "Subtareas (opcional)", crearSubtareas()));
        inferiores.add(seccion("adjuntos", "Archivos adjuntos (opcional)", crearAdjuntos()));
        contenido.add(inferiores);
        contenido.add(Box.createVerticalStrut(7));
        contenido.add(seccion("observaciones", "Observaciones (opcional)", crearObservaciones()));
        JScrollPane desplazamiento = new JScrollPane(contenido);
        desplazamiento.setBorder(null);
        desplazamiento.setOpaque(false);
        desplazamiento.getViewport().setOpaque(false);
        desplazamiento.getVerticalScrollBar().setUnitIncrement(14);
        panel.add(desplazamiento, BorderLayout.CENTER);

        JPanel botones = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
        botones.setOpaque(false);
        JButton guardar = boton("Guardar tarea", true, "guardar");
        guardar.addActionListener(e -> guardar());
        botones.add(guardar);
        panel.add(botones, BorderLayout.SOUTH);
        estilizarCampos();
        return panel;
    }

    private JPanel crearInformacionGeneral() {
        JPanel panel = rejilla(2);
        agregar(panel, 0, 0, "TÍTULO / DESCRIPCIÓN BREVE *", titulo, 1);
        agregar(panel, 1, 0, "TIPO DE TAREA *", tipo, 1);
        agregarArea(panel, 0, 1, "Descripción detallada", descripcion, 1);
        JPanel folio = new JPanel(new BorderLayout(5, 0)); folio.setOpaque(false);
        folio.add(relacionado, BorderLayout.CENTER);
        JButton buscarFolio = boton("", true, "buscar"); buscarFolio.setPreferredSize(new Dimension(38, 28));
        folio.add(buscarFolio, BorderLayout.EAST);
        agregar(panel, 1, 1, "FOLIO / DOCUMENTO RELACIONADO", folio, 1);
        return panel;
    }

    private JPanel crearAsignacion() {
        JPanel panel = rejilla(4);
        agregar(panel, 0, 0, "PRIORIDAD *", prioridad, 1);
        agregar(panel, 1, 0, "ESTADO *", estado, 1);
        agregar(panel, 2, 0, "RESPONSABLE *", responsable, 1);
        agregar(panel, 3, 0, "CREADO POR", creadoPor, 1);
        agregar(panel, 0, 1, "FECHA DE INICIO *", inicio, 1);
        agregar(panel, 1, 1, "FECHA DE VENCIMIENTO *", vencimiento, 1);
        return panel;
    }

    private JPanel crearSubtareas() {
        JPanel panel = new JPanel(new BorderLayout(0, 5)); panel.setOpaque(false);
        JPanel fila = new JPanel(new BorderLayout(8, 0)); fila.setOpaque(false);
        JTextField subtarea = new JTextField();
        subtarea.setToolTipText("Escribe una subtarea");
        fila.add(subtarea, BorderLayout.CENTER);
        JButton agregar = boton("Agregar subtarea", false, "agregar");
        agregar.addActionListener(e -> {
            if (!subtarea.getText().isBlank()) {
                subtarea.setText("");
            }
        });
        fila.add(agregar, BorderLayout.EAST);
        panel.add(fila, BorderLayout.NORTH);
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
        JButton seleccionar = boton("Seleccionar archivos", false, "adjunto");
        seleccionar.addActionListener(e -> seleccionarAdjuntos());
        panel.add(drop, BorderLayout.CENTER);
        panel.add(seleccionar, BorderLayout.SOUTH);
        return panel;
    }

    private JPanel crearObservaciones() {
        JPanel panel = new JPanel(new BorderLayout()); panel.setOpaque(false);
        agregarArea(panel, "INFORMACIÓN ADICIONAL, COMENTARIOS O NOTAS...", observaciones);
        return panel;
    }

    private JPanel rejilla(int columnas) {
        JPanel panel = new JPanel(new GridBagLayout()); panel.setOpaque(false);
        panel.putClientProperty("columnas", columnas);
        return panel;
    }

    private JPanel seccion(String tipoIcono, String tituloSeccion, JComponent contenido) {
        JPanel panel = new JPanel(new BorderLayout(0, 6)); panel.setOpaque(false);
        panel.setBorder(BorderFactory.createLineBorder(new Color(225, 234, 244)));
        JPanel encabezado = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 4));
        encabezado.setBackground(AZUL_SECCION);
        encabezado.setBorder(BorderFactory.createEmptyBorder(1, 7, 1, 7));
        JLabel iconoLabel = new JLabel(new IconoTarea(tipoIcono, 16));
        iconoLabel.setForeground(AZUL); iconoLabel.setFont(new Font("Segoe UI Symbol", Font.BOLD, 16));
        JLabel texto = new JLabel(tituloSeccion); texto.setForeground(AZUL); texto.setFont(new Font("Segoe UI", Font.BOLD, 12));
        encabezado.add(iconoLabel); encabezado.add(texto);
        panel.add(encabezado, BorderLayout.NORTH);
        JPanel cuerpo = new JPanel(new BorderLayout());
        cuerpo.setOpaque(false);
        cuerpo.setBorder(BorderFactory.createEmptyBorder(8, 9, 8, 9));
        cuerpo.add(contenido, BorderLayout.CENTER);
        panel.add(cuerpo, BorderLayout.CENTER);
        return panel;
    }

    private void agregar(JPanel p, int x, int y, String label, Component input, int span) {
        int columnas = (int) p.getClientProperty("columnas");
        GridBagConstraints c = new GridBagConstraints();
        c.gridx = x; c.gridy = y; c.gridwidth = span; c.weightx = span;
        c.fill = GridBagConstraints.HORIZONTAL; c.anchor = GridBagConstraints.NORTH;
        c.insets = new Insets(2, x == 0 ? 0 : 8, 5, x + span == columnas ? 0 : 8);
        JPanel box = new JPanel(new BorderLayout(0, 3)); box.setOpaque(false);
        JPanel etiqueta = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
        etiqueta.setOpaque(false);
        String texto = label.endsWith(" *") ? label.substring(0, label.length() - 2) : label;
        JLabel l = new JLabel(texto.toUpperCase(java.util.Locale.ROOT));
        l.setFont(new Font("Segoe UI", Font.BOLD, 11));
        l.setForeground(TEXTO_ETIQUETA);
        etiqueta.add(l);
        if (label.endsWith(" *")) {
            JLabel requerido = new JLabel("*");
            requerido.setFont(new Font("Segoe UI", Font.BOLD, 11));
            requerido.setForeground(new Color(217, 43, 43));
            etiqueta.add(requerido);
        }
        box.add(etiqueta, BorderLayout.NORTH); box.add(input, BorderLayout.CENTER); p.add(box, c);
    }

    private void agregarArea(JPanel p, int x, int y, String label, JTextArea area, int span) {
        area.setLineWrap(true); area.setWrapStyleWord(true);
        JScrollPane scroll = new JScrollPane(area);
        scroll.setPreferredSize(new Dimension(0, 58));
        scroll.setBorder(BorderFactory.createLineBorder(BORDE_CAMPO));
        scroll.setViewportBorder(BorderFactory.createEmptyBorder());
        agregar(p, x, y, label, scroll, span);
    }

    private void agregarArea(JPanel p, String textoInicial, JTextArea area) {
        area.setLineWrap(true); area.setWrapStyleWord(true); area.setToolTipText(textoInicial);
        JScrollPane scroll = new JScrollPane(area);
        scroll.setPreferredSize(new Dimension(0, 58));
        scroll.setBorder(BorderFactory.createLineBorder(BORDE_CAMPO));
        scroll.setViewportBorder(BorderFactory.createEmptyBorder());
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

    private JButton boton(String texto, boolean primario, String tipoIcono) {
        JButton boton = boton(texto, primario);
        boton.setIcon(new IconoTarea(tipoIcono, 16));
        boton.setIconTextGap(6);
        return boton;
    }

    private void estilizarCampos() {
        for (JComponent campo : new JComponent[]{titulo, responsable, relacionado, tipo, estado, prioridad, inicio, vencimiento}) {
            estilizarCampo(campo, false);
        }
        estilizarCampo(creadoPor, true);
        titulo.setToolTipText("Escribe un título claro y conciso");
        descripcion.setToolTipText("Describe la tarea, los objetivos y cualquier información relevante");
        relacionado.setToolTipText("Busca un folio, orden o proveedor");
        estiloArea(descripcion); estiloArea(observaciones);
    }

    private void estilizarCampo(JComponent campo, boolean soloLectura) {
        campo.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        campo.setPreferredSize(new Dimension(0, 28));
        campo.setBackground(soloLectura ? FONDO_SOLO_LECTURA : Color.WHITE);
        if (campo instanceof JTextField campoTexto) {
            campoTexto.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(BORDE_CAMPO),
                    BorderFactory.createEmptyBorder(4, 10, 4, 10)));
            campoTexto.setEditable(!soloLectura || campoTexto.isEditable());
        } else if (campo instanceof JComboBox<?> combo) {
            javax.swing.border.Border bordeCombo = UIManager.getBorder("ComboBox.border");
            if (bordeCombo != null) {
                combo.setBorder(bordeCombo);
            }
            combo.setOpaque(true);
            combo.setBackground(Color.WHITE);
            combo.setRenderer(new DefaultListCellRenderer() {
                @Override
                public Component getListCellRendererComponent(JList<?> lista, Object valor, int indice,
                        boolean seleccionado, boolean foco) {
                    JLabel etiqueta = (JLabel) super.getListCellRendererComponent(
                            lista, valor, indice, seleccionado, foco);
                    etiqueta.setFont(new Font("Segoe UI", Font.PLAIN, 12));
                    etiqueta.setBorder(BorderFactory.createEmptyBorder(2, 8, 2, 8));
                    return etiqueta;
                }
            });
        } else if (campo instanceof JDateChooser selector) {
            selector.setBorder(BorderFactory.createLineBorder(BORDE_CAMPO));
            selector.setBackground(soloLectura ? FONDO_SOLO_LECTURA : Color.WHITE);
            Component editor = selector.getDateEditor().getUiComponent();
            editor.setFont(new Font("Segoe UI", Font.PLAIN, 14));
            editor.setBackground(soloLectura ? FONDO_SOLO_LECTURA : Color.WHITE);
            if (editor instanceof JTextField campoFecha) {
                campoFecha.setBorder(BorderFactory.createEmptyBorder(4, 10, 4, 10));
            }
        }
        if (campo instanceof javax.swing.text.JTextComponent campoTexto) {
            EstiloAplicacion.aplicarEstiloCamposTexto(campoTexto);
        } else if (campo instanceof JDateChooser selector) {
            EstiloAplicacion.aplicarEstiloCamposTexto(
                    selector.getDateEditor().getUiComponent());
        }
    }

    private void configurarCamposProtegidos() {
        boolean administrador = esAdministrador();
        String nombre = SesionActual.nombre();
        if (nombre.isBlank()) {
            nombre = SesionActual.usuario();
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
        return SesionActual.esAdministrador();
    }

    private void estiloArea(JTextArea area) {
        area.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        area.setLineWrap(true);
        area.setWrapStyleWord(true);
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
        selector.setLocale(java.util.Locale.forLanguageTag("es-MX"));
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

    private static final class IconoTarea implements javax.swing.Icon {
        private final String tipo;
        private final int tamano;

        private IconoTarea(String tipo, int tamano) {
            this.tipo = tipo;
            this.tamano = tamano;
        }

        @Override
        public int getIconWidth() {
            return tamano;
        }

        @Override
        public int getIconHeight() {
            return tamano;
        }

        @Override
        public void paintIcon(Component componente, Graphics g, int x, int y) {
            Graphics2D grafico = (Graphics2D) g.create();
            grafico.setColor(componente == null ? TEXTO : componente.getForeground());
            grafico.setStroke(new BasicStroke(Math.max(1.4f, tamano / 11f),
                    BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
            grafico.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
                    RenderingHints.VALUE_ANTIALIAS_ON);

            int s = tamano;
            switch (tipo) {
                case "editar":
                    grafico.drawRoundRect(x + s / 5, y + s / 5, s * 3 / 5, s * 3 / 5, 3, 3);
                    grafico.drawLine(x + s * 3 / 5, y + s / 5, x + s * 4 / 5, y + s * 2 / 5);
                    grafico.drawLine(x + s * 2 / 5, y + s * 3 / 5, x + s * 7 / 10, y + s / 3);
                    break;
                case "informacion":
                    grafico.drawRoundRect(x + s / 5, y + s / 5, s * 3 / 5, s * 3 / 5, 2, 2);
                    grafico.drawLine(x + s * 3 / 10, y + s * 2 / 5, x + s * 7 / 10, y + s * 2 / 5);
                    grafico.drawLine(x + s * 3 / 10, y + s * 3 / 5, x + s * 7 / 10, y + s * 3 / 5);
                    break;
                case "asignacion":
                    grafico.fillOval(x + s * 2 / 5, y + s / 5, s / 5, s / 5);
                    grafico.fillRoundRect(x + s / 4, y + s * 2 / 5, s / 2, s * 2 / 5, 4, 4);
                    break;
                case "subtareas":
                    for (int i = 0; i < 3; i++) {
                        grafico.fillRect(x + s / 5, y + s / 4 + i * s / 5, s / 10, s / 10);
                        grafico.drawLine(x + s * 2 / 5, y + s / 4 + i * s / 5 + s / 20,
                                x + s * 4 / 5, y + s / 4 + i * s / 5 + s / 20);
                    }
                    break;
                case "adjuntos":
                    grafico.drawOval(x + s * 3 / 10, y + s / 5, s * 2 / 5, s * 3 / 5);
                    grafico.drawOval(x + s * 2 / 5, y + s * 3 / 10, s / 5, s * 2 / 5);
                    break;
                case "observaciones":
                    grafico.drawRoundRect(x + s / 5, y + s / 5, s * 3 / 5, s * 3 / 5, 2, 2);
                    grafico.drawLine(x + s * 3 / 10, y + s * 2 / 5, x + s * 7 / 10, y + s * 2 / 5);
                    grafico.drawLine(x + s * 3 / 10, y + s * 3 / 5, x + s * 3 / 5, y + s * 3 / 5);
                    break;
                case "buscar":
                    grafico.drawOval(x + s / 5, y + s / 5, s * 2 / 5, s * 2 / 5);
                    grafico.drawLine(x + s * 1 / 2, y + s * 1 / 2, x + s * 4 / 5, y + s * 4 / 5);
                    break;
                case "agregar":
                    grafico.drawLine(x + s / 2, y + s / 5, x + s / 2, y + s * 4 / 5);
                    grafico.drawLine(x + s / 5, y + s / 2, x + s * 4 / 5, y + s / 2);
                    break;
                case "atras":
                    grafico.drawLine(x + s * 4 / 5, y + s / 2, x + s / 5, y + s / 2);
                    grafico.drawLine(x + s / 5, y + s / 2, x + s * 2 / 5, y + s / 4);
                    grafico.drawLine(x + s / 5, y + s / 2, x + s * 2 / 5, y + s * 3 / 4);
                    break;
                case "guardar":
                    grafico.drawRoundRect(x + s / 5, y + s / 5, s * 3 / 5, s * 3 / 5, 2, 2);
                    grafico.drawRect(x + s * 3 / 10, y + s / 5, s * 2 / 5, s / 4);
                    grafico.fillRect(x + s * 3 / 10, y + s * 3 / 5, s * 2 / 5, s / 5);
                    break;
                case "adjunto":
                    grafico.drawRect(x + s / 4, y + s / 5, s / 2, s * 3 / 5);
                    grafico.drawLine(x + s * 3 / 10, y + s * 2 / 5, x + s * 7 / 10, y + s * 2 / 5);
                    break;
                default:
                    grafico.drawOval(x + s / 4, y + s / 4, s / 2, s / 2);
                    break;
            }
            grafico.dispose();
        }
    }
}