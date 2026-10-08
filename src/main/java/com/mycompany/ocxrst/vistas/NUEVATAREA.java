package com.mycompany.ocxrst.vistas;

import com.mycompany.ocxrst.model.Tarea;
import com.mycompany.ocxrst.repository.TareaRepository;
import java.awt.*;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import javax.swing.*;

public class NUEVATAREA extends JDialog {
    private final TareaRepository repository;
    private final Tarea original;
    private boolean guardado;
    private final JTextField titulo = new JTextField();
    private final JTextArea descripcion = new JTextArea();
    private final JComboBox<String> tipo = new JComboBox<>(Tarea.TIPOS.toArray(String[]::new));
    private final JComboBox<String> estado = new JComboBox<>(Tarea.ESTADOS.toArray(String[]::new));
    private final JComboBox<String> prioridad = new JComboBox<>(Tarea.PRIORIDADES.toArray(String[]::new));
    private final JTextField responsable = new JTextField();
    private final JTextField inicio = new JTextField();
    private final JTextField vencimiento = new JTextField();
    private final JTextField relacionado = new JTextField();
    private final JTextArea observaciones = new JTextArea();
    private final java.util.List<String> adjuntos = new ArrayList<>();

    public NUEVATAREA(Window owner, Tarea original, TareaRepository repository) {
        super(owner, "NUEVA TAREA", ModalityType.APPLICATION_MODAL);
        this.original = original; this.repository = repository;
        setContentPane(crearContenido()); setSize(760, 580); setLocationRelativeTo(owner);
        if (original != null) cargarOriginal();
    }
    private JPanel crearContenido() {
        JPanel panel = new JPanel(new BorderLayout(0, 12)); panel.setBackground(Color.WHITE);
        panel.setBorder(BorderFactory.createEmptyBorder(20, 24, 18, 24));
        JLabel tituloVista = new JLabel(original == null ? "NUEVA TAREA" : "EDITAR TAREA");
        tituloVista.setFont(new Font("Segoe UI", Font.BOLD, 24)); tituloVista.setForeground(new Color(28,82,132));
        panel.add(tituloVista, BorderLayout.NORTH);
        JPanel campos = new JPanel(new GridBagLayout()); campos.setOpaque(false);
        GridBagConstraints c = new GridBagConstraints(); c.insets = new Insets(5, 5, 5, 5); c.fill = GridBagConstraints.HORIZONTAL; c.weightx = 1;
        agregar(campos,c,0,0,"* TÍTULO / DESCRIPCIÓN BREVE",titulo,2);
        agregar(campos,c,0,1,"* TIPO",tipo,1); agregar(campos,c,1,1,"* PRIORIDAD",prioridad,1); agregar(campos,c,2,1,"* ESTADO",estado,1);
        agregar(campos,c,0,2,"* RESPONSABLE",responsable,1); agregar(campos,c,1,2,"* FECHA DE INICIO (dd/MM/yyyy)",inicio,1); agregar(campos,c,2,2,"* FECHA DE VENCIMIENTO (dd/MM/yyyy)",vencimiento,1);
        agregar(campos,c,0,3,"FOLIO RELACIONADO",relacionado,3);
        agregarArea(campos,c,0,4,"DESCRIPCIÓN DETALLADA",descripcion,3); agregarArea(campos,c,0,5,"OBSERVACIONES",observaciones,3);
        JPanel botones = new JPanel(new FlowLayout(FlowLayout.RIGHT)); botones.setOpaque(false);
        JButton cancelar = new JButton("ATRÁS"); cancelar.addActionListener(e -> dispose());
        JButton guardar = new JButton("GUARDAR"); guardar.setBackground(new Color(31,112,184)); guardar.setForeground(Color.WHITE);
        guardar.addActionListener(e -> guardar()); botones.add(cancelar); botones.add(guardar);
        panel.add(new JScrollPane(campos), BorderLayout.CENTER); panel.add(botones, BorderLayout.SOUTH); return panel;
    }
    private void agregar(JPanel p, GridBagConstraints base, int x, int y, String label, Component input, int span) {
        GridBagConstraints c = (GridBagConstraints) base.clone(); c.gridx=x; c.gridy=y; c.gridwidth=span; c.weightx=span; JPanel box = new JPanel(new BorderLayout(0,3)); box.setOpaque(false);
        JLabel l = new JLabel(label); l.setFont(new Font("Segoe UI", Font.BOLD, 10)); l.setForeground(new Color(28,82,132)); box.add(l,BorderLayout.NORTH); box.add(input,BorderLayout.CENTER); p.add(box,c);
    }
    private void agregarArea(JPanel p, GridBagConstraints base, int x, int y, String label, JTextArea area, int span) {
        area.setLineWrap(true); area.setWrapStyleWord(true); agregar(p,base,x,y,label,new JScrollPane(area),span);
    }
    private void cargarOriginal() {
        titulo.setText(original.titulo()); descripcion.setText(original.descripcion()); tipo.setSelectedItem(original.tipo()); estado.setSelectedItem(original.estado()); prioridad.setSelectedItem(original.prioridad()); responsable.setText(original.responsable()); inicio.setText(original.inicio().toString()); vencimiento.setText(original.vencimiento().toString()); relacionado.setText(original.folioRelacionado()); observaciones.setText(original.observaciones()); adjuntos.addAll(original.adjuntos());
    }
    private void guardar() {
        try {
            LocalDate desde = fecha(inicio.getText()), hasta = fecha(vencimiento.getText());
            Tarea tarea = new Tarea(original == null ? "" : original.folio(), titulo.getText(), (String)tipo.getSelectedItem(), (String)estado.getSelectedItem(), (String)prioridad.getSelectedItem(), responsable.getText(), original == null ? LocalDate.now() : original.creacion(), hasta, descripcion.getText(), desde, relacionado.getText(), observaciones.getText(), adjuntos);
            repository.guardar(tarea); guardado = true; dispose();
        } catch (IllegalArgumentException | IOException ex) { JOptionPane.showMessageDialog(this, ex.getMessage(), "No se pudo guardar", JOptionPane.WARNING_MESSAGE); }
    }
    private LocalDate fecha(String valor) {
        String limpio = valor.trim(); if (limpio.matches("\\d{2}/\\d{2}/\\d{4}")) { String[] p=limpio.split("/"); return LocalDate.of(Integer.parseInt(p[2]),Integer.parseInt(p[1]),Integer.parseInt(p[0])); } return LocalDate.parse(limpio);
    }
    public boolean guardado() { return guardado; }
}
