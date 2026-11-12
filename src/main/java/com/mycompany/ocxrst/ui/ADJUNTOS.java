package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.util.EstiloAplicacion;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Font;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;
import javax.swing.BorderFactory;
import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.ListSelectionModel;

public class ADJUNTOS extends JFrame {
    private static final Color AZUL = new Color(28, 82, 132);
    private static final Color AZUL_ACCION = new Color(31, 112, 184);
    private static final Path CARPETA_ADJUNTOS = Paths.get(
            "C:\\OCXRST\\OrdendeComprasRST\\MESATRABAJO\\Adjuntos");
    private final DefaultListModel<String> modelo = new DefaultListModel<>();
    private final JList<String> lista = new JList<>(modelo);
    private final JLabel estado = new JLabel();

    public ADJUNTOS() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        setTitle("ADJUNTOS");
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        setContentPane(crearContenido());
        EstiloAplicacion.aplicarEstiloBotonesRestoPantallas(getContentPane());
        setSize(760, 520);
        setMinimumSize(new Dimension(620, 420));
        setLocationRelativeTo(null);
        cargarCarpetas();
    }

    private JPanel crearContenido() {
        JPanel fondo = new JPanel(new BorderLayout(0, 10));
        fondo.setBackground(EstiloAplicacion.FONDO_PANTALLA);
        fondo.setBorder(BorderFactory.createEmptyBorder(18, 20, 16, 20));

        JPanel encabezado = new JPanel(new BorderLayout());
        encabezado.setOpaque(false);
        JLabel titulo = new JLabel("<html><b>Adjuntos</b><br>"
                + "<font size='3'>Carpetas y archivos asociados a las solicitudes.</font></html>");
        titulo.setForeground(AZUL);
        titulo.setFont(new Font("Segoe UI", Font.BOLD, 20));
        JButton actualizar = new JButton("Actualizar");
        actualizar.setForeground(AZUL);
        actualizar.addActionListener(e -> cargarCarpetas());
        encabezado.add(titulo, BorderLayout.CENTER);
        encabezado.add(actualizar, BorderLayout.EAST);
        fondo.add(encabezado, BorderLayout.NORTH);

        lista.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        lista.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        lista.setForeground(new Color(24, 58, 104));
        lista.setBackground(Color.WHITE);
        lista.setBorder(BorderFactory.createEmptyBorder(8, 10, 8, 10));
        lista.addListSelectionListener(e -> actualizarBotonAbrir());
        fondo.add(new JScrollPane(lista), BorderLayout.CENTER);

        JPanel pie = new JPanel(new BorderLayout(8, 0));
        pie.setOpaque(false);
        estado.setForeground(AZUL);
        JButton abrir = new JButton("Abrir carpeta");
        abrir.setForeground(AZUL_ACCION);
        abrir.setEnabled(false);
        abrir.addActionListener(e -> abrirCarpetaSeleccionada());
        lista.putClientProperty("botonAbrir", abrir);
        pie.add(estado, BorderLayout.CENTER);
        pie.add(abrir, BorderLayout.EAST);
        fondo.add(pie, BorderLayout.SOUTH);
        return fondo;
    }

    private void cargarCarpetas() {
        modelo.clear();
        try {
            Files.createDirectories(CARPETA_ADJUNTOS);
            List<Path> carpetas = Files.list(CARPETA_ADJUNTOS)
                    .filter(Files::isDirectory)
                    .sorted(Comparator.comparing(path -> path.getFileName().toString()))
                    .collect(Collectors.toList());
            for (Path carpeta : carpetas) {
                long archivos = Files.list(carpeta).filter(Files::isRegularFile).count();
                modelo.addElement(carpeta.getFileName() + "  (" + archivos + " archivo(s))");
            }
            estado.setText(carpetas.size() + " carpeta(s) en " + CARPETA_ADJUNTOS);
        } catch (IOException ex) {
            estado.setText("No se pudieron leer los adjuntos.");
            javax.swing.JOptionPane.showMessageDialog(this,
                    "No se pudo cargar la carpeta de adjuntos:\n" + ex.getMessage(),
                    "Error", javax.swing.JOptionPane.ERROR_MESSAGE);
        }
        actualizarBotonAbrir();
    }

    private void actualizarBotonAbrir() {
        JButton abrir = (JButton) lista.getClientProperty("botonAbrir");
        if (abrir != null) {
            abrir.setEnabled(lista.getSelectedIndex() >= 0);
        }
    }

    private void abrirCarpetaSeleccionada() {
        int indice = lista.getSelectedIndex();
        if (indice < 0) {
            return;
        }
        String texto = modelo.getElementAt(indice);
        String nombre = texto.substring(0, texto.indexOf("  ("));
        try {
            if (!Desktop.isDesktopSupported()) {
                throw new IOException("El sistema no permite abrir carpetas.");
            }
            Desktop.getDesktop().open(CARPETA_ADJUNTOS.resolve(nombre).toFile());
        } catch (IOException ex) {
            javax.swing.JOptionPane.showMessageDialog(this,
                    "No se pudo abrir la carpeta:\n" + ex.getMessage(),
                    "Error", javax.swing.JOptionPane.ERROR_MESSAGE);
        }
    }
}
