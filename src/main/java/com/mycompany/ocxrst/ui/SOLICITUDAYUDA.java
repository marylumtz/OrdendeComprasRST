package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.model.Tarea;
import com.mycompany.ocxrst.repository.TareaRepository;
import com.mycompany.ocxrst.service.SesionActual;
import com.mycompany.ocxrst.util.EstiloAplicacion;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Insets;
import java.io.IOException;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.TransferHandler;

public class SOLICITUDAYUDA extends JFrame {
    private static final long MAX_ADJUNTO_BYTES = 2_500_000L;
    private static final Color AZUL = new Color(7, 57, 127);
    private static final Color AZUL_ACCION = new Color(10, 116, 211);
    private static final Color AZUL_SECCION = new Color(232, 243, 253);
    private static final Color BORDE = new Color(190, 208, 229);
    private static final Color TEXTO = new Color(24, 58, 104);
    private static final TareaRepository REPOSITORY = new TareaRepository();

    private final JComboBox<String> tipoSolicitud = new JComboBox<>(
            new String[]{"Reporte de problema", "Solicitud de apoyo", "Requerimiento"});
    private final JComboBox<String> categoria = new JComboBox<>(
            new String[]{"Sistemas / ERP", "Compras", "Inventarios", "Proveedores", "General"});
    private final JTextField asunto = new JTextField();
    private final JTextArea descripcion = new JTextArea();
    private final JTextField solicitante = new JTextField();
    private final JComboBox<String> area = new JComboBox<>(
            new String[]{"Compras", "Sistemas", "Inventarios", "Proveedores", "Administración"});
    private final JTextField telefono = new JTextField();
    private final JLabel archivosTexto = new JLabel("Arrastra archivos aquí o haz clic para seleccionar",
            SwingConstants.CENTER);
    private final List<String> adjuntos = new ArrayList<>();
    private final JPanel archivosLista = new JPanel();

    public SOLICITUDAYUDA() {
        EstiloAplicacion.aplicarTipografiaGlobal();
        setTitle("SOLICITUD DE AYUDA");
        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);
        setContentPane(crearContenido());
        setSize(980, 720);
        setMinimumSize(new Dimension(850, 620));
        setLocationRelativeTo(null);
        EstiloAplicacion.aplicarTipografia(getContentPane());
        EstiloAplicacion.aplicarEstiloCamposTexto(getContentPane());
        cargarUsuario();
    }

    private JPanel crearContenido() {
        JPanel fondo = new JPanel(new BorderLayout(0, 10));
        fondo.setBackground(EstiloAplicacion.FONDO_PANTALLA);
        fondo.setBorder(BorderFactory.createEmptyBorder(0, 0, 14, 0));

        JPanel cuerpo = new JPanel(new BorderLayout(0, 10));
        cuerpo.setOpaque(false);
        cuerpo.setBorder(BorderFactory.createEmptyBorder(0, 14, 0, 14));
        cuerpo.add(crearAviso(), BorderLayout.NORTH);

        JPanel formularios = new JPanel(new BorderLayout(0, 10));
        formularios.setOpaque(false);
        formularios.add(seccion("1. Información de la solicitud", crearInformacion()), BorderLayout.NORTH);
        JPanel inferior = new JPanel(new GridLayout(1, 2, 10, 0));
        inferior.setOpaque(false);
        inferior.add(seccion("2. Información del solicitante", crearSolicitante()));
        inferior.add(seccion("3. Archivos adjuntos (opcional)", crearAdjuntos()));
        formularios.add(inferior, BorderLayout.CENTER);
        cuerpo.add(formularios, BorderLayout.CENTER);
        fondo.add(cuerpo, BorderLayout.CENTER);
        fondo.add(crearBotones(), BorderLayout.SOUTH);
        return fondo;
    }

    private JPanel crearAviso() {
        JPanel aviso = new JPanel(new BorderLayout());
        aviso.setBackground(new Color(225, 239, 255));
        aviso.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(BORDE),
                BorderFactory.createEmptyBorder(10, 16, 10, 16)));
        JLabel texto = new JLabel("<html><b>Tu solicitud será registrada y atendida por el área correspondiente.</b><br>"
                + "Podrás dar seguimiento desde la opción Mis solicitudes.</html>");
        texto.setForeground(AZUL);
        texto.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        aviso.add(texto, BorderLayout.CENTER);
        return aviso;
    }

    private JPanel crearInformacion() {
        JPanel panel = formulario(2);
        agregar(panel, 0, 0, "Tipo de solicitud *", tipoSolicitud, 1);
        agregar(panel, 1, 0, "Categoría *", categoria, 1);
        agregar(panel, 0, 1, "Asunto (título) *", asunto, 2);
        JScrollPane descripcionScroll = new JScrollPane(descripcion);
        descripcionScroll.setPreferredSize(new Dimension(0, 76));
        agregar(panel, 0, 2, "Descripción del requerimiento o problema *", descripcionScroll, 2);
        descripcion.setLineWrap(true);
        descripcion.setWrapStyleWord(true);
        descripcion.setBorder(BorderFactory.createEmptyBorder(7, 9, 7, 9));
        return panel;
    }

    private JPanel crearSolicitante() {
        JPanel panel = formulario(2);
        solicitante.setEditable(false);
        solicitante.setBackground(new Color(247, 250, 253));
        agregar(panel, 0, 0, "Solicitante", solicitante, 2);
        agregar(panel, 0, 1, "Área / Departamento *", area, 1);
        agregar(panel, 1, 1, "Teléfono / Extensión", telefono, 1);
        return panel;
    }

    private JPanel crearAdjuntos() {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setOpaque(false);
        JPanel zonaArchivos = new JPanel(new BorderLayout(0, 6));
        zonaArchivos.setOpaque(false);
        archivosTexto.setForeground(AZUL);
        archivosTexto.setBorder(BorderFactory.createEmptyBorder(8, 8, 4, 8));
        TransferHandler arrastre = crearControladorArrastre();
        zonaArchivos.setBorder(BorderFactory.createDashedBorder(BORDE));
        zonaArchivos.setTransferHandler(arrastre);
        archivosTexto.setTransferHandler(arrastre);
        archivosLista.setLayout(new javax.swing.BoxLayout(archivosLista, javax.swing.BoxLayout.Y_AXIS));
        archivosLista.setTransferHandler(arrastre);
        archivosLista.setForeground(AZUL);
        archivosLista.setBackground(new Color(247, 250, 253));
        archivosLista.setBorder(BorderFactory.createEmptyBorder(4, 8, 4, 8));
        JScrollPane listaScroll = new JScrollPane(archivosLista);
        listaScroll.setBorder(BorderFactory.createEmptyBorder());
        listaScroll.setPreferredSize(new Dimension(0, 105));
        zonaArchivos.add(archivosTexto, BorderLayout.NORTH);
        zonaArchivos.add(listaScroll, BorderLayout.CENTER);
        JButton seleccionar = new JButton("Seleccionar archivos");
        seleccionar.setForeground(AZUL_ACCION);
        seleccionar.addActionListener(e -> seleccionarAdjuntos());
        panel.add(zonaArchivos, BorderLayout.CENTER);
        panel.add(seleccionar, BorderLayout.SOUTH);
        return panel;
    }

    private TransferHandler crearControladorArrastre() {
        return new TransferHandler() {
            @Override
            public boolean canImport(TransferSupport soporte) {
                return soporte.isDataFlavorSupported(DataFlavor.javaFileListFlavor);
            }

            @Override
            public boolean importData(TransferSupport soporte) {
                if (!canImport(soporte)) {
                    return false;
                }
                try {
                    Transferable datos = soporte.getTransferable();
                    @SuppressWarnings("unchecked")
                    List<File> archivos = (List<File>) datos.getTransferData(DataFlavor.javaFileListFlavor);
                    return agregarAdjuntos(archivos) > 0;
                } catch (Exception ex) {
                    javax.swing.JOptionPane.showMessageDialog(
                            SOLICITUDAYUDA.this,
                            "No se pudieron cargar los archivos arrastrados:\n" + ex.getMessage(),
                            "Error al adjuntar", javax.swing.JOptionPane.ERROR_MESSAGE);
                    return false;
                }
            }
        };
    }

    private int agregarAdjuntos(List<File> archivos) {
        List<String> formatosNoPermitidos = new ArrayList<>();
        List<String> archivosDemasiadoGrandes = new ArrayList<>();
        int agregados = 0;
        for (File archivo : archivos) {
            if (!archivo.isFile() || adjuntos.contains(archivo.getAbsolutePath())) {
                continue;
            }
            if (!esImagenPermitida(archivo)) {
                formatosNoPermitidos.add(archivo.getName());
                continue;
            }
            if (archivo.length() > MAX_ADJUNTO_BYTES) {
                archivosDemasiadoGrandes.add(archivo.getName());
                continue;
            }
            adjuntos.add(archivo.getAbsolutePath());
            agregados++;
        }
        archivosLista.removeAll();
        for (String ruta : adjuntos) {
            archivosLista.add(crearFilaAdjunto(ruta));
        }
        archivosTexto.setText(adjuntos.isEmpty()
                ? "Arrastra archivos aquí o haz clic para seleccionar"
                : "Archivos que se copiarán en la carpeta de la solicitud:");
        archivosLista.revalidate();
        archivosLista.repaint();
        if (!formatosNoPermitidos.isEmpty() || !archivosDemasiadoGrandes.isEmpty()) {
            StringBuilder mensaje = new StringBuilder();
            if (!formatosNoPermitidos.isEmpty()) {
                mensaje.append("Estos archivos no son JPG o PNG y no fueron adjuntados:\n")
                        .append(String.join("\n", formatosNoPermitidos)).append("\n\n");
            }
            if (!archivosDemasiadoGrandes.isEmpty()) {
                mensaje.append("Estos archivos superan el límite de 2.5 MB y no fueron adjuntados:\n")
                        .append(String.join("\n", archivosDemasiadoGrandes));
            }
            javax.swing.JOptionPane.showMessageDialog(this,
                    mensaje.toString(), "Archivo no permitido",
                    javax.swing.JOptionPane.WARNING_MESSAGE);
        }
        return agregados;
    }

    private JPanel crearFilaAdjunto(String ruta) {
        JPanel fila = new JPanel(new BorderLayout(8, 0));
        fila.setOpaque(false);
        fila.setBorder(BorderFactory.createEmptyBorder(2, 0, 2, 0));
        JLabel nombre = new JLabel(new File(ruta).getName());
        nombre.setForeground(TEXTO);
        nombre.setFont(new Font("Segoe UI", Font.PLAIN, 12));

        JPanel acciones = new JPanel(new FlowLayout(FlowLayout.RIGHT, 4, 0));
        acciones.setOpaque(false);
        JButton ver = new JButton("Ver");
        ver.addActionListener(e -> verAdjunto(ruta));
        JButton quitar = new JButton("Quitar");
        quitar.addActionListener(e -> {
            adjuntos.remove(ruta);
            actualizarListadoAdjuntos();
        });
        acciones.add(ver);
        acciones.add(quitar);
        fila.add(nombre, BorderLayout.CENTER);
        fila.add(acciones, BorderLayout.EAST);
        return fila;
    }

    private void actualizarListadoAdjuntos() {
        archivosLista.removeAll();
        for (String ruta : adjuntos) {
            archivosLista.add(crearFilaAdjunto(ruta));
        }
        archivosTexto.setText(adjuntos.isEmpty()
                ? "Arrastra archivos aquí o haz clic para seleccionar"
                : "Archivos que se copiarán en la carpeta de la solicitud:");
        archivosLista.revalidate();
        archivosLista.repaint();
    }

    private void verAdjunto(String ruta) {
        try {
            if (!Desktop.isDesktopSupported()) {
                throw new IOException("El sistema no permite abrir imágenes desde la aplicación.");
            }
            Desktop.getDesktop().open(new File(ruta));
        } catch (IOException ex) {
            javax.swing.JOptionPane.showMessageDialog(this,
                    "No se pudo abrir la imagen:\n" + ex.getMessage(),
                    "No se pudo ver el archivo", javax.swing.JOptionPane.ERROR_MESSAGE);
        }
    }

    private boolean esImagenPermitida(File archivo) {
        String nombre = archivo.getName().toLowerCase(java.util.Locale.ROOT);
        return nombre.endsWith(".jpg") || nombre.endsWith(".jpeg") || nombre.endsWith(".png");
    }

    private JPanel crearBotones() {
        JPanel botones = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
        botones.setOpaque(false);
        JButton cancelar = new JButton("Cancelar");
        cancelar.addActionListener(e -> dispose());
        JButton enviar = new JButton("Enviar solicitud");
        enviar.setBackground(AZUL_ACCION);
        enviar.setForeground(Color.WHITE);
        enviar.addActionListener(e -> enviarSolicitud());
        botones.add(cancelar);
        botones.add(enviar);
        return botones;
    }

    private JPanel seccion(String titulo, JPanel contenido) {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setOpaque(false);
        JLabel encabezado = new JLabel(titulo);
        encabezado.setOpaque(true);
        encabezado.setBackground(AZUL_SECCION);
        encabezado.setForeground(AZUL);
        encabezado.setFont(new Font("Segoe UI", Font.BOLD, 14));
        encabezado.setBorder(BorderFactory.createEmptyBorder(7, 12, 7, 12));
        panel.add(encabezado, BorderLayout.NORTH);
        panel.add(contenido, BorderLayout.CENTER);
        panel.setBorder(BorderFactory.createLineBorder(BORDE));
        return panel;
    }

    private JPanel formulario(int columnas) {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setOpaque(false);
        panel.setBorder(BorderFactory.createEmptyBorder(8, 12, 10, 12));
        return panel;
    }

    private void agregar(JPanel panel, int columna, int fila, String texto,
            java.awt.Component componente, int ancho) {
        JPanel campo = new JPanel(new BorderLayout(0, 4));
        campo.setOpaque(false);
        JLabel etiqueta = new JLabel(texto);
        etiqueta.setForeground(TEXTO);
        etiqueta.setFont(new Font("Segoe UI", Font.BOLD, 11));
        campo.add(etiqueta, BorderLayout.NORTH);
        campo.add(componente, BorderLayout.CENTER);
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = columna;
        gbc.gridy = fila;
        gbc.gridwidth = ancho;
        gbc.weightx = ancho;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(3, 5, 3, 5);
        panel.add(campo, gbc);
    }

    private void cargarUsuario() {
        String nombre = SesionActual.nombre();
        solicitante.setText(nombre.isBlank() ? SesionActual.usuario() : nombre);
    }

    private void seleccionarAdjuntos() {
        JFileChooser selector = new JFileChooser();
        selector.setMultiSelectionEnabled(true);
        selector.setFileFilter(new FileNameExtensionFilter(
                "Imágenes JPG o PNG", "jpg", "jpeg", "png"));
        if (selector.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        agregarAdjuntos(List.of(selector.getSelectedFiles()));
    }

    private void enviarSolicitud() {
        String titulo = asunto.getText().trim();
        String detalle = descripcion.getText().trim();
        if (titulo.isBlank() || detalle.isBlank()) {
            javax.swing.JOptionPane.showMessageDialog(this,
                    "Completa el asunto y la descripción de la solicitud.",
                    "Datos incompletos", javax.swing.JOptionPane.WARNING_MESSAGE);
            return;
        }
        LocalDate hoy = LocalDate.now();
        Tarea tarea = new Tarea("", titulo, tipoTarea(), "Pendiente", "Media",
                "Pendiente de asignar", hoy, hoy.plusDays(7), detalle, hoy, "",
                "Solicitud de ayuda - Área: " + area.getSelectedItem()
                + (telefono.getText().isBlank() ? "" : " - Teléfono: " + telefono.getText().trim()),
                adjuntos);
        try {
            Tarea guardada = REPOSITORY.guardar(tarea);
            copiarFotosAMesaTrabajo(guardada);
            javax.swing.JOptionPane.showMessageDialog(this,
                    "Solicitud registrada correctamente con folio " + guardada.folio() + ".",
                    "Solicitud enviada", javax.swing.JOptionPane.INFORMATION_MESSAGE);
            dispose();
        } catch (IOException | IllegalArgumentException ex) {
            javax.swing.JOptionPane.showMessageDialog(this,
                    "No se pudo registrar la solicitud:\n" + ex.getMessage(),
                    "Error al guardar", javax.swing.JOptionPane.ERROR_MESSAGE);
        }
    }

    private void copiarFotosAMesaTrabajo(Tarea tarea) throws IOException {
        Path carpetaDestino = Paths.get(
                "C:\\OCXRST\\OrdendeComprasRST\\MESATRABAJO\\Adjuntos",
                tarea.folio());
        Files.createDirectories(carpetaDestino);
        for (String ruta : tarea.adjuntos()) {
            Path origen = Paths.get(ruta);
            if (!Files.isRegularFile(origen)) {
                throw new IOException("No se encontró la imagen adjunta: " + origen.getFileName());
            }
            String nombre = origen.getFileName().toString();
            Path destino = carpetaDestino.resolve(nombre);
            if (Files.exists(destino)) {
                destino = carpetaDestino.resolve(System.currentTimeMillis() + "_" + nombre);
            }
            Files.copy(origen, destino);
        }
    }

    private String tipoTarea() {
        return switch (String.valueOf(categoria.getSelectedItem())) {
            case "Compras" -> "Compra";
            case "Inventarios" -> "Inventario";
            case "Proveedores" -> "Proveedor";
            case "Sistemas / ERP" -> "Soporte";
            default -> "General";
        };
    }
}
