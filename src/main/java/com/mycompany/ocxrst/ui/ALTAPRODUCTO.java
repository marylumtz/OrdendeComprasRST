package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.repository.InventarioRepository;
import com.mycompany.ocxrst.util.EstiloAplicacion;
import com.mycompany.ocxrst.validation.InventarioValidator;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Insets;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import javax.imageio.ImageIO;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ALTAPRODUCTO extends JPanel {
    private static final Color AZUL = new Color(11, 92, 181);
    private static final Color AZUL_OSCURO = new Color(7, 57, 127);
    private static final Color FONDO = new Color(245, 249, 255);
    private static final Color BORDE = new Color(172, 202, 244);
    private final InventarioRepository repository = new InventarioRepository();
    private final JTextField codigo = new JTextField();
    private final JTextField sku = new JTextField();
    private final JTextField descripcion = new JTextField();
    private final JComboBox<String> categoria = new JComboBox<>(new String[]{"General", "Electrónica",
        "Ferretería", "Oficina", "Limpieza", "Consumibles", "Herramientas", "Refacciones"});
    private final JTextField marca = new JTextField();
    private final JComboBox<String> unidad = new JComboBox<>(new String[]{"PZA", "KG", "LT", "MT",
        "CAJA", "PAQUETE", "SERVICIO"});
    private final JTextField precio = new JTextField();
    private final JTextField cantidad = new JTextField();
    private final JTextField stockMinimo = new JTextField();
    private final JTextArea observaciones = new JTextArea();
    private final JLabel imagen = new JLabel("Agregar imagen", SwingConstants.CENTER);
    private final Runnable volver;
    private final String codigoInicial;
    private String rutaImagen = "";

    public ALTAPRODUCTO(Runnable volver) {
        this(volver, null);
    }

    public ALTAPRODUCTO(Runnable volver, String codigoExistente) {
        this.volver = volver;
        this.codigoInicial = codigoExistente;
        EstiloAplicacion.aplicarTipografiaGlobal();
        setLayout(new BorderLayout());
        add(crearContenido(), BorderLayout.CENTER);
        EstiloAplicacion.aplicarTipografia(this);
        EstiloAplicacion.aplicarEstiloBotones(this);
        EstiloAplicacion.aplicarEstiloCamposTexto(this);
        if (codigoInicial != null && !codigoInicial.isBlank()) {
            codigo.setEditable(false);
            cargarProductoExistente(codigoInicial);
        }
    }

    private JPanel crearContenido() {
        JPanel fondo = new JPanel(new BorderLayout(0, 10));
        fondo.setBackground(FONDO);
        fondo.setBorder(BorderFactory.createEmptyBorder(12, 12, 12, 12));
        fondo.add(crearEncabezado(), BorderLayout.NORTH);

        JPanel cuerpo = new JPanel(new GridLayout(1, 2, 10, 0));
        cuerpo.setOpaque(false);
        cuerpo.add(crearColumnaIzquierda());
        cuerpo.add(crearColumnaDerecha());
        fondo.add(cuerpo, BorderLayout.CENTER);

        JPanel botones = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
        botones.setOpaque(false);
        JButton cancelar = new JButton("Cancelar");
        JButton limpiar = new JButton("Limpiar");
        JButton guardar = new JButton("Guardar");
        cancelar.addActionListener(e -> volver.run());
        limpiar.addActionListener(e -> limpiar());
        guardar.addActionListener(e -> guardar());
        botones.add(cancelar);
        botones.add(limpiar);
        botones.add(guardar);
        fondo.add(botones, BorderLayout.SOUTH);
        return fondo;
    }

    private JPanel crearEncabezado() {
        JPanel encabezado = new JPanel(new BorderLayout(12, 0));
        encabezado.setOpaque(false);
        JLabel icono = new JLabel("◇");
        icono.setForeground(Color.WHITE);
        icono.setFont(new Font("Segoe UI Symbol", Font.BOLD, 30));
        JPanel insignia = new JPanel(new GridBagLayout());
        insignia.setBackground(AZUL);
        insignia.setPreferredSize(new Dimension(54, 46));
        insignia.add(icono);
        encabezado.add(insignia, BorderLayout.WEST);
        String textoEncabezado = codigoInicial == null || codigoInicial.isBlank()
                ? "Registro de Producto" : "Editar Producto";
        String subtitulo = codigoInicial == null || codigoInicial.isBlank()
                ? "Capture la información del producto." : "Actualice la información del producto.";
        JLabel titulo = new JLabel("<html><b>" + textoEncabezado + "</b><br>"
                + "<font size='3'>" + subtitulo + "</font></html>");
        titulo.setForeground(AZUL_OSCURO);
        titulo.setFont(new Font("Segoe UI", Font.PLAIN, 18));
        encabezado.add(titulo, BorderLayout.CENTER);
        return encabezado;
    }

    private JPanel crearColumnaIzquierda() {
        JPanel columna = new JPanel(new BorderLayout(0, 10));
        columna.setOpaque(false);
        columna.add(crearSeccion("Imagen del Producto", crearPanelImagen()), BorderLayout.NORTH);
        columna.add(crearSeccion("Control de Inventario", crearPanelInventario()), BorderLayout.CENTER);
        return columna;
    }

    private JPanel crearColumnaDerecha() {
        JPanel columna = new JPanel(new BorderLayout(0, 10));
        columna.setOpaque(false);
        JPanel general = new JPanel(new GridBagLayout());
        general.setOpaque(false);
        GridBagConstraints gc = restricciones();
        agregar(general, gc, 0, 0, "Código *", codigo);
        agregar(general, gc, 1, 0, "SKU", sku);
        agregar(general, gc, 0, 1, "Descripción *", descripcion, 2);
        agregar(general, gc, 0, 2, "Categoría *", categoria);
        agregar(general, gc, 1, 2, "Marca *", marca);
        agregar(general, gc, 0, 3, "Unidad *", unidad);
        agregar(general, gc, 1, 3, "Precio unitario *", precio);
        columna.add(crearSeccion("Información General", general), BorderLayout.NORTH);
        JPanel adicional = new JPanel(new BorderLayout(0, 8));
        adicional.setOpaque(false);
        adicional.add(new JLabel("Observaciones"), BorderLayout.NORTH);
        observaciones.setLineWrap(true);
        observaciones.setWrapStyleWord(true);
        adicional.add(new JScrollPane(observaciones), BorderLayout.CENTER);
        columna.add(crearSeccion("Información Adicional", adicional), BorderLayout.CENTER);
        return columna;
    }

    private JPanel crearPanelImagen() {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setOpaque(false);
        imagen.setOpaque(true);
        imagen.setBackground(Color.WHITE);
        imagen.setPreferredSize(new Dimension(240, 180));
        imagen.setBorder(BorderFactory.createLineBorder(BORDE));
        JButton cambiar = new JButton("Cambiar imagen");
        cambiar.addActionListener(e -> seleccionarImagen());
        panel.add(imagen, BorderLayout.CENTER);
        panel.add(cambiar, BorderLayout.SOUTH);
        return panel;
    }

    private JPanel crearPanelInventario() {
        JPanel panel = new JPanel(new GridLayout(2, 2, 8, 8));
        panel.setOpaque(false);
        panel.add(campo("Cantidad en inventario", cantidad));
        panel.add(campo("Stock mínimo", stockMinimo));
        panel.add(new JLabel("Estado de stock: Disponible"));
        return panel;
    }

    private JPanel campo(String texto, JTextField campo) {
        JPanel panel = new JPanel(new BorderLayout(0, 4));
        panel.setOpaque(false);
        panel.add(new JLabel(texto), BorderLayout.NORTH);
        panel.add(campo, BorderLayout.CENTER);
        return panel;
    }

    private JPanel crearSeccion(String titulo, JPanel contenido) {
        JPanel panel = new JPanel(new BorderLayout(0, 8));
        panel.setBackground(Color.WHITE);
        panel.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(BORDE),
                BorderFactory.createEmptyBorder(10, 12, 12, 12)));
        JLabel encabezado = new JLabel(titulo);
        encabezado.setForeground(AZUL_OSCURO);
        encabezado.setFont(new Font("Segoe UI", Font.BOLD, 15));
        panel.add(encabezado, BorderLayout.NORTH);
        panel.add(contenido, BorderLayout.CENTER);
        return panel;
    }

    private GridBagConstraints restricciones() {
        GridBagConstraints gc = new GridBagConstraints();
        gc.fill = GridBagConstraints.HORIZONTAL;
        gc.insets = new Insets(5, 5, 5, 5);
        gc.weightx = 1;
        return gc;
    }

    private void agregar(JPanel panel, GridBagConstraints base, int columna, int fila,
            String etiqueta, java.awt.Component componente) {
        agregar(panel, base, columna, fila, etiqueta, componente, 1);
    }

    private void agregar(JPanel panel, GridBagConstraints base, int columna, int fila,
            String etiqueta, java.awt.Component componente, int ancho) {
        GridBagConstraints gc = (GridBagConstraints) base.clone();
        gc.gridx = columna;
        gc.gridy = fila;
        gc.gridwidth = ancho;
        JPanel campo = new JPanel(new BorderLayout(0, 4));
        campo.setOpaque(false);
        campo.add(new JLabel(etiqueta), BorderLayout.NORTH);
        campo.add(componente, BorderLayout.CENTER);
        panel.add(campo, gc);
    }

    private void seleccionarImagen() {
        JFileChooser selector = new JFileChooser();
        if (selector.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        File archivo = selector.getSelectedFile();
        try {
            BufferedImage original = ImageIO.read(archivo);
            if (original == null) {
                throw new IOException("Formato no compatible.");
            }
            mostrarImagen(archivo, original);
        } catch (IOException ex) {
            javax.swing.JOptionPane.showMessageDialog(this, "No se pudo cargar la imagen.");
        }
    }

    private void cargarProductoExistente(String codigoExistente) {
        try (Workbook libro = repository.cargarLibro(repository.obtenerArchivoInventarios())) {
            Sheet hoja = repository.obtenerHojaInventarios(libro);
            Row fila = null;
            for (int indice = 1; indice <= hoja.getLastRowNum(); indice++) {
                Row candidata = hoja.getRow(indice);
                if (codigoExistente.equalsIgnoreCase(repository.leerCelda(candidata, 0))) {
                    fila = candidata;
                    break;
                }
            }
            if (fila == null) {
                throw new IOException("No se encontró el producto " + codigoExistente + ".");
            }
            codigo.setText(repository.leerCelda(fila, 0));
            sku.setText(repository.leerCelda(fila, 11));
            descripcion.setText(repository.leerCelda(fila, 1));
            asignarCombo(unidad, repository.leerCelda(fila, 2));
            precio.setText(repository.leerCelda(fila, 3));
            cantidad.setText(repository.leerCelda(fila, 4));
            asignarCombo(categoria, repository.leerCelda(fila, 6));
            marca.setText(repository.leerCelda(fila, 7));
            stockMinimo.setText(repository.leerCelda(fila, 8));
            observaciones.setText(repository.leerCelda(fila, 9));
            rutaImagen = repository.leerCelda(fila, 10);
            if (!rutaImagen.isBlank()) {
                File archivoImagen = new File(rutaImagen);
                if (archivoImagen.isFile()) {
                    mostrarImagen(archivoImagen, ImageIO.read(archivoImagen));
                }
            }
        } catch (IOException ex) {
            javax.swing.JOptionPane.showMessageDialog(this, ex.getMessage(),
                    "No se pudo cargar el producto.", javax.swing.JOptionPane.ERROR_MESSAGE);
            volver.run();
        }
    }

    private void asignarCombo(JComboBox<String> combo, String valor) {
        for (int indice = 0; indice < combo.getItemCount(); indice++) {
            if (valor.equalsIgnoreCase(combo.getItemAt(indice))) {
                combo.setSelectedIndex(indice);
                return;
            }
        }
        if (!valor.isBlank()) {
            combo.insertItemAt(valor, 0);
            combo.setSelectedIndex(0);
        }
    }

    private void mostrarImagen(File archivo, BufferedImage original) throws IOException {
        if (original == null) {
            throw new IOException("Formato no compatible.");
        }
        rutaImagen = archivo.getAbsolutePath();
        int ancho = Math.max(1, Math.min(210, original.getWidth()));
        int alto = Math.max(1, Math.min(150, original.getHeight()));
        double escala = Math.min((double) ancho / original.getWidth(),
                (double) alto / original.getHeight());
        Image escalada = original.getScaledInstance(
                Math.max(1, (int) (original.getWidth() * escala)),
                Math.max(1, (int) (original.getHeight() * escala)), Image.SCALE_SMOOTH);
        imagen.setText(null);
        imagen.setIcon(new javax.swing.ImageIcon(escalada));
    }

    private void guardar() {
        String error = InventarioValidator.validar(codigo.getText(), descripcion.getText());
        if (error != null) {
            javax.swing.JOptionPane.showMessageDialog(this, error);
            return;
        }
        try {
            repository.guardarProducto(codigo.getText(), descripcion.getText(), texto(unidad),
                    precio.getText(), cantidad.getText(), texto(categoria), marca.getText().trim(),
                    stockMinimo.getText(), observaciones.getText(), rutaImagen, sku.getText());
            String mensaje = codigoInicial == null || codigoInicial.isBlank()
                    ? "Producto creado correctamente." : "Producto actualizado correctamente.";
            javax.swing.JOptionPane.showMessageDialog(this, mensaje);
            volver.run();
        } catch (IOException | IllegalArgumentException ex) {
            javax.swing.JOptionPane.showMessageDialog(this, ex.getMessage(),
                    "No se pudo guardar el producto.", javax.swing.JOptionPane.ERROR_MESSAGE);
        }
    }

    private String texto(JComboBox<String> combo) {
        Object valor = combo.getEditor().getItem();
        return valor == null ? "" : valor.toString().trim();
    }

    private void limpiar() {
        codigo.setText("");
        sku.setText("");
        descripcion.setText("");
        precio.setText("");
        cantidad.setText("");
        stockMinimo.setText("");
        observaciones.setText("");
        categoria.setSelectedIndex(0);
        marca.setText("");
        unidad.setSelectedIndex(0);
        rutaImagen = "";
        imagen.setIcon(null);
        imagen.setText("Agregar imagen");
    }
}
