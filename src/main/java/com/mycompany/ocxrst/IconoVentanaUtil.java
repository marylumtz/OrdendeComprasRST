package com.mycompany.ocxrst;

import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.Taskbar;
import java.awt.image.BufferedImage;
import java.io.File;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import javax.swing.ImageIcon;
import javax.swing.JFrame;

public final class IconoVentanaUtil {

    private static final String RUTA_ICONO_PRINCIPAL = "/com/mycompany/ocxrst/IMAGENES/Compu.png";
    private static final String RUTA_ICONO_RESPALDO = "/com/mycompany/ocxrst/IMAGENES/LOGO.jpg";
    private static final String RUTA_ICONO_PRINCIPAL_LOCAL = "src/main/java/com/mycompany/ocxrst/IMAGENES/Compu.png";
    private static final String RUTA_ICONO_RESPALDO_LOCAL = "src/main/java/com/mycompany/ocxrst/IMAGENES/LOGO.jpg";
    private static final int[] TAMANOS_ICONO = {16, 24, 32, 48, 64, 128, 256, 512};

    private static volatile List<Image> iconosCache;

    private IconoVentanaUtil() {
    }

    public static void aplicar(JFrame frame) {
        if (frame == null) {
            return;
        }

        List<Image> iconos = obtenerIconos();
        if (!iconos.isEmpty()) {
            frame.setIconImages(iconos);
            aplicarTaskbar(iconos.get(iconos.size() - 1));
        }
    }

    private static List<Image> obtenerIconos() {
        if (iconosCache != null) {
            return iconosCache;
        }

        synchronized (IconoVentanaUtil.class) {
            if (iconosCache != null) {
                return iconosCache;
            }

            Image iconoBase = cargarDesdeClasspath();
            if (iconoBase == null) {
                iconoBase = cargarDesdeArchivo();
            }
            if (iconoBase == null) {
                iconosCache = Collections.emptyList();
                return iconosCache;
            }

            iconosCache = construirVariantes(iconoBase);
            return iconosCache;
        }
    }

    private static List<Image> construirVariantes(Image iconoBase) {
        BufferedImage base = convertirABuffered(iconoBase);
        if (base == null) {
            return List.of(iconoBase);
        }

        List<Image> iconos = new ArrayList<>();
        for (int tamano : TAMANOS_ICONO) {
            iconos.add(escalar(base, tamano));
        }
        return Collections.unmodifiableList(iconos);
    }

    private static BufferedImage convertirABuffered(Image imagen) {
        if (imagen instanceof BufferedImage bufferedImage) {
            return bufferedImage;
        }

        ImageIcon icon = new ImageIcon(imagen);
        int ancho = icon.getIconWidth();
        int alto = icon.getIconHeight();
        if (ancho <= 0 || alto <= 0) {
            return null;
        }

        BufferedImage buffered = new BufferedImage(ancho, alto, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = buffered.createGraphics();
        g2.drawImage(imagen, 0, 0, null);
        g2.dispose();
        return buffered;
    }

    private static BufferedImage escalar(BufferedImage origen, int tamano) {
        BufferedImage salida = new BufferedImage(tamano, tamano, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = salida.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        g2.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

        int anchoOrigen = origen.getWidth();
        int altoOrigen = origen.getHeight();
        double escala = Math.min((double) tamano / anchoOrigen, (double) tamano / altoOrigen);
        int anchoDestino = Math.max(1, (int) Math.round(anchoOrigen * escala));
        int altoDestino = Math.max(1, (int) Math.round(altoOrigen * escala));
        int x = (tamano - anchoDestino) / 2;
        int y = (tamano - altoDestino) / 2;

        g2.drawImage(origen, x, y, anchoDestino, altoDestino, null);
        g2.dispose();
        return salida;
    }

    private static void aplicarTaskbar(Image icono) {
        try {
            if (!Taskbar.isTaskbarSupported()) {
                return;
            }
            Taskbar taskbar = Taskbar.getTaskbar();
            if (taskbar.isSupported(Taskbar.Feature.ICON_IMAGE)) {
                taskbar.setIconImage(icono);
            }
        } catch (UnsupportedOperationException | SecurityException ex) {
            // Si el sistema no soporta icono global, se mantiene solo por ventana.
        }
    }

    private static Image cargarDesdeClasspath() {
        Image iconoPrincipal = cargarImagenClasspath(RUTA_ICONO_PRINCIPAL);
        if (iconoPrincipal != null) {
            return iconoPrincipal;
        }
        return cargarImagenClasspath(RUTA_ICONO_RESPALDO);
    }

    private static Image cargarImagenClasspath(String ruta) {
        URL url = IconoVentanaUtil.class.getResource(ruta);
        if (url == null) {
            return null;
        }
        return new ImageIcon(url).getImage();
    }

    private static Image cargarDesdeArchivo() {
        Image iconoPrincipal = cargarImagenArchivo(RUTA_ICONO_PRINCIPAL_LOCAL);
        if (iconoPrincipal != null) {
            return iconoPrincipal;
        }
        return cargarImagenArchivo(RUTA_ICONO_RESPALDO_LOCAL);
    }

    private static Image cargarImagenArchivo(String ruta) {
        File archivo = new File(ruta);
        if (!archivo.exists()) {
            return null;
        }
        return new ImageIcon(archivo.getAbsolutePath()).getImage();
    }
}
