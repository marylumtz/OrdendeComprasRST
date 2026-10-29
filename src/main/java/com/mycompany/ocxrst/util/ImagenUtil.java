package com.mycompany.ocxrst.util;

import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BaseMultiResolutionImage;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.net.URL;
import javax.imageio.ImageIO;
import javax.swing.ImageIcon;

public final class ImagenUtil {

    private ImagenUtil() {
    }

    public static ImageIcon cargarEscalada(String rutaClasspath, int ancho, int alto) {
        URL url = ImagenUtil.class.getResource(rutaClasspath);
        if (url == null) {
            return IconoVentanaUtil.obtenerIconoSeguro(rutaClasspath);
        }

        try {
            BufferedImage original = ImageIO.read(url);
            if (original == null) {
                return IconoVentanaUtil.obtenerIconoSeguro(rutaClasspath);
            }

            BufferedImage base = redimensionar(original, ancho, alto);
            BufferedImage alta = redimensionar(original, ancho * 2, alto * 2);
            BufferedImage muyAlta = redimensionar(original, ancho * 3, alto * 3);
            return new ImageIcon(new BaseMultiResolutionImage(base, alta, muyAlta));
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudo cargar la imagen: " + rutaClasspath, ex);
        }
    }

    private static BufferedImage redimensionar(BufferedImage original, int ancho, int alto) {
        BufferedImage escalada = new BufferedImage(ancho, alto, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = escalada.createGraphics();
        try {
            graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION,
                    RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            graphics.setRenderingHint(RenderingHints.KEY_RENDERING,
                    RenderingHints.VALUE_RENDER_QUALITY);
            graphics.setRenderingHint(RenderingHints.KEY_COLOR_RENDERING,
                    RenderingHints.VALUE_COLOR_RENDER_QUALITY);
            graphics.setRenderingHint(RenderingHints.KEY_DITHERING,
                    RenderingHints.VALUE_DITHER_DISABLE);
            graphics.setRenderingHint(RenderingHints.KEY_ALPHA_INTERPOLATION,
                    RenderingHints.VALUE_ALPHA_INTERPOLATION_QUALITY);
            graphics.drawImage(original, 0, 0, ancho, alto, null);
            return escalada;
        } finally {
            graphics.dispose();
        }
    }
}
