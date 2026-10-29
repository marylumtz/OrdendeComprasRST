package com.mycompany.ocxrst.util;

import java.awt.Font;
import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.util.Locale;
import javax.swing.AbstractButton;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.UIManager;
import javax.swing.border.AbstractBorder;
import javax.swing.JComponent;
import javax.swing.ButtonModel;
import javax.swing.plaf.basic.BasicButtonUI;
import javax.swing.text.JTextComponent;

public final class EstiloAplicacion {

    private static final String FAMILIA_TIPOGRAFICA = "Segoe UI";
    public static final Color FONDO_PANTALLA = new Color(240, 246, 253);
    private static final Font FUENTE_BASE = fuente(Font.PLAIN, 12);
    private static final Locale LOCALE_ES = Locale.forLanguageTag("es-MX");

    private EstiloAplicacion() {
    }

    /** Crea una fuente de la familia tipográfica oficial de la aplicación. */
    public static Font fuente(int estilo, float tamano) {
        return new Font(FAMILIA_TIPOGRAFICA, estilo, Math.round(tamano));
    }

    /**
     * Normaliza la familia tipográfica de todos los controles de una pantalla,
     * conservando el tamaño y el estilo definidos por cada control.
     */
    public static void aplicarTipografia(Component raiz) {
        if (raiz == null) {
            return;
        }
        Font fuenteActual = raiz.getFont();
        if (fuenteActual != null && !"Segoe UI Symbol".equals(fuenteActual.getFamily())) {
            raiz.setFont(fuente(fuenteActual.getStyle(), fuenteActual.getSize2D()));
        }
        if (raiz instanceof Container contenedor) {
            for (Component hijo : contenedor.getComponents()) {
                aplicarTipografia(hijo);
            }
        }
    }

    /** Aplica la tipografía compartida a un botón creado dinámicamente. */
    public static void aplicarTipografiaBoton(AbstractButton boton) {
        if (boton != null) {
            Font fuenteActual = boton.getFont();
            int estilo = fuenteActual == null ? Font.PLAIN : fuenteActual.getStyle();
            float tamano = fuenteActual == null ? FUENTE_BASE.getSize2D() : fuenteActual.getSize2D();
            boton.setFont(fuente(estilo, tamano));
        }
    }

    /**
     * Aplica el campo de escritura libre compartido por todas las pantallas.
     * Los combos, tablas y botones no se modifican.
     */
    public static void aplicarEstiloCamposTexto(Component raiz) {
        if (raiz == null) {
            return;
        }
        if (raiz instanceof javax.swing.JComboBox<?> combo) {
            combo.setFont(fuente(Font.PLAIN, 14));
            javax.swing.border.Border bordeOriginal = combo.getBorder();
            combo.setBorder(new javax.swing.border.CompoundBorder(
                    bordeOriginal, new TrianguloComboBorder()));
            java.awt.Dimension preferido = combo.getPreferredSize();
            combo.setPreferredSize(new java.awt.Dimension(preferido.width, 30));
            combo.setMinimumSize(new java.awt.Dimension(0, 30));
            return;
        }
        if (raiz instanceof JTextComponent campo) {
            campo.setFont(fuente(Font.PLAIN, 14));
            campo.setForeground(new Color(24, 58, 104));
            campo.setBackground(Color.WHITE);
            campo.setOpaque(true);
            campo.setBorder(new CampoTextoBorder());
            campo.setMargin(new java.awt.Insets(4, 10, 4, 10));
            if (campo instanceof javax.swing.JTextField) {
                java.awt.Dimension preferido = campo.getPreferredSize();
                campo.setPreferredSize(new java.awt.Dimension(preferido.width, 30));
                campo.setMinimumSize(new java.awt.Dimension(0, 30));
            }
        }
        if (raiz instanceof Container contenedor) {
            for (Component hijo : contenedor.getComponents()) {
                aplicarEstiloCamposTexto(hijo);
            }
        }
    }

    public static void aplicarTipografiaGlobal() {
        Locale.setDefault(LOCALE_ES);
        UIManager.getDefaults().setDefaultLocale(LOCALE_ES);
        UIManager.put("OptionPane.yesButtonText", "Sí");
        UIManager.put("OptionPane.noButtonText", "No");
        UIManager.put("OptionPane.okButtonText", "Aceptar");
        UIManager.put("OptionPane.cancelButtonText", "Cancelar");
        UIManager.put("FileChooser.openButtonText", "Abrir");
        UIManager.put("FileChooser.saveButtonText", "Guardar");
        UIManager.put("FileChooser.cancelButtonText", "Cancelar");
        UIManager.put("FileChooser.saveDialogTitleText", "Guardar");
        UIManager.put("FileChooser.openDialogTitleText", "Abrir");
        UIManager.put("FileChooser.lookInLabelText", "Buscar en:");
        UIManager.put("FileChooser.saveInLabelText", "Guardar en:");
        UIManager.put("FileChooser.fileNameLabelText", "Nombre:");
        UIManager.put("FileChooser.filesOfTypeLabelText", "Tipo:");
        UIManager.put("FileChooser.upFolderToolTipText", "Subir un nivel");
        UIManager.put("FileChooser.homeFolderToolTipText", "Carpeta principal");
        UIManager.put("FileChooser.newFolderToolTipText", "Crear carpeta");
        UIManager.put("FileChooser.listViewButtonToolTipText", "Lista");
        UIManager.put("FileChooser.detailsViewButtonToolTipText", "Detalles");
        String[] componentes = {
            "Label", "Button", "ToggleButton", "CheckBox", "RadioButton",
            "ComboBox", "TextField", "PasswordField", "TextArea", "TextPane",
            "EditorPane", "FormattedTextField", "Spinner", "Table",
            "TableHeader", "List", "Tree", "Menu", "MenuItem",
            "CheckBoxMenuItem", "RadioButtonMenuItem", "TabbedPane",
            "TitledBorder", "ToolTip", "OptionPane", "FileChooser"
        };
        for (String componente : componentes) {
            UIManager.put(componente + ".font", FUENTE_BASE);
        }
        configurarEstiloCamposTextoGlobal();
    }

    private static void configurarEstiloCamposTextoGlobal() {
        String[] camposTexto = {
            "TextField", "PasswordField", "FormattedTextField",
            "TextArea", "TextPane", "EditorPane"
        };
        javax.swing.border.Border borde = new CampoTextoBorder();
        java.awt.Insets margen = new java.awt.Insets(4, 10, 4, 10);
        Color texto = new Color(24, 58, 104);

        for (String campoTexto : camposTexto) {
            UIManager.put(campoTexto + ".font", fuente(Font.PLAIN, 14));
            UIManager.put(campoTexto + ".foreground", texto);
            UIManager.put(campoTexto + ".background", Color.WHITE);
            UIManager.put(campoTexto + ".border", borde);
            UIManager.put(campoTexto + ".margin", margen);
        }
    }

    public static JFileChooser crearSelectorArchivos() {
        JFileChooser selector = new JFileChooser();
        selector.setLocale(LOCALE_ES);
        selector.updateUI();
        return selector;
    }

    /**
     * Applies the shared button language used throughout the application.
     * Existing non-white backgrounds identify primary actions; all other
     * buttons use the outlined secondary variant.
     */
    public static void aplicarEstiloBotones(Container raiz) {
        for (Component componente : raiz.getComponents()) {
            if (componente instanceof JButton boton) {
                estilizarBoton(boton);
            }
            if (componente instanceof Container contenedor) {
                aplicarEstiloBotones(contenedor);
            }
        }
    }

    private static void estilizarBoton(AbstractButton boton) {
        Color fondoOriginal = boton.getBackground();
        boolean peligro = fondoOriginal != null
                && fondoOriginal.getRed() > fondoOriginal.getBlue() + 35
                && fondoOriginal.getRed() > fondoOriginal.getGreen() + 20;
        boolean primario = fondoOriginal != null
                && fondoOriginal.getAlpha() > 0
                && !esClaro(fondoOriginal)
                && !peligro;
        Color azul = new Color(31, 112, 184);
        Color fondo = peligro ? new Color(211, 65, 72) : (primario ? azul : Color.WHITE);
        Color texto = primario || peligro ? Color.WHITE : azul;

        boton.setUI(new BotonRedondeadoUI());
        boton.setOpaque(false);
        boton.setContentAreaFilled(false);
        boton.setFocusPainted(false);
        boton.setRolloverEnabled(true);
        boton.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        boton.setFont(fuente(Font.BOLD, 11));
        boton.setForeground(texto);
        boton.setBackground(fondo);
        boton.setBorder(new BordeBoton(9, peligro ? fondo : azul, 1.4f));
        boton.setMargin(new java.awt.Insets(9, 16, 9, 16));
    }

    private static boolean esClaro(Color color) {
        return color.getRed() > 220 && color.getGreen() > 220 && color.getBlue() > 220;
    }

    private static final class BordeBoton extends AbstractBorder {
        private final int radio;
        private final Color color;
        private final float grosor;

        private BordeBoton(int radio, Color color, float grosor) {
            this.radio = radio;
            this.color = color;
            this.grosor = grosor;
        }

        @Override
        public void paintBorder(Component componente, java.awt.Graphics graphics,
                int x, int y, int width, int height) {
            java.awt.Graphics2D g2 = (java.awt.Graphics2D) graphics.create();
            try {
                g2.setColor(color);
                g2.setStroke(new java.awt.BasicStroke(grosor));
                g2.drawRoundRect(x + 1, y + 1, width - 3, height - 3, radio, radio);
            } finally {
                g2.dispose();
            }
        }
    }

    private static final class CampoTextoBorder extends AbstractBorder {
        private static final Color COLOR = new Color(198, 210, 225);

        @Override
        public java.awt.Insets getBorderInsets(Component componente) {
            return new java.awt.Insets(1, 1, 1, 1);
        }

        @Override
        public void paintBorder(Component componente, java.awt.Graphics graphics,
                int x, int y, int width, int height) {
            java.awt.Graphics2D g2 = (java.awt.Graphics2D) graphics.create();
            try {
                g2.setColor(COLOR);
                g2.setStroke(new java.awt.BasicStroke(1.2f));
                g2.drawRoundRect(x, y, width - 1, height - 1, 8, 8);
            } finally {
                g2.dispose();
            }
        }
    }

    private static final class TrianguloComboBorder extends AbstractBorder {
        @Override
        public java.awt.Insets getBorderInsets(Component componente) {
            return new java.awt.Insets(0, 0, 0, 0);
        }

        @Override
        public void paintBorder(Component componente, java.awt.Graphics graphics,
                int x, int y, int width, int height) {
            java.awt.Graphics2D g2 = (java.awt.Graphics2D) graphics.create();
            try {
                g2.setColor(new Color(31, 82, 132));
                int centroX = x + width - 11;
                int centroY = y + height / 2;
                java.awt.Polygon triangulo = new java.awt.Polygon(
                        new int[]{centroX - 5, centroX + 5, centroX},
                        new int[]{centroY - 2, centroY - 2, centroY + 4}, 3);
                g2.fillPolygon(triangulo);
            } finally {
                g2.dispose();
            }
        }
    }

    private static final class BotonRedondeadoUI extends BasicButtonUI {
        @Override
        public void paint(java.awt.Graphics graphics, JComponent componente) {
            AbstractButton boton = (AbstractButton) componente;
            ButtonModel modelo = boton.getModel();
            Color fondo = boton.getBackground();
            if (modelo.isPressed()) {
                fondo = mezclar(fondo, Color.BLACK, 0.16f);
            } else if (modelo.isRollover()) {
                fondo = mezclar(fondo, Color.BLACK, 0.07f);
            }

            java.awt.Graphics2D g2 = (java.awt.Graphics2D) graphics.create();
            try {
                g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                        java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
                g2.setColor(fondo);
                g2.fillRoundRect(0, 0, componente.getWidth() - 1,
                        componente.getHeight() - 1, 9, 9);
            } finally {
                g2.dispose();
            }
            super.paint(graphics, componente);
        }
    }

    private static Color mezclar(Color base, Color destino, float proporcion) {
        return new Color(
                base.getRed() + Math.round((destino.getRed() - base.getRed()) * proporcion),
                base.getGreen() + Math.round((destino.getGreen() - base.getGreen()) * proporcion),
                base.getBlue() + Math.round((destino.getBlue() - base.getBlue()) * proporcion));
    }
}
