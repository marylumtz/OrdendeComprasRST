package com.mycompany.ocxrst;

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
import javax.swing.plaf.basic.BasicButtonUI;

public final class EstiloAplicacion {

    private static final Font FUENTE_BASE = new Font("Segoe UI", Font.PLAIN, 12);
    private static final Locale LOCALE_ES = Locale.forLanguageTag("es-MX");

    private EstiloAplicacion() {
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
        boolean primario = fondoOriginal != null
                && fondoOriginal.getAlpha() > 0
                && !esClaro(fondoOriginal);
        Color azul = new Color(31, 112, 184);

        boton.setUI(new BasicButtonUI());
        boton.setOpaque(true);
        boton.setContentAreaFilled(true);
        boton.setFocusPainted(false);
        boton.setRolloverEnabled(true);
        boton.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        boton.setFont(new Font("Segoe UI", Font.BOLD, 11));
        boton.setForeground(primario ? Color.WHITE : azul);
        boton.setBackground(primario ? azul : Color.WHITE);
        boton.setBorder(new BordeBoton(8, primario ? azul : azul, 1.4f));
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
}
