package com.mycompany.ocxrst;

import java.awt.Font;
import java.util.Locale;
import javax.swing.JFileChooser;
import javax.swing.UIManager;

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
}
