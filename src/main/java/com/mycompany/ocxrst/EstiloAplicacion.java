package com.mycompany.ocxrst;

import java.awt.Font;
import javax.swing.UIManager;

public final class EstiloAplicacion {

    private static final Font FUENTE_BASE = new Font("Segoe UI", Font.PLAIN, 12);

    private EstiloAplicacion() {
    }

    public static void aplicarTipografiaGlobal() {
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
}
