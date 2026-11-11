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
import javax.swing.text.JTextComponent;

public final class EstiloAplicacion {

    private static final String FAMILIA_TIPOGRAFICA = "Segoe UI";
    private static final String BOTON_ESTILO_EXENTO = "estiloAplicacion.botonExento";
    private static final String[] ESTADOS_FLECHA_COMBO = {
        "Enabled", "MouseOver", "Disabled", "Pressed", "Selected"
    };
    private static final javax.swing.Painter<JComponent> PINTOR_FLECHA_COMBO_VACIO
            = (graphics, componente, ancho, alto) -> { };
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
     * Aplica la apariencia estándar de las tablas del ERP: encabezado
     * gris azulado, filas alternadas, selección institucional y ordenamiento
     * al pulsar los encabezados.
     */
    public static void aplicarEstiloTabla(javax.swing.JTable tabla) {
            if (tabla == null) {
                return;
            }
            Color azulOscuro = new Color(28, 82, 132);
            Color azul = new Color(31, 112, 184);
            Color azulClaro = new Color(225, 239, 255);
            Color borde = new Color(225, 232, 241);
            Color fondoFila = new Color(248, 250, 253);
            Color fondoAlterno = new Color(242, 242, 242);
            Color texto = new Color(24, 58, 104);

            tabla.setRowHeight(32);
            tabla.setFont(fuente(Font.PLAIN, 11));
            tabla.setForeground(texto);
            tabla.setBackground(fondoFila);
            tabla.setGridColor(borde);
            tabla.setShowVerticalLines(true);
            tabla.setSelectionBackground(azulClaro);
            tabla.setSelectionForeground(texto);
            tabla.setAutoCreateRowSorter(true);

            javax.swing.table.JTableHeader encabezado = tabla.getTableHeader();
            encabezado.setFont(fuente(Font.BOLD, 10));
            encabezado.setBackground(new Color(232, 238, 247));
            encabezado.setForeground(azulOscuro);
            encabezado.setOpaque(true);

            tabla.setDefaultRenderer(Object.class, (t, valor, seleccionado, foco, fila, columna) -> {
                javax.swing.JLabel celda = new javax.swing.JLabel(valor == null ? "" : String.valueOf(valor));
                celda.setOpaque(true);
                celda.setBorder(javax.swing.BorderFactory.createEmptyBorder(0, 7, 0, 7));
                celda.setFont(fuente(Font.PLAIN, 11));
                celda.setForeground(seleccionado ? Color.WHITE : texto);
                celda.setBackground(seleccionado ? azul : (fila % 2 == 0 ? Color.WHITE : fondoAlterno));
                return celda;
            });
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
            anularPintoresFlechaCombo(UIManager.getDefaults());
            combo.setFont(fuente(Font.PLAIN, 14));
            javax.swing.border.Border bordeOriginal = combo.getBorder();
            combo.setBorder(new javax.swing.border.CompoundBorder(
                    bordeOriginal, new TrianguloComboBorder()));
            for (Component hijo : combo.getComponents()) {
                if (hijo instanceof JComponent componenteSwing
                        && "ComboBox.arrowButton".equals(componenteSwing.getName())) {
                    javax.swing.UIDefaults overrides = new javax.swing.UIDefaults();
                    anularPintoresFlechaCombo(overrides);
                    componenteSwing.putClientProperty("Nimbus.Overrides", overrides);
                    componenteSwing.putClientProperty("Nimbus.Overrides.InheritDefaults", Boolean.TRUE);
                    componenteSwing.updateUI();
                }
            }
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

    private static void anularPintoresFlechaCombo(javax.swing.UIDefaults defaults) {
        for (String estado : ESTADOS_FLECHA_COMBO) {
            defaults.put(
                    "ComboBox:\"ComboBox.arrowButton\"[" + estado + "].foregroundPainter",
                    PINTOR_FLECHA_COMBO_VACIO);
        }
    }

    /**
     * Aplica la apariencia compartida a un selector de fecha y a su calendario.
     * La vista muestra únicamente los siete días de la semana, sin columna de
     * números de semana.
     */
    public static void estilizarSelectorFecha(com.toedter.calendar.JDateChooser selector) {
        if (selector == null) {
            return;
        }

        Color azul = new Color(31, 112, 184);
        Color azulOscuro = new Color(28, 82, 132);
        Color azulClaro = new Color(225, 238, 250);
        Color borde = new Color(180, 202, 226);
        Font fuente = fuente(Font.PLAIN, 12);

        selector.setLocale(LOCALE_ES);
        selector.setDateFormatString("d MMM. yyyy");
        selector.setFont(fuente);
        selector.setOpaque(false);
        selector.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(borde),
                javax.swing.BorderFactory.createEmptyBorder(1, 4, 1, 1)));

        Component editor = selector.getDateEditor().getUiComponent();
        if (editor instanceof javax.swing.JTextField campoFecha) {
            campoFecha.setFont(fuente);
            campoFecha.setForeground(azulOscuro);
            campoFecha.setBackground(Color.WHITE);
            campoFecha.setOpaque(true);
            campoFecha.setBorder(javax.swing.BorderFactory.createEmptyBorder(4, 7, 4, 4));
        }

        javax.swing.JButton botonCalendario = selector.getCalendarButton();
        botonCalendario.setFont(fuente(Font.BOLD, 12));
        botonCalendario.setForeground(Color.WHITE);
        botonCalendario.setBackground(azul);
        botonCalendario.setOpaque(true);
        botonCalendario.setContentAreaFilled(true);
        botonCalendario.setBorder(javax.swing.BorderFactory.createEmptyBorder(2, 10, 2, 10));
        botonCalendario.setFocusPainted(false);
        botonCalendario.setPreferredSize(new java.awt.Dimension(40, 30));
        botonCalendario.setMinimumSize(new java.awt.Dimension(40, 30));
        botonCalendario.setMaximumSize(new java.awt.Dimension(40, 30));

        com.toedter.calendar.JCalendar calendario = selector.getJCalendar();
        calendario.setLocale(LOCALE_ES);
        calendario.setFont(fuente);
        calendario.setBackground(Color.WHITE);
        calendario.setForeground(azulOscuro);
        calendario.setDecorationBackgroundColor(azul);
        calendario.setWeekdayForeground(Color.WHITE);
        calendario.setSundayForeground(new Color(220, 235, 250));
        calendario.setDecorationBordersVisible(false);
        calendario.setTodayButtonVisible(true);
        calendario.setTodayButtonText("Hoy");
        calendario.setNullDateButtonVisible(true);
        calendario.setNullDateButtonText("Limpiar");
        calendario.setBorder(javax.swing.BorderFactory.createEmptyBorder(6, 6, 6, 6));
        calendario.setWeekOfYearVisible(false);

        com.toedter.calendar.JDayChooser dias = calendario.getDayChooser();
        dias.setDayBordersVisible(false);
        dias.setDecorationBordersVisible(false);
        dias.setFont(fuente);
        dias.setBackground(Color.WHITE);
        dias.setForeground(azulOscuro);
        dias.setDecorationBackgroundColor(azul);
        dias.setWeekdayForeground(Color.WHITE);
        dias.setSundayForeground(new Color(220, 235, 250));
        dias.setWeekOfYearVisible(false);
        dias.getDayPanel().setBorder(
                javax.swing.BorderFactory.createEmptyBorder(0, 4, 0, 4));

        agregarListaAnios(calendario);
        estilizarControlesCalendario(calendario, azul, azulOscuro, azulClaro, borde, fuente);
        blanquearFondosCalendario(calendario);
        quitarBordesInternosCalendario(calendario);
        refrescarDiasCalendario(selector, azul, azulOscuro);
        java.beans.PropertyChangeListener actualizarDias = e ->
                refrescarDiasCalendario(selector, azul, azulOscuro);
        selector.addPropertyChangeListener("date", actualizarDias);
        calendario.addPropertyChangeListener("calendar", actualizarDias);
        botonCalendario.addActionListener(e -> javax.swing.SwingUtilities.invokeLater(
                () -> {
                    quitarBordePopupCalendario(calendario);
                    refrescarDiasCalendario(selector, azul, azulOscuro);
                }));
        quitarBordePopupCalendario(calendario);
    }

    private static void agregarListaAnios(com.toedter.calendar.JCalendar calendario) {
        if (calendario.getClientProperty("listaAnios") != null) {
            return;
        }
        com.toedter.calendar.JYearChooser selectorAnio = calendario.getYearChooser();
        int anioActual = selectorAnio.getYear();
        int anioFinal = java.time.Year.now().getValue() + 10;
        selectorAnio.setStartYear(2000);
        selectorAnio.setEndYear(anioFinal);
        javax.swing.JComboBox<Integer> listaAnios = new javax.swing.JComboBox<>();
        for (int anio = 2000; anio <= anioFinal; anio++) {
            listaAnios.addItem(anio);
        }
        listaAnios.setSelectedItem(anioActual);
        listaAnios.setFont(fuente(Font.PLAIN, 12));
        listaAnios.setMaximumRowCount(8);
        listaAnios.setPrototypeDisplayValue(anioActual);
        listaAnios.setPreferredSize(new java.awt.Dimension(68, 24));
        listaAnios.setMinimumSize(new java.awt.Dimension(68, 24));
        listaAnios.setMaximumSize(new java.awt.Dimension(68, 24));
        listaAnios.addActionListener(e -> {
            Object seleccionado = listaAnios.getSelectedItem();
            if (seleccionado instanceof Integer anio) {
                calendario.getYearChooser().setYear(anio);
            }
        });
        calendario.getYearChooser().getSpinner().setVisible(false);
        calendario.getYearChooser().getSpinner().setPreferredSize(new java.awt.Dimension(0, 0));
        calendario.getYearChooser().getSpinner().setMinimumSize(new java.awt.Dimension(0, 0));
        calendario.getYearChooser().getSpinner().setMaximumSize(new java.awt.Dimension(0, 0));
        calendario.getYearChooser().getParent().add(listaAnios);
        calendario.putClientProperty("listaAnios", listaAnios);
    }

    private static void blanquearFondosCalendario(java.awt.Container contenedor) {
        for (Component componente : contenedor.getComponents()) {
            if (componente instanceof javax.swing.JButton
                    || componente instanceof javax.swing.JComboBox<?>) {
                continue;
            }
            componente.setBackground(Color.WHITE);
            if (componente instanceof JComponent componenteSwing) {
                componenteSwing.setOpaque(true);
            }
            if (componente instanceof Container hijo) {
                blanquearFondosCalendario(hijo);
            }
        }
    }

    private static void quitarBordesInternosCalendario(
            com.toedter.calendar.JCalendar calendario) {
        quitarBordeInterno(calendario.getMonthChooser());
        quitarBordeInterno(calendario.getDayChooser());
        quitarBordeInterno(calendario.getDayChooser().getDayPanel());
    }

    private static void quitarBordeInterno(Component componente) {
        if (componente instanceof JComponent componenteSwing) {
            componenteSwing.setBorder(javax.swing.BorderFactory.createEmptyBorder());
        }
        if (componente instanceof Container contenedor) {
            for (Component hijo : contenedor.getComponents()) {
                if (!(hijo instanceof javax.swing.JButton)
                        && !(hijo instanceof javax.swing.JComboBox<?>)
                        && !(hijo instanceof javax.swing.JSpinner)) {
                    quitarBordeInterno(hijo);
                }
            }
        }
    }

    private static void quitarBordePopupCalendario(Component calendario) {
        Container padre = calendario.getParent();
        while (padre != null) {
            if (padre instanceof javax.swing.JPopupMenu popup) {
                popup.setBorder(javax.swing.BorderFactory.createEmptyBorder());
                popup.setOpaque(true);
                popup.setBackground(Color.WHITE);
                popup.revalidate();
                popup.repaint();
                return;
            }
            padre = padre.getParent();
        }
    }

    private static void estilizarControlesCalendario(
            com.toedter.calendar.JCalendar calendario, Color azul, Color azulOscuro,
            Color azulClaro, Color borde, Font fuente) {
        estilizarContenedorCalendario(calendario.getMonthChooser(), azul, Color.WHITE, fuente);
        estilizarContenedorCalendario(calendario.getYearChooser(), azul, Color.WHITE, fuente);
        ocultarFlechasSpinner(calendario.getMonthChooser().getSpinner());
        ocultarFlechasSpinner(calendario.getYearChooser().getSpinner());
        estilizarEditorSpinner(calendario.getMonthChooser().getSpinner(), azulOscuro, fuente, borde);
        estilizarEditorSpinner(calendario.getYearChooser().getSpinner(), azulOscuro, fuente, borde);
        Component mesesComponente = calendario.getMonthChooser().getComboBox();
        if (mesesComponente instanceof javax.swing.JComboBox<?> meses) {
            calendario.getMonthChooser().setOpaque(false);
            calendario.getMonthChooser().setBackground(Color.WHITE);
            meses.updateUI();
            meses.setFont(fuente);
            meses.setForeground(javax.swing.UIManager.getColor("ComboBox.foreground"));
            meses.setBackground(javax.swing.UIManager.getColor("ComboBox.background"));
            meses.setOpaque(false);
            meses.setBorder(javax.swing.UIManager.getBorder("ComboBox.border"));
            int altoSelectorAnio = 24;
            Object selectorAnio = calendario.getClientProperty("listaAnios");
            if (selectorAnio instanceof Component componente) {
                altoSelectorAnio = componente.getPreferredSize().height;
            }
            java.awt.Dimension tamanoMes = new java.awt.Dimension(132, altoSelectorAnio);
            meses.setPreferredSize(tamanoMes);
            meses.setMinimumSize(tamanoMes);
            meses.setMaximumSize(tamanoMes);
        }
        estilizarBotonesCalendario(calendario, azulClaro, azulOscuro, fuente);
        igualarAlturaControlesCalendario(calendario);
    }

    private static void igualarAlturaControlesCalendario(
            com.toedter.calendar.JCalendar calendario) {
        int alturaBoton = 24;
        for (Component componente : calendario.getComponents()) {
            alturaBoton = Math.max(alturaBoton, obtenerAlturaBotonAuxiliar(componente));
        }

        Component meses = calendario.getMonthChooser().getComboBox();
        aplicarAltura(meses, alturaBoton);

        Object selectorAnio = calendario.getClientProperty("listaAnios");
        if (selectorAnio instanceof Component componente) {
            aplicarAltura(componente, alturaBoton);
        }
    }

    private static int obtenerAlturaBotonAuxiliar(Component componente) {
        if (componente instanceof javax.swing.JButton boton) {
            String texto = boton.getText() == null
                    ? "" : boton.getText().trim().toLowerCase(LOCALE_ES);
            if ("hoy".equals(texto) || "limpiar".equals(texto)) {
                return boton.getPreferredSize().height;
            }
        }
        if (componente instanceof Container contenedor) {
            int altura = 24;
            for (Component hijo : contenedor.getComponents()) {
                altura = Math.max(altura, obtenerAlturaBotonAuxiliar(hijo));
            }
            return altura;
        }
        return 24;
    }

    private static void aplicarAltura(Component componente, int altura) {
        java.awt.Dimension preferido = componente.getPreferredSize();
        java.awt.Dimension tamano = new java.awt.Dimension(preferido.width, altura);
        componente.setPreferredSize(tamano);
        componente.setMinimumSize(tamano);
        componente.setMaximumSize(tamano);
    }

    private static void ocultarFlechasSpinner(Component componente) {
        if (!(componente instanceof javax.swing.JSpinner spinner)) {
            return;
        }
        for (Component hijo : spinner.getComponents()) {
            if (hijo instanceof javax.swing.JButton boton) {
                boton.setVisible(false);
                boton.setPreferredSize(new java.awt.Dimension(0, 0));
                boton.setMinimumSize(new java.awt.Dimension(0, 0));
                boton.setMaximumSize(new java.awt.Dimension(0, 0));
            }
        }
        spinner.revalidate();
    }

    private static void estilizarEditorSpinner(Component componente,
            Color texto, Font fuente, Color borde) {
        if (!(componente instanceof javax.swing.JSpinner spinner)
                || !(spinner.getEditor() instanceof javax.swing.JSpinner.DefaultEditor editor)) {
            return;
        }
        javax.swing.JTextField campo = editor.getTextField();
        campo.setFont(fuente);
        campo.setForeground(texto);
        campo.setBackground(Color.WHITE);
        campo.setOpaque(true);
        campo.setHorizontalAlignment(javax.swing.JTextField.CENTER);
        campo.setBorder(javax.swing.BorderFactory.createCompoundBorder(
                javax.swing.BorderFactory.createLineBorder(borde, 1, true),
                javax.swing.BorderFactory.createEmptyBorder(1, 6, 1, 6)));
    }

    private static void estilizarContenedorCalendario(
            java.awt.Container contenedor, Color fondo, Color texto, Font fuente) {
        contenedor.setBackground(fondo);
        contenedor.setForeground(texto);
        contenedor.setFont(fuente);
        for (Component componente : contenedor.getComponents()) {
            componente.setFont(fuente);
            componente.setBackground(fondo);
            componente.setForeground(texto);
            if (componente instanceof JComponent componenteSwing) {
                componenteSwing.setOpaque(true);
                componenteSwing.setBorder(javax.swing.BorderFactory.createEmptyBorder(2, 4, 2, 4));
            }
            if (componente instanceof Container hijo) {
                estilizarContenedorCalendario(hijo, fondo, texto, fuente);
            }
        }
    }

    private static void estilizarBotonesCalendario(
            java.awt.Container contenedor, Color fondo, Color texto, Font fuente) {
        for (Component componente : contenedor.getComponents()) {
            if (componente instanceof javax.swing.JButton boton) {
                String etiqueta = boton.getText() == null
                        ? "" : boton.getText().trim().toLowerCase(LOCALE_ES);
                if ("hoy".equals(etiqueta) || "limpiar".equals(etiqueta)) {
                    boton.setFont(fuente);
                    boton.setForeground(texto);
                    boton.setBackground(fondo);
                    boton.setOpaque(true);
                    boton.setContentAreaFilled(true);
                    boton.setFocusPainted(false);
                    boton.setBorder(javax.swing.BorderFactory.createEmptyBorder(5, 10, 5, 10));
                }
            }
            if (componente instanceof Container hijo) {
                estilizarBotonesCalendario(hijo, fondo, texto, fuente);
            }
        }
    }

    private static void refrescarDiasCalendario(
            com.toedter.calendar.JDateChooser selector, Color azul, Color azulOscuro) {
        com.toedter.calendar.JDayChooser dias = selector.getJCalendar().getDayChooser();
        int diaSeleccionado = selector.getDate() == null ? -1 : dias.getDay();
        for (Component componente : dias.getDayPanel().getComponents()) {
            if (!(componente instanceof javax.swing.JButton boton)) {
                continue;
            }
            String texto = boton.getText() == null ? "" : boton.getText().trim();
            boolean esDia = texto.matches("\\d{1,2}");
            boton.setFocusPainted(false);
            boton.setBorderPainted(false);
            boton.setContentAreaFilled(true);
            boton.setRolloverEnabled(false);
            boton.setUI(new javax.swing.plaf.basic.BasicButtonUI());
            boton.setBorder(javax.swing.BorderFactory.createEmptyBorder(4, 5, 4, 5));
            boton.setFont(fuente(Font.PLAIN, 12));
            if (!esDia) {
                boton.setBackground(azul);
                boton.setForeground(Color.WHITE);
                continue;
            }
            boolean seleccionado = Integer.parseInt(texto) == diaSeleccionado;
            boton.setOpaque(true);
            boton.setBackground(seleccionado ? azul : Color.WHITE);
            boton.setForeground(seleccionado ? Color.WHITE : azulOscuro);
        }
        dias.getDayPanel().revalidate();
        dias.getDayPanel().repaint();
    }

    public static void aplicarTipografiaGlobal() {
        aplicarLookAndFeelNimbus();
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

    private static void aplicarLookAndFeelNimbus() {
        if (UIManager.getLookAndFeel() != null
                && "Nimbus".equals(UIManager.getLookAndFeel().getName())) {
            return;
        }
        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    return;
                }
            }
        } catch (ReflectiveOperationException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(EstiloAplicacion.class.getName())
                    .log(java.util.logging.Level.SEVERE, "No se pudo aplicar Nimbus", ex);
        }
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

    /** Applies a shared rounded shape while preserving each button's color role. */
    public static void aplicarEstiloBotones(Container raiz) {
        if (raiz == null) {
            return;
        }
        if (raiz instanceof JButton boton) {
            aplicarEstiloBoton(boton);
        }
        for (Component componente : raiz.getComponents()) {
            if (componente instanceof JButton boton) {
                aplicarEstiloBoton(boton);
            } else if (componente instanceof Container contenedor) {
                aplicarEstiloBotones(contenedor);
            }
        }
    }

    public static void aplicarEstiloBoton(AbstractButton boton) {
        if (boton != null
                && !Boolean.TRUE.equals(boton.getClientProperty(BOTON_ESTILO_EXENTO))) {
            estilizarBoton(boton);
        }
    }

    public static void excluirBotonDelEstiloComun(AbstractButton boton) {
        if (boton != null) {
            boton.putClientProperty(BOTON_ESTILO_EXENTO, Boolean.TRUE);
        }
    }

    private static void estilizarBoton(AbstractButton boton) {
        Color fondoOriginal = boton.getBackground();
        Color fondoPredeterminado = UIManager.getColor("Button.background");
        boolean esPredeterminado = fondoOriginal == null
                || fondoOriginal.equals(fondoPredeterminado);
        Color azul = new Color(31, 112, 184);
        Color fondo = esPredeterminado ? Color.WHITE : fondoOriginal;
        Color texto = esPredeterminado || esClaro(fondo) ? azul : Color.WHITE;

        boton.setFocusPainted(false);
        boton.setRolloverEnabled(true);
        boton.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        boton.setFont(fuente(Font.BOLD, 11));
        boton.setForeground(texto);
        boton.setBackground(fondo);
        boton.setBorder(javax.swing.BorderFactory.createEmptyBorder(7, 12, 7, 12));
    }

    private static boolean esClaro(Color color) {
        return color.getRed() > 220 && color.getGreen() > 220 && color.getBlue() > 220;
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

}
