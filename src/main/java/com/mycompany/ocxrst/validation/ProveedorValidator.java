package com.mycompany.ocxrst.validation;

public final class ProveedorValidator {

    private ProveedorValidator() {
    }

    public static String validar(String nombreRazonSocial, String rfc, boolean solicitanteSeleccionado) {
        if (nombreRazonSocial == null || nombreRazonSocial.trim().isEmpty()) {
            return "NOMBRE O RAZON SOCIAL es obligatorio.";
        }
        if (rfc == null || rfc.trim().isEmpty()) {
            return "RFC es obligatorio.";
        }
        if (!solicitanteSeleccionado) {
            return "Selecciona el NOMBRE DEL SOLICITANTE desde USUARIOS.xlsx.";
        }
        return null;
    }
}
