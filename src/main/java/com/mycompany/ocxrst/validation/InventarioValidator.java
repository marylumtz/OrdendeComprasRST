package com.mycompany.ocxrst.validation;

public final class InventarioValidator {

    private InventarioValidator() {
    }

    public static String validar(String codigo, String descripcion) {
        if (codigo == null || codigo.trim().isEmpty()) {
            return "CÓDIGO es obligatorio.";
        }
        if (descripcion == null || descripcion.trim().isEmpty()) {
            return "DESCRIPCIÓN es obligatoria.";
        }
        return null;
    }
}
