package com.mycompany.ocxrst.validation;

public final class UsuarioValidator {

    private UsuarioValidator() {
    }

    public static String validar(String usuario, String nombre, String area, String puesto) {
        if (usuario == null || usuario.trim().isEmpty()) {
            return "NOMBRE DE USUARIO es obligatorio.";
        }
        if (nombre == null || nombre.trim().isEmpty()) {
            return "NOMBRE COMPLETO es obligatorio.";
        }
        if (area == null || area.trim().isEmpty()) {
            return "AREA O DEPARTAMENTO es obligatorio.";
        }
        if (puesto == null || puesto.trim().isEmpty()) {
            return "PUESTO es obligatorio.";
        }
        return null;
    }
}
