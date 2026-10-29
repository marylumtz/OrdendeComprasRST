package com.mycompany.ocxrst.service;

import com.mycompany.ocxrst.model.Usuario;

public final class SesionActual {

    private static Usuario usuario;

    private SesionActual() {
    }

    public static void iniciar(Usuario usuarioAutenticado) {
        usuario = usuarioAutenticado;
    }

    public static String usuario() {
        return usuario == null ? "" : usuario.getUsuario();
    }

    public static String nombre() {
        return usuario == null ? "" : usuario.getNombre();
    }

    public static boolean esAdministrador() {
        return "admin".equalsIgnoreCase(usuario());
    }

    public static void cerrar() {
        usuario = null;
    }
}
