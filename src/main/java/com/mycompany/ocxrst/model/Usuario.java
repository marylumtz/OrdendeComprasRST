package com.mycompany.ocxrst.model;

public final class Usuario {

    private final String usuario;
    private final String nombre;

    public Usuario(String usuario, String nombre) {
        this.usuario = usuario == null ? "" : usuario.trim();
        this.nombre = nombre == null ? "" : nombre.trim();
    }

    public String getUsuario() {
        return usuario;
    }

    public String getNombre() {
        return nombre;
    }
}
