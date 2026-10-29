package com.mycompany.ocxrst.service;

import com.mycompany.ocxrst.config.AppConfig;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.Reader;
import java.nio.charset.StandardCharsets;

public final class TerminosService {

    private static final String RECURSO_TERMINOS =
            "/com/mycompany/ocxrst/RECURSOS/Términos y Condiciones.txt";

    public String cargar() throws java.io.IOException {
        File archivo = AppConfig.terminosFile();
        if (archivo.exists()) {
            try (InputStream entrada = new FileInputStream(archivo)) {
                return leer(entrada);
            }
        }

        InputStream recurso = TerminosService.class.getResourceAsStream(RECURSO_TERMINOS);
        if (recurso == null) {
            throw new java.io.FileNotFoundException("No se encontró el archivo de términos.");
        }
        try (InputStream entrada = recurso) {
            return leer(entrada);
        }
    }

    public void guardar(String contenido) throws java.io.IOException {
        File archivo = AppConfig.terminosFile();
        File parent = archivo.getParentFile();
        if (parent != null) {
            parent.mkdirs();
        }

        try (java.io.Writer salida = new OutputStreamWriter(
                new FileOutputStream(archivo), StandardCharsets.UTF_8);
                BufferedWriter writer = new BufferedWriter(salida)) {
            writer.write(contenido == null ? "" : contenido);
        }
    }

    private String leer(InputStream entrada) throws java.io.IOException {
        StringBuilder contenido = new StringBuilder();
        try (Reader input = new InputStreamReader(entrada, StandardCharsets.UTF_8);
                BufferedReader reader = new BufferedReader(input)) {
            String linea;
            while ((linea = reader.readLine()) != null) {
                if (contenido.length() > 0) {
                    contenido.append(System.lineSeparator());
                }
                contenido.append(linea);
            }
        }
        return contenido.toString();
    }
}
