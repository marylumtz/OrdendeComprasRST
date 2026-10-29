package com.mycompany.ocxrst.service;

import com.mycompany.ocxrst.config.AppConfig;
import com.mycompany.ocxrst.model.Usuario;
import com.mycompany.ocxrst.repository.UsuarioRepository;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class AutenticacionService {

    private final UsuarioRepository usuarioRepository;

    public AutenticacionService() {
        this(new UsuarioRepository());
    }

    public AutenticacionService(UsuarioRepository usuarioRepository) {
        this.usuarioRepository = usuarioRepository;
    }

    public Optional<Usuario> autenticar(String usuario, String contrasena) {
        File archivo = AppConfig.resolveBaseFile("USUARIOS.xlsx");
        if (!archivo.exists()) {
            return Optional.empty();
        }

        try (FileInputStream entrada = new FileInputStream(archivo);
                Workbook libro = new XSSFWorkbook(entrada)) {
            return autenticar(libro, usuario, contrasena);
        } catch (IOException ex) {
            throw new IllegalStateException("No se pudo leer el archivo de usuarios.", ex);
        }
    }

    public Optional<Usuario> autenticar(Workbook libro, String usuario, String contrasena) {
        if (libro == null || usuario == null || contrasena == null) {
            return Optional.empty();
        }

        for (int indice = 1; indice <= libro.getSheetAt(0).getLastRowNum(); indice++) {
            Row fila = libro.getSheetAt(0).getRow(indice);
            if (fila == null) {
                continue;
            }

            String usuarioExcel = usuarioRepository.leerCelda(fila, 0);
            String contrasenaExcel = usuarioRepository.leerCelda(fila, 1);
            if (usuario.equals(usuarioExcel) && contrasena.equals(contrasenaExcel)) {
                return Optional.of(new Usuario(usuarioExcel, usuarioRepository.leerCelda(fila, 2)));
            }
        }
        return Optional.empty();
    }

    public String obtenerNombre(Workbook libro, String usuario, String contrasena) {
        return autenticar(libro, usuario, contrasena)
                .map(Usuario::getNombre)
                .orElse("");
    }
}
