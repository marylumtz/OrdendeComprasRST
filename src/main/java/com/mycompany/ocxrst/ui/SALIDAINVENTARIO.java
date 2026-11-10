package com.mycompany.ocxrst.ui;

import com.mycompany.ocxrst.config.AppConfig;
import com.mycompany.ocxrst.repository.ActividadRepository;
import com.mycompany.ocxrst.repository.InventarioRepository;
import java.io.File;

public class SALIDAINVENTARIO extends ENTRADAINVENTARIO {
    public SALIDAINVENTARIO(Runnable volver) {
        this(volver, new InventarioRepository(), new ActividadRepository(),
                AppConfig.resolveBaseFile("MOVIMIENTOSINVENTARIO.xlsx"));
    }

    SALIDAINVENTARIO(Runnable volver, InventarioRepository repository,
            ActividadRepository actividadRepository, File archivoMovimientos) {
        super(volver, repository, actividadRepository, archivoMovimientos, TipoMovimiento.SALIDA);
    }
}
