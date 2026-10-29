package com.mycompany.ocxrst;

import com.mycompany.ocxrst.util.IconoVentanaUtil;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import org.junit.jupiter.api.Test;

class IconoVentanaUtilTest {

    @Test
    void aplicarConVentanaNulaNoLanzaExcepcion() {
        assertDoesNotThrow(() -> IconoVentanaUtil.aplicar(null));
    }
}