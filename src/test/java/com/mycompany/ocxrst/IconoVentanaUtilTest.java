package com.mycompany.ocxrst;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import org.junit.jupiter.api.Test;

class IconoVentanaUtilTest {

    @Test
    void aplicarConVentanaNulaNoLanzaExcepcion() {
        assertDoesNotThrow(() -> IconoVentanaUtil.aplicar(null));
    }
}