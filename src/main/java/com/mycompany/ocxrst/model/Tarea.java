package com.mycompany.ocxrst.model;

import java.time.LocalDate;
import java.util.List;

public record Tarea(String folio, String titulo, String tipo, String estado,
        String prioridad, String responsable, LocalDate creacion, LocalDate vencimiento,
        String descripcion, LocalDate inicio, String folioRelacionado,
        String observaciones, List<String> adjuntos) {

    public static final List<String> TIPOS = List.of("Inventario", "Compra", "Proveedor", "Soporte", "General");
    public static final List<String> ESTADOS = List.of("Pendiente", "En proceso", "Completada", "Cancelada");
    public static final List<String> PRIORIDADES = List.of("Alta", "Media", "Baja");

    public Tarea {
        adjuntos = List.copyOf(adjuntos);
    }

    public void validar() {
        if (titulo == null || titulo.isBlank() || responsable == null || responsable.isBlank()) {
            throw new IllegalArgumentException("Completa el t\u00edtulo y el responsable.");
        }
        if (!TIPOS.contains(tipo) || !ESTADOS.contains(estado) || !PRIORIDADES.contains(prioridad)) {
            throw new IllegalArgumentException("Selecciona un tipo, estado y prioridad v\u00e1lidos.");
        }
        if (creacion == null || inicio == null || vencimiento == null) {
            throw new IllegalArgumentException("Completa las fechas de inicio y vencimiento.");
        }
        if (vencimiento.isBefore(inicio)) {
            throw new IllegalArgumentException("El vencimiento no puede ser anterior a la fecha de inicio.");
        }
    }
}
