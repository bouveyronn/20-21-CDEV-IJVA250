package com.example.demo.service;

import com.example.demo.entity.Facture;

import java.util.List;

public interface FactureService {
    List<Facture> findAllFactures();

    Facture findById(Long idFacture);
}
