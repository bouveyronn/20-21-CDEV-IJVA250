package com.example.demo.service;

import com.example.demo.entity.Client;

import java.util.List;

public interface ClientService {
    List<Client> findAllClients();
    Client findClientById(Long id);
}
