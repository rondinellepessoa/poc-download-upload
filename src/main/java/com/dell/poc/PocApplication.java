package com.dell.poc;

import java.util.Arrays;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.dell.poc.entity.Produto;
import com.dell.poc.repository.ProdutoRepository;

@SpringBootApplication
public class PocApplication implements CommandLineRunner{

	@Autowired
	private ProdutoRepository produtoRepository;
	
	public static void main(String[] args) {
		SpringApplication.run(PocApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {
		Produto p1 = new Produto(null, "teste", 20.00);
		Produto p2 = new Produto(null, "teste", 20.00);
		Produto p3 = new Produto(null, "teste", 20.00);
		Produto p4 = new Produto(null, "teste", 20.00);
		produtoRepository.saveAll(Arrays.asList(p1, p2, p3, p4));
	}

}
