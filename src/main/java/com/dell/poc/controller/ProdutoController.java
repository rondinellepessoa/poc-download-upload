package com.dell.poc.controller;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.dell.poc.download.ProdutoDownload;
import com.dell.poc.entity.Produto;
import com.dell.poc.service.ProdutoService;

@RestController
@RequestMapping(value="/produtos")
public class ProdutoController {
	
	@Autowired
	private ProdutoService produtoService;
	
	@GetMapping
	public ResponseEntity<List<Produto>> findAll(){
		List<Produto> produtos = produtoService.findAll();
		return ResponseEntity.ok().body(produtos);
	}
	
	@GetMapping("/export/excel")
    public void exportToExcel(HttpServletResponse response) throws IOException {        
         
        List<Produto> listaProdutos = produtoService.findAll();
        
        List<String> params = Arrays.asList("Produto", "Nome", "", "Quantidade", "", "Drop Down Test");
         
        ProdutoDownload excelExporter = new ProdutoDownload(listaProdutos, params);
         
        excelExporter.export(response);    
    } 
}
