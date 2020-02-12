package com.example.demo.controller;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.impl.ClientServiceImpl;
import com.example.demo.service.ArticleService;
import com.example.demo.service.FactureService;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;

import com.lowagie.text.BadElementException;
import com.lowagie.text.Element;
import com.lowagie.text.PageSize;
import com.lowagie.text.Table;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.archive.RAFolderOutputStream;
import org.openqa.selenium.remote.http.HttpResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpRequest;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
/**
 * Controller principale pour affichage des clients / factures sur la page d'acceuil.
 */
@Controller
public class HomeController {

    //Déclaration des services
    private ArticleService articleService;
    private ClientServiceImpl clientServiceImpl;
    private FactureService factureService;

    //Constructeur
    public HomeController(ArticleService articleService, ClientServiceImpl clientService, FactureService factureService) {
        this.articleService = articleService;
        this.clientServiceImpl = clientService;
        this.factureService = factureService;
    }

    //Home
    @GetMapping("/")
    public ModelAndView home() {
        ModelAndView modelAndView = new ModelAndView("home");

        List<Article> articles = articleService.findAll();
        modelAndView.addObject("articles", articles);

        List<Client> toto = clientServiceImpl.findAllClients();
        modelAndView.addObject("clients", toto);

        List<Facture> factures = factureService.findAllFactures();
        modelAndView.addObject("factures", factures);

        return modelAndView;
    }

    //Lien "Inventaire des articles"
    @GetMapping("/articles")
    public ModelAndView inventaire(){
        ModelAndView modelAndView = new ModelAndView("articles");

        List<Article> lstArticles = articleService.findAll();
        modelAndView.addObject("articles", lstArticles);

        return modelAndView;
    }

    //Lien "Faire des achats"
    @GetMapping("/acheter")
    public ModelAndView shop(){
        ModelAndView modelAndView = new ModelAndView("acheter");

        List<Article> lstArticles = articleService.findAll();
        modelAndView.addObject("articles", lstArticles);

        List<Facture> lstFactures = factureService.findAllFactures();
        modelAndView.addObject("quantite", lstFactures);

        return modelAndView;
    }

    //Export des articles au format CSV
    @GetMapping("/articles/csv")
    public void articlesCSV(HttpServletRequest request, HttpServletResponse response)throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition","attachement; filename=\"export-articles.csv\"");
        PrintWriter writer = response.getWriter();

        //Appel du service
        List<Article> lstArticles = articleService.findAll();
        writer.println("Libellé du produit;Prix");
        for (Integer i = 0; i< lstArticles.size(); i++){
            writer.println(lstArticles.get(i).getLibelle() + ";" + lstArticles.get(i).getPrix());
        }

    }

    //Export des clients au format CSV
    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response)throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition","attachement; filename=\"export-clients.csv\"");
        PrintWriter writer = response.getWriter();

        //Appel du service
        List<Client> lstClients = clientServiceImpl.findAllClients();
        writer.println("NOM;PRENOM;AGE");
        for (Integer i = 0; i< lstClients.size(); i++){
            LocalDate dateNaissance = lstClients.get(i).getDateNaissance();
            writer.println(lstClients.get(i).getNom() + ";" + lstClients.get(i).getPrenom() + ";" + lstClients.get(i).getAge(dateNaissance));
        }

    }

    //Export des articles au format XLSX
    @GetMapping("/articles/xlsx")
    public void articlesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException{

        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"articles.xlsx\"");

        //1. Créer un Document vide
        XSSFWorkbook wb = new XSSFWorkbook();
        //2. Créer une Feuille de calcul vide
        Sheet feuille = wb.createSheet("export-articles");
        //3. Créer une ligne et mettre qlq chose dedans
        Row ligne_header = feuille.createRow((short)0);
        //4. Créer une Nouvelle cellule
        Cell cell_1 = ligne_header.createCell(0);
        Cell cell_2 = ligne_header.createCell(1);
        //5. Donner la valeur
        cell_1.setCellValue("Article");
        cell_2.setCellValue("Prix");

        //Index des lignes après en-tête
        Integer Row_Index = 1;

        //Appel du service
        List<Article> lstArticles = articleService.findAll();
        for (Integer i = 0; i< lstArticles.size(); i++){

            //Création d'une nouvelle ligne
            Row row = feuille.createRow(Row_Index);

            //Cellules articles et prix
            Cell cell_article = row.createCell(0);
            Cell cell_prix = row.createCell(1);

            cell_article.setCellValue(lstArticles.get(i).getLibelle());
            cell_prix.setCellValue(lstArticles.get(i).getPrix());

            //Ligne suivante
            Row_Index += 1;
        }

        try{
            wb.write(response.getOutputStream());
            wb.close();
        }catch (IOException e){
            e.printStackTrace();
        }

    }

    //Export des clients au format XLSX
    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException{
        //tentative infructueuse de la création du tableau ^^
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");

        //1. Créer un Document vide
        XSSFWorkbook wb = new XSSFWorkbook();
        //2. Créer une Feuille de calcul vide
        Sheet feuille = wb.createSheet("export-clients");
        //3. Créer une ligne et mettre qlq chose dedans
        Row ligne_header = feuille.createRow((short)0);
        //4. Créer une Nouvelle cellule
        Cell cell_1 = ligne_header.createCell(0);
        Cell cell_2 = ligne_header.createCell(1);
        Cell cell_3 = ligne_header.createCell(2);
        //5. Donner la valeur
        cell_1.setCellValue("NOM");
        cell_2.setCellValue("PRENOM");
        cell_3.setCellValue("AGE");

        //Index des lignes après en-tête
        Integer Row_Index = 1;

        //Appel du service
        List<Client> lstClients = clientServiceImpl.findAllClients();
        for (Integer i = 0; i< lstClients.size(); i++){

            //Création d'une nouvelle ligne
            Row row = feuille.createRow(Row_Index);

            //Cellules articles et prix
            Cell cell_nom = row.createCell(0);
            Cell cell_prenom = row.createCell(1);
            Cell cell_age = row.createCell(2);

            cell_nom.setCellValue(lstClients.get(i).getNom());
            cell_prenom.setCellValue(lstClients.get(i).getPrenom());

            LocalDate dateNaissance = lstClients.get(i).getDateNaissance();
            cell_age.setCellValue(lstClients.get(i).getAge(dateNaissance));

            //Ligne suivante
            Row_Index += 1;
        }

        try{
            wb.write(response.getOutputStream());
            wb.close();
        }catch (IOException e){
            e.printStackTrace();
        }

    }

    @GetMapping("/articles/create")
    public ModelAndView createArticle(HttpServletRequest request, HttpServletResponse response)throws IOException{
        ModelAndView modelAndView = new ModelAndView("article");

        List<Article> lstArticles = articleService.findAll();
        modelAndView.addObject("article", lstArticles);

        List<Facture> lstFactures = factureService.findAllFactures();
        modelAndView.addObject("quantite", lstFactures);

        return modelAndView;
    }

    @GetMapping("factures/{id}/pdf")
    public void facturespdf(HttpServletRequest request, HttpServletResponse response, @PathVariable("id") Long idFacture)throws IOException{
        response.setContentType("application/pdf");
        response.setHeader("Content-Disposition", "attachement;filename=\"facture.pdf\"");



        com.itextpdf.text.Document doc = new com.itextpdf.text.Document(com.itextpdf.text.PageSize.A4);
        try{

            //Récup des infos de la facture
           Facture facture = factureService.findById(idFacture);

            PdfWriter.getInstance(doc, response.getOutputStream());
            doc.open();
            doc.add(new Paragraph("Nom : " + facture.getClient().getNom()));
            doc.add(new Paragraph("Prénom : " + facture.getClient().getPrenom()));

            // Creating a table
            float [] pointColumnWidths = {150F, 150F, 150F};

            Table table = new Table(3);
            table.setBorderWidth(1);
            table.setBorderColor(new Color(0, 0, 255));
            table.setPadding(5);
            table.setSpacing(5);

            com.lowagie.text.Cell cell_Designation = new com.lowagie.text.Cell("Designation");
            cell_Designation.setHeader(true);
            cell_Designation.setColspan(1);

            com.lowagie.text.Cell cell_Quantite = new com.lowagie.text.Cell("Quantité");
            cell_Quantite.setHeader(true);
            cell_Quantite.setColspan(1);

            com.lowagie.text.Cell cell_PUHT = new com.lowagie.text.Cell("PUHT");
            cell_PUHT.setHeader(true);
            cell_PUHT.setColspan(1);

            table.addCell(cell_Designation);
            table.addCell(cell_Quantite);
            table.addCell(cell_PUHT);

            doc.add(new Paragraph("Désignation - quantité - PUHT"));
            for(LigneFacture lf : facture.getLigneFactures()){

                doc.add(new Paragraph(lf.getArticle().getLibelle() + " - " + lf.getQuantite() + " - " + lf.getSousTotal()/lf.getQuantite()));

                String article_libelle = lf.getArticle().getLibelle();
                cell_Designation = new com.lowagie.text.Cell(article_libelle);
                cell_Designation.setRowspan(2);
                cell_Designation.setColspan(1);

                Integer quantite = lf.getQuantite();
                cell_Quantite = new com.lowagie.text.Cell(quantite.toString());
                cell_Quantite.setRowspan(2);
                cell_Quantite.setColspan(1);

                Double PUHT = lf.getSousTotal()/lf.getQuantite();
                cell_PUHT = new com.lowagie.text.Cell(PUHT.toString());
                cell_PUHT.setRowspan(2);
                cell_PUHT.setColspan(1);

                table.addCell(cell_Designation);
                table.addCell(cell_Quantite);
                table.addCell(cell_PUHT);


            }
        } catch (DocumentException | BadElementException de){
            de.printStackTrace();
        }
        doc.close();
    }

    @GetMapping("/factures/xlsx")
    public void facturesXLSX(HttpServletRequest request, HttpServletResponse response)throws IOException{
        //TODO Export des factures EXCEL

    }
}
