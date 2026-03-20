package com.example.planilhas;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.*;
import javafx.geometry.*;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.layout.*;
import javafx.stage.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.pdf.*;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;

import java.io.*;
import java.net.URL;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.prefs.Preferences;
import org.apache.poi.ss.usermodel.Cell;
import com.itextpdf.io.font.PdfEncodings;

public class GeradorPdfAlunos extends Application {

    private File arquivoExcel;
    private File pastaDestino;
    private Label lblArquivo;
    private Label lblPasta;
    private Label statusLabel;
    private ProgressBar progressBar;
    private TableView<Map<String, String>> tableView;
    private Preferences prefs;
    private static final SimpleDateFormat FORMATO_BR = new SimpleDateFormat("dd/MM/yyyy HH:mm");

    @Override
    public void start(Stage stage) {
        prefs = Preferences.userNodeForPackage(GeradorPdfAlunos.class);

        stage.setTitle("Gerador de PDFs por Aluno");
        InputStream iconStream = getClass().getResourceAsStream("/icon.png");
        if (iconStream != null) {
            stage.getIcons().add(new Image(iconStream));
        }

        BorderPane root = new BorderPane();
        root.setStyle("-fx-background-color: #121212;");

        Label titulo = new Label("Gerador de PDFs por Aluno");
        titulo.setStyle("-fx-font-size: 22px; -fx-font-weight: bold; -fx-text-fill: white;");

        Button btnExcel = new Button("📂 Selecionar Excel");
        Button btnPasta = new Button("📁 Pasta Destino");
        Button btnGerar = new Button("🚀 Gerar PDFs");

        stylePrimary(btnExcel, "#2563eb");
        stylePrimary(btnPasta, "#7c3aed");
        stylePrimary(btnGerar, "#16a34a");

        lblArquivo = new Label("Nenhum arquivo selecionado");
        lblPasta = new Label("Nenhuma pasta selecionada");
        styleMuted(lblArquivo);
        styleMuted(lblPasta);

        progressBar = new ProgressBar(0);
        progressBar.setPrefWidth(400);
        progressBar.setVisible(false);

        statusLabel = new Label("");
        statusLabel.setStyle("-fx-text-fill: #38bdf8;");

        tableView = new TableView<>();
        tableView.setPrefHeight(200);
        tableView.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        tableView.setStyle("-fx-background-color: #1e1e1e; -fx-control-inner-background: #1e1e1e; -fx-text-fill: white;");

        VBox topBox = new VBox(10, titulo);
        topBox.setPadding(new Insets(20));
        root.setTop(topBox);

        VBox centerBox = new VBox(12,
                btnExcel, lblArquivo,
                btnPasta, lblPasta,
                btnGerar,
                progressBar,
                statusLabel,
                new Label("Pré-visualização dos dados:"),
                tableView
        );
        centerBox.setPadding(new Insets(20));
        root.setCenter(centerBox);

        btnExcel.setOnAction(e -> selecionarExcel(stage));
        btnPasta.setOnAction(e -> selecionarPasta(stage));
        btnGerar.setOnAction(e -> gerarPdfs());

        restaurarPreferencias();

        Scene scene = new Scene(root, 750, 600);
        stage.setScene(scene);
        stage.show();
    }

    private void stylePrimary(Button btn, String color) {
        btn.setPrefWidth(240);
        btn.setStyle("""
            -fx-background-color: %s;
            -fx-text-fill: white;
            -fx-font-size: 14px;
            -fx-font-weight: bold;
            -fx-background-radius: 8;
            -fx-cursor: hand;
        """.formatted(color));
    }

    private void styleMuted(Label lbl) {
        lbl.setStyle("-fx-text-fill: #9ca3af; -fx-font-size: 12px;");
    }

    private void selecionarExcel(Stage stage) {
        FileChooser fc = new FileChooser();
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Arquivos Excel", "*.xlsx"));
        File file = fc.showOpenDialog(stage);
        if (file != null) {
            arquivoExcel = file;
            lblArquivo.setText("📄 " + file.getName());
            prefs.put("ultimoExcel", file.getAbsolutePath());
            carregarPreview();
        }
    }

    private void selecionarPasta(Stage stage) {
        DirectoryChooser dc = new DirectoryChooser();
        File dir = dc.showDialog(stage);
        if (dir != null) {
            pastaDestino = dir;
            lblPasta.setText("📂 " + dir.getAbsolutePath());
            prefs.put("ultimaPasta", dir.getAbsolutePath());
        }
    }

    private void restaurarPreferencias() {
        String ultimoExcel = prefs.get("ultimoExcel", null);
        if (ultimoExcel != null && new File(ultimoExcel).exists()) {
            arquivoExcel = new File(ultimoExcel);
            lblArquivo.setText("📄 " + arquivoExcel.getName());
            carregarPreview();
        }

        String ultimaPasta = prefs.get("ultimaPasta", null);
        if (ultimaPasta != null && new File(ultimaPasta).exists()) {
            pastaDestino = new File(ultimaPasta);
            lblPasta.setText("📂 " + pastaDestino.getAbsolutePath());
        }
    }

    private void carregarPreview() {
        try (FileInputStream fis = new FileInputStream(arquivoExcel);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            tableView.getItems().clear();
            tableView.getColumns().clear();

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return;

            int colCount = headerRow.getLastCellNum();

            for (int col = 0; col < colCount; col++) {
                final int index = col;
                TableColumn<Map<String, String>, String> column =
                        new TableColumn<>(headerRow.getCell(col).getStringCellValue());
                column.setCellValueFactory(data ->
                        new SimpleStringProperty(data.getValue().getOrDefault("col" + index, "")));
                tableView.getColumns().add(column);
            }

            ObservableList<Map<String, String>> data = FXCollections.observableArrayList();

            for (int i = 1; i <= Math.min(sheet.getLastRowNum(), 20); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, String> rowData = new HashMap<>();
                for (int col = 0; col < colCount; col++) {
                    Cell cell = row.getCell(col);
                    rowData.put("col" + col, formatarCelula(cell));
                }
                data.add(rowData);
            }

            tableView.setItems(data);

        } catch (Exception e) {
            statusLabel.setText("Erro ao carregar pré-visualização: " + e.getMessage());
            statusLabel.setStyle("-fx-text-fill: red;");
        }
    }

    private void gerarPdfs() {
        if (arquivoExcel == null || pastaDestino == null) {
            statusLabel.setText("⚠️ Selecione o Excel e a pasta destino.");
            statusLabel.setStyle("-fx-text-fill: orange;");
            return;
        }

        progressBar.setVisible(true);
        progressBar.setProgress(0);
        statusLabel.setText("Processando...");
        statusLabel.setStyle("-fx-text-fill: #38bdf8;");

        new Thread(() -> {
            try (FileInputStream fis = new FileInputStream(arquivoExcel);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(0);
                Map<String, List<Row>> alunos = new LinkedHashMap<>();

                Row headerRow = sheet.getRow(0);
                if (headerRow == null) throw new RuntimeException("Planilha sem cabeçalho.");

                int colCount = headerRow.getLastCellNum();

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    Cell cellNome = row.getCell(1);
                    if (cellNome == null) continue;

                    String nomeAluno = cellNome.getStringCellValue().trim();
                    alunos.computeIfAbsent(nomeAluno, k -> new ArrayList<>()).add(row);
                }

                int total = alunos.size();
                int count = 0;

                for (Map.Entry<String, List<Row>> entry : alunos.entrySet()) {
                    String nomeAluno = entry.getKey();
                    List<Row> linhas = entry.getValue();

                    String nomeArquivoSeguro = nomeAluno.replaceAll("[\\\\/:*?\"<>|]", "_");
                    Path caminhoPdf = pastaDestino.toPath().resolve(nomeArquivoSeguro + ".pdf");

                    PdfWriter writer = new PdfWriter(caminhoPdf.toString());
                    PdfDocument pdf = new PdfDocument(writer);
                    Document document = new Document(pdf);

                    // Margens profissionais
                    document.setMargins(36, 36, 36, 36);

                    // Fontes
                    PdfFont fonteNormal = PdfFontFactory.createFont(StandardFonts.TIMES_ROMAN);
                    PdfFont fonteNegrito = PdfFontFactory.createFont(StandardFonts.TIMES_BOLD);


                    // Cores institucionais
                    com.itextpdf.kernel.colors.Color azulInstitucional =
                            new com.itextpdf.kernel.colors.DeviceRgb(30, 64, 175);
                    com.itextpdf.kernel.colors.Color cinza =
                            com.itextpdf.kernel.colors.ColorConstants.GRAY;

                    document.add(new Paragraph("Aluno: " + nomeAluno)
                            .setFont(fonteNegrito)
                            .setFontSize(14).setFontColor(azulInstitucional)
                            .setMarginTop(10));


                    // Separador visual
                    document.add(new Paragraph("────────────────────────────────────────────")
                            .setFont(fonteNormal)
                            .setFontSize(9)
                            .setMarginBottom(12));

                    int indiceRegistro = 1;

                    for (Row row : linhas) {

                        // Título da ficha

                        for (int col = 0; col < colCount; col++) {
                            String tituloColuna = headerRow.getCell(col).getStringCellValue();
                            String valorCelula = formatarCelula(row.getCell(col));

                            if (!valorCelula.isBlank()) {
                                document.add(new Paragraph(tituloColuna)
                                        .setFont(fonteNegrito)
                                        .setFontSize(10));

                                document.add(new Paragraph(valorCelula)
                                        .setFont(fonteNormal)
                                        .setFontSize(10)
                                        .setMarginBottom(8));
                            }
                        }

                        // Separador entre fichas
                        document.add(new Paragraph("────────────────────────────────────────────")
                                .setFont(fonteNormal)
                                .setFontSize(9)
                                .setMarginBottom(14));

                        indiceRegistro++;
                    }

                    document.close();






                    count++;
                    double progresso = (double) count / total;
                    final int finalCount = count;

                    Platform.runLater(() -> {
                        progressBar.setProgress(progresso);
                        statusLabel.setText("Gerado " + finalCount + " de " + total + " PDFs...");
                    });
                }

                Platform.runLater(() -> {
                    statusLabel.setText("✅ PDFs gerados com sucesso!");
                    statusLabel.setStyle("-fx-text-fill: #22c55e;");
                });

            } catch (Exception e) {
                e.printStackTrace();
                Platform.runLater(() -> {
                    statusLabel.setText("❌ Erro: " + e.getMessage());
                    statusLabel.setStyle("-fx-text-fill: red;");
                });
            }
        }).start();
    }

    private String formatarCelula(Cell cell) {
        if (cell == null) return "";

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    yield FORMATO_BR.format(cell.getDateCellValue());
                } else {
                    double valor = cell.getNumericCellValue();
                    if (valor == Math.floor(valor)) {
                        yield String.valueOf((long) valor);
                    } else {
                        yield String.valueOf(valor);
                    }
                }
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }
    public static void main(String[] args) {
        launch(args);
    }
}
