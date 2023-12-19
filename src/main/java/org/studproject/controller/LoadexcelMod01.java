package org.studproject.controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumnModel;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;

/**
 *  Container do JFrame
 * * Este método realiza a comparação de duas colunas em uma planilha Excel, A e B
 * * importação de dados a partir de um arquivo Excel selecionado
 * * pelo usuário. Os dados são exibidos em uma tabela Swing e as diferenças nas colunas
 * * A e B são identificadas e apresentadas em uma área de JTable.
 *
 * JTable elemento apra gerar a tabela
 * JTextArea elemento para definir o resultado, da busca
 * JPanel Botão para importar a planilha
 * JProgressBar Elemento apra realizar Load do Processo de Busca da diferença entre as colunas
 * DefaultTableModel elemento apra Apresentar o resultado da leitura na tabela. apresenta os daods da tabela
 * JPanel Controlador Painel View Principal
 * JProgressBar Progress bar Load Linhas
 * DefaultTableModel Tabela visualizar dados da Tabela Excel Importada
 * JFileChooser
 * Set Lendo Coluna A da Tabela Excel - Elemento Pre Programado. sem elementos de Seleção de Coluna
 * Set Lendo Coluna B da Tabela Excel - Elemento Pre Programado. sem elementos de Seleção de Coluna
 * JFrame Tela do Load Carregando Separadamente do Projeto. JPanel
 *
 * */
public class LoadexcelMod01 extends JFrame {
    private JTable resultTable = new JTable();
    private JTextArea differencesTextArea = new JTextArea();
    private JPanel buttonPanel = new JPanel();
    private JProgressBar progressBar = new JProgressBar(0);
    private DefaultTableModel tableModel = new DefaultTableModel();
    private JFileChooser fileChooser = new JFileChooser();
    private Set<String> setColumnA = new HashSet<>();
    private Set<String> setColumnB = new HashSet<>();
    private JFrame loadingFrame = new JFrame("Carregando...");


    /**
     * Construtor da classe LoadexcelMod01 que inicializa a interface gráfica
     * e configura os elementos visuais, como a tabela e o botão de importação.
     *
     *  *tableModel: DefaultTableModel definido no escopo da classe
     *  *this: container da classe sendo acionado
     * */
    public LoadexcelMod01() {

        tableModel.addColumn("Linha");
        tableModel.addColumn("Coluna A");
        tableModel.addColumn("Coluna B");

        this.setTitle("Resultados da Comparação");
        this.setSize(800, 600);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setLayout(new BorderLayout());

        JButton importButton = new JButton("Importar");
        importButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {importExcel();}
        });

        this.buttonPanel.add(importButton);
        this.add(buttonPanel, BorderLayout.SOUTH);
        this.setVisible(true);
    }

    /** Method: importExcel() Contêiner operacional interno da classe responsável por carregar todos Elementos da tabela
     * processar os campos condicionais separados, para encontrar os elementos Distintos
     *
     * */
    private void importExcel() {
        fileChooser.setDialogTitle("Selecione o arquivo Excel");
        fileChooser.setFileFilter(new FileNameExtensionFilter("Arquivos Excel", "xls", "xlsx"));

        loadingFrame.setSize(350, 120);
        loadingFrame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
        loadingFrame.setLayout(new BorderLayout());
        loadingFrame.add(progressBar, BorderLayout.CENTER);
        loadingFrame.setLocationRelativeTo(null);
        loadingFrame.setVisible(true);
        progressBar.setStringPainted(true);
        progressBar.setOrientation(JProgressBar.HORIZONTAL);


        int userSelection = fileChooser.showOpenDialog(null);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            try {
                FileInputStream fis = new FileInputStream(selectedFile);
                Workbook workbook;
                if (selectedFile.getName().endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                } else if (selectedFile.getName().endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fis);
                } else {
                    throw new IllegalArgumentException("Formato de arquivo não suportado: " + selectedFile.getName());
                }
                Sheet sheet = workbook.getSheetAt(0);
                int columnAIndex = 0; // Coluna A
                int columnBIndex = 1; // Coluna B

                int totalRows = sheet.getLastRowNum(); // Total de linhas na planilha
                progressBar.setMaximum(totalRows); // Define o valor máximo do progressBar

                System.out.println("progressBar: " + totalRows);
                SwingWorker<Void, Integer> barLoad = new SwingWorker<Void, Integer>() {
                    @Override
                    protected Void doInBackground() throws Exception {
                        Iterator<Row> rowIterator = sheet.iterator();
                        rowIterator.next();
                        int rowCount = 1;
                        AtomicInteger finalRowCount= new AtomicInteger();
                        while (rowIterator.hasNext()) {
                            finalRowCount.getAndIncrement();
                            Row row = rowIterator.next();
                            Object[] rowData = {rowCount++, row.getCell(columnAIndex), row.getCell(columnBIndex)};
                            tableModel.addRow(rowData);

                            Cell cellA = row.getCell(columnAIndex);
                            Cell cellB = row.getCell(columnBIndex);

                            if (cellA != null) {
                                setColumnA.add(cellA.toString());
                            }

                            if (cellB != null) {
                                setColumnB.add(cellB.toString());
                            }
                            this.publish(rowCount);
                        }
                       return null;
                    }
                    @Override
                    protected void process(java.util.List<Integer> chunks) {
                        for (int chunk : chunks) {
                            progressBar.setValue(chunk);
                            progressBar.setString("Carregando linha: " + chunk + " de " + progressBar.getMaximum());
                        }
                    }
                    @Override
                    protected void done() {
                        resultTable = new JTable(tableModel);
                       JScrollPane tableScrollPane = new JScrollPane(resultTable);
                       add(tableScrollPane, BorderLayout.CENTER);

                        differencesTextArea = new JTextArea();
                        differencesTextArea.setEditable(false);

                        differencesTextArea.append("\nDiferenças da Coluna A para B:\n");
                        for (String difference : setColumnA) {
                            if (!setColumnB.contains(difference)) {
                                differencesTextArea.append(difference + "\n");
                            }
                        }

                        differencesTextArea.append("\nDiferenças da Coluna B para A:\n");
                        for (String difference : setColumnB) {
                            if (!setColumnA.contains(difference)) {
                                differencesTextArea.append(difference + "\n");
                            }
                        }

                        JScrollPane differencesScrollPane = new JScrollPane( differencesTextArea);
                        add(differencesScrollPane, BorderLayout.SOUTH);

                        TableColumnModel columnModel =  resultTable.getColumnModel();
                        for (int column = 0; column <  resultTable.getColumnCount(); column++) {
                            int width = 15;
                            for (int row = 0; row <  resultTable.getRowCount(); row++) {
                                TableCellRenderer renderer =  resultTable.getCellRenderer(row, column);
                                Component comp =  resultTable.prepareRenderer(renderer, row, column);
                                width = Math.max(comp.getPreferredSize().width + 1, width);
                            }
                            columnModel.getColumn(column).setPreferredWidth(width);
                        }
                        loadingFrame.setVisible(false);
                        loadingFrame.dispose();
                        setSize(800, 620);
                    }
                };
                barLoad.execute();
                System.out.println("Retor do Load --------------");

                this.setVisible(true);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("Nenhum arquivo selecionado. Encerrando o programa.");
        }
    }

}
