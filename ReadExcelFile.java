package br.com.rfp.services;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;

public class ReadExcelFile {
    public static void main(String[] args) {

        try (FileInputStream file = new FileInputStream("C:\\arquivos\\Base_robo_Botti.xlsx")) {
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheet("histórico de venda");

            // Encontrar índices das colunas desejadas
            int nomeClienteIndex = -1;
            int cidadeIndex = -1;
            int descProdutoIndex = -1;
            int codProdutoIndex = -1;
            int vendedorIndex = -1;
            int gradeIndex = -1;
            int dataEmissaoIndex = -1;

            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                String header = cell.getStringCellValue().trim();
                if (header.equalsIgnoreCase("NOME_CLIENTE")) {
                    nomeClienteIndex = cell.getColumnIndex();
                } else if (header.equalsIgnoreCase("CIDADE")) {
                    cidadeIndex = cell.getColumnIndex();
                } else if (header.equalsIgnoreCase("DESC_PRODUTO")) {
                    descProdutoIndex = cell.getColumnIndex();
                } else if (header.equalsIgnoreCase("COD_PRODUTO")) {
                    codProdutoIndex = cell.getColumnIndex();
                } else if (header.equalsIgnoreCase("Vendedor")) {
                    vendedorIndex = cell.getColumnIndex();
                } else if (header.equalsIgnoreCase("Grade")) {
                    gradeIndex = cell.getColumnIndex();
                } else if (header.equalsIgnoreCase("DATA_EMISSAO")) {
                    dataEmissaoIndex = cell.getColumnIndex();
                }
            }

            // Verificar se todas as colunas foram encontradas
            if (nomeClienteIndex == -1 || cidadeIndex == -1 || descProdutoIndex == -1 || codProdutoIndex == -1 ||
                    vendedorIndex == -1 || gradeIndex == -1 || dataEmissaoIndex == -1) {
                System.out.println("Colunas não encontradas no arquivo Excel.");
            } else {
                System.out.println("Deu bom!!");
            }

            // Criar uma nova planilha filtrada
            Workbook filteredWorkbook = new XSSFWorkbook();
            Sheet filteredSheet = filteredWorkbook.createSheet("PlanilhaFiltrada");

            // Adicionar cabeçalho à nova planilha
            Row headerRowFiltered = filteredSheet.createRow(0);
            headerRowFiltered.createCell(0).setCellValue("Nome do Cliente");
            headerRowFiltered.createCell(1).setCellValue("Cidade");
            headerRowFiltered.createCell(2).setCellValue("Descrição do Produto");
            headerRowFiltered.createCell(3).setCellValue("Código Produto");
            headerRowFiltered.createCell(4).setCellValue("Vendedor");
            headerRowFiltered.createCell(5).setCellValue("Tamanho");
            headerRowFiltered.createCell(6).setCellValue("Data da última compra");
            headerRowFiltered.createCell(7).setCellValue("Codigo Produto Botti");


            int filteredRowIndex = 1; // Inicia a partir da segunda linha na nova planilha

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                String nomeCliente = row.getCell(nomeClienteIndex).getStringCellValue();
                String cidade = row.getCell(cidadeIndex).getStringCellValue();
                String descProduto = row.getCell(descProdutoIndex).getStringCellValue();
                String codProduto = getCellValueAsString(row.getCell(codProdutoIndex));
                String vendedor = getCellValueAsString(row.getCell(vendedorIndex), "Vendedor não informado");
                String codProdutoBotti = newCell();

                Cell grade = row.getCell(gradeIndex);
                // Avaliar a fórmula da célula de grade
                String tamanho = evaluateFormula(grade, sheet.getWorkbook().getCreationHelper().createFormulaEvaluator());

                SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
                formato.setLenient(false);

                Date dataEmissao =  row.getCell(dataEmissaoIndex).getDateCellValue();
                String dateFormat = formato.format(dataEmissao);
                Date date = formato.parse(dateFormat);
                System.out.println(date);

                Date dataAtualSubtract = subtractDate();
                System.out.println(dataAtualSubtract);

                if (date.before(dataAtualSubtract)) {
                    // Adicionar dados à nova planilha filtrada
                    String novaDataFormatada = formato.format(date);
                    Row filteredRow = filteredSheet.createRow(filteredRowIndex);
                    filteredRow.createCell(0).setCellValue(nomeCliente);
                    filteredRow.createCell(1).setCellValue(cidade);
                    filteredRow.createCell(2).setCellValue(descProduto);
                    filteredRow.createCell(3).setCellValue(codProduto);
                    filteredRow.createCell(4).setCellValue(vendedor);
                    filteredRow.createCell(5).setCellValue(tamanho);
                    filteredRow.createCell(6).setCellValue(novaDataFormatada);
                    filteredRow.createCell(7).setCellValue(codProdutoBotti);

                    filteredRowIndex++;
                    System.out.println("-----------------cliente add -------------");
                    System.out.println("Data da última compra: " + dataEmissao);
                    System.out.println("------------------------------");

                } else {
                    System.out.println("-----------------cliente falhou-------------");
                    System.out.println("Data da última compra: " + dataEmissao);
                    System.out.println("------------------------------");
                }
            }

            // Salvar a nova planilha filtrada
            try (FileOutputStream filteredFile = new FileOutputStream("C:\\arquivos\\PlanilhaFiltrada1.xlsx")) {
                filteredWorkbook.write(filteredFile);
                filteredWorkbook.close();
                System.out.println("Nova planilha filtrada criada com sucesso.");
            } catch (IOException e) {
                e.printStackTrace();
            }

            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }
    }

    private static String getCellValueAsString(Cell cell, String defaultValue) {
        if (cell == null) {
            return defaultValue;
        }

        if (cell.getCellType() == CellType.FORMULA) {
            // Avalia a fórmula e obtém o resultado como um valor numérico
            double numericValue = cell.getNumericCellValue();
            return String.valueOf(numericValue);

        } else if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            // Se a célula contiver uma data numérica, converte para LocalDate
            LocalDate dataEmissaoDate = cell.getLocalDateTimeCellValue().toLocalDate();
            return dataEmissaoDate.format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));

        } else {
            // Se não, obtém como uma string
            return cell.getStringCellValue();
        }
    }

    private static String evaluateFormula(Cell cell, FormulaEvaluator evaluator) {
        if (cell != null && cell.getCellType() == CellType.FORMULA) {
            try {
                CellValue cellValue = evaluator.evaluate(cell);
                if (cellValue.getCellType() == CellType.NUMERIC) {
                    return String.valueOf(cellValue.getNumberValue());

                } else if (cellValue.getCellType() == CellType.STRING) {
                    return cellValue.getStringValue();

                } else if (cellValue.getCellType() == CellType.BOOLEAN) {
                    return String.valueOf(cellValue.getBooleanValue());

                } else if (cellValue.getCellType() == CellType.ERROR) {
                    // Trata o caso em que a célula contém um valor de erro
                    return "Erro na célula: " + FormulaError.forInt(cellValue.getErrorValue()).getString();
                }

            } catch (Exception e) {
                // Trata o caso em que a fórmula não pode ser avaliada
                return "Erro ao avaliar a fórmula";
            }
        }
        return getCellValueAsString(cell, "");
    }

    // Sobrecarga do método para tratar casos em que o valor padrão é uma string
    private static String getCellValueAsString(Cell cell) {
        return getCellValueAsString(cell, "");
    }

    private static Date subtractDate() throws ParseException {
        SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
        formato.setLenient(false);

        String dataAtual = "22/12/2023";
        Date dataAtualDate = formato.parse(dataAtual);

        // Criar um objeto Calendar e definir a data atual
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(dataAtualDate);

        // Subtrair 30 dias
        calendar.add(Calendar.DAY_OF_MONTH, -30);

        // Obter a nova data após a subtração
        Date novaData = calendar.getTime();

        // Formatar a nova data antes de retorná-la
        String novaDataFormatada = formato.format(novaData);

        // Retorna a nova data já formatada
        return novaData;
    }

    private static String newCell() {
        String codigoProduto2 = null;
        try (FileInputStream arquivoExcel = new FileInputStream("C:\\arquivos\\Botti.xlsx")) {
            Workbook workbook = WorkbookFactory.create(arquivoExcel);
            Sheet sheet = workbook.getSheetAt(0); // Pode variar dependendo do índice da folha

            Row row = sheet.getRow(1); // Segunda linha (índice 1)
            Cell cell = row.getCell(2); // Terceira coluna (índice 2)

            if (cell != null) {
                codigoProduto2 = cell.getStringCellValue();
                System.out.println("Valor da célula C2: " + codigoProduto2);
            } else {
                System.out.println("Célula C2 está vazia.");
            }

            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return codigoProduto2;
    }

}
