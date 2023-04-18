package manipulandoDoc;

import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PrincipalDoc {


    public static void main(String[] args) throws IOException {

        BuscaArquivo busca = new BuscaArquivo();
        Aluno aluno = new Aluno();

        List<String> listaNomes1 = new ArrayList<>();
        List<String> listaNomes2 = new ArrayList<>();
        List<String> listaNomes3 = new ArrayList<>();
        List<String> listaNomes4 = new ArrayList<>();

        List<Integer> listaNotas1 = new ArrayList<>();
        List<Integer> listaNotas2 = new ArrayList<>();
        List<Integer> listaNotas3 = new ArrayList<>();
        List<Integer> listaNotas4 = new ArrayList<>();

        List<Integer> listaFinalNotas = new ArrayList<>();
        ArrayList<String> listaFinalNomes = new ArrayList<>();


        ArrayList<String> todosAlunosTodasPlan = new ArrayList<>();
        ArrayList<Integer> todasNotasTodasPlan = new ArrayList<>();
        int abas = 0;
        while (abas <= 7) {

            int fimPrograma = 0;
            int planilhaNum = 1;
            int cont = 0;
            int valid = 0;
            int numeroDeAlunos = 0;
            while (fimPrograma == 0) {
                String caminho;
                busca.Busca();
                caminho = busca.getCaminho();
                try {
                    FileInputStream arquivo = new FileInputStream(caminho);

                    XSSFWorkbook workbook = new XSSFWorkbook(arquivo);

                    //pega a primeira aba da planilha, por isso o valor igual a 0, outras abas valor 1.. 2.. por diante
                    XSSFSheet sheetAlunos = workbook.getSheetAt(abas);

                    //retorna todas as linhas da planilha 0 ou aba 1
                    Iterator<Row> rowIterator = sheetAlunos.iterator();

                    //criar um objeto da classe FormulaEvaluator que será usado para avaliar fórmulas em uma planilha do Excel
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

                    //varre todas as linhas da planilha 0 ou aba 1
                    while (rowIterator.hasNext()) {

                        //recebe cada linha da planilha
                        Row row = rowIterator.next();

                        //pega todas as celulas desssa linha
                        Iterator<Cell> cellIterator = row.cellIterator();

                        //varrendo todas as celulas da linha atual
                        while (cellIterator.hasNext()) {
                            Cell cell = cellIterator.next();
                            if (planilhaNum == 1) {
                                if (cell.getColumnIndex() == 1) {
                                    if (cell.getRowIndex() >= 60 && cell.getRowIndex() < 101) {
                                        if (evaluator.evaluateFormulaCell(cell) == CellType.STRING && evaluator.evaluateFormulaCell(cell) != null) {
                                            aluno.setNome(cell.getStringCellValue());
                                            if (aluno.getNome() != null && !aluno.getNome().equals("")) {
                                                listaNomes1.add(cell.getStringCellValue());
                                                Row linha = sheetAlunos.getRow(cell.getRowIndex());
                                                Cell celula = linha.getCell(9);
                                                if (evaluator.evaluateFormulaCell(celula) == CellType.NUMERIC && evaluator.evaluateFormulaCell(celula) != null) {
                                                    listaNotas1.add((int) Math.round(celula.getNumericCellValue()));
                                                } else {
                                                    listaNotas1.add(0);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (planilhaNum == 2) {
                                if (cell.getColumnIndex() == 1) {
                                    if (cell.getRowIndex() >= 60 && cell.getRowIndex() < 101) {
                                        if (evaluator.evaluateFormulaCell(cell) == CellType.STRING && evaluator.evaluateFormulaCell(cell) != null) {
                                            aluno.setNome(cell.getStringCellValue());
                                            if (aluno.getNome() != null && !aluno.getNome().equals("")) {
                                                listaNomes2.add(cell.getStringCellValue());
                                                Row linha = sheetAlunos.getRow(cell.getRowIndex());
                                                Cell celula = linha.getCell(9);
                                                if (evaluator.evaluateFormulaCell(celula) == CellType.NUMERIC && evaluator.evaluateFormulaCell(celula) != null) {
                                                    listaNotas2.add((int) Math.round(celula.getNumericCellValue()));
                                                } else {
                                                    listaNotas2.add(0);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (planilhaNum == 3) {
                                if (cell.getColumnIndex() == 1) {
                                    if (cell.getRowIndex() >= 60 && cell.getRowIndex() < 101) {
                                        if (evaluator.evaluateFormulaCell(cell) == CellType.STRING && evaluator.evaluateFormulaCell(cell) != null) {
                                            aluno.setNome(cell.getStringCellValue());
                                            if (aluno.getNome() != null && !aluno.getNome().equals("")) {
                                                listaNomes3.add(cell.getStringCellValue());
                                                Row linha = sheetAlunos.getRow(cell.getRowIndex());
                                                Cell celula = linha.getCell(10);
                                                if (evaluator.evaluateFormulaCell(celula) == CellType.NUMERIC && evaluator.evaluateFormulaCell(celula) != null) {
                                                    listaNotas3.add((int) Math.round(celula.getNumericCellValue()));
                                                } else {
                                                    listaNotas3.add(0);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (planilhaNum == 4) {
                                if (cell.getColumnIndex() == 1) {
                                    if (cell.getRowIndex() >= 60 && cell.getRowIndex() < 101) {
                                        if (evaluator.evaluateFormulaCell(cell) == CellType.STRING && evaluator.evaluateFormulaCell(cell) != null) {
                                            aluno.setNome(cell.getStringCellValue());
                                            if (aluno.getNome() != null && !aluno.getNome().equals("")) {
                                                listaNomes4.add(cell.getStringCellValue());
                                                Row linha = sheetAlunos.getRow(cell.getRowIndex());
                                                Cell celula = linha.getCell(11);
                                                if (evaluator.evaluateFormulaCell(celula) == CellType.NUMERIC && evaluator.evaluateFormulaCell(celula) != null) {
                                                    listaNotas4.add((int) Math.round(celula.getNumericCellValue()));
                                                } else {
                                                    listaNotas4.add(0);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (cont <= 3) {
                        cont++;
                        planilhaNum++;
                    } else {
                        fimPrograma++;
                    }
                    arquivo.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                    System.out.println("Arquivo Excel não encontrado!");
                }
            }
            if (listaNomes1.size() == 0) {
                System.out.println("Nenhum aluno encontrado!");
            } else {

                HashSet<String> todosNomes = new HashSet();
                todosNomes.addAll(listaNomes1);
                todosNomes.addAll(listaNomes2);
                todosNomes.addAll(listaNomes3);
                todosNomes.addAll(listaNomes4);
                listaFinalNomes.addAll(todosNomes);
                Collections.sort(listaFinalNomes);

                for (int i = 0; i < listaFinalNomes.size(); i++) {
                    if (listaNomes1.size() < listaFinalNomes.size()) {
                        listaNomes1.add("");
                    }
                    if (listaNomes2.size() < listaFinalNomes.size()) {
                        listaNomes2.add("");
                    }
                    if (listaNomes3.size() < listaFinalNomes.size()) {
                        listaNomes3.add("");
                    }
                    if (listaNomes4.size() <= listaFinalNomes.size()) {
                        listaNomes4.add("");
                    }
                }
                for (int i = 0; i < listaFinalNomes.size(); i++) {
                    listaFinalNotas.add(0);
                }
                for (int i = 0; i < listaFinalNomes.size(); i++) {

                    if (listaFinalNomes.get(i).equals(listaNomes1.get(i))) {
                        listaFinalNotas.add(i, listaNotas1.get(i));
                        continue;
                    } else {
                        for (int j = 0; j < listaFinalNomes.size(); j++) {
                            if (listaFinalNomes.get(i).equals(listaNomes1.get(j))) {
                                listaFinalNotas.add(i, listaNotas1.get(j));
                                valid++;
                            }
                        }
                    }
                    if (valid == 0) {
                        listaFinalNotas.add(0);
                    }
                    valid = 0;
                }
                for (int i = 0; i < listaFinalNomes.size(); i++) {
                    if (listaFinalNomes.get(i).equals(listaNomes2.get(i))) {
                        listaFinalNotas.set(i, listaNotas2.get(i) + listaFinalNotas.get(i));
                        continue;
                    } else {
                        for (int j = 0; j < listaFinalNomes.size(); j++) {
                            if (listaFinalNomes.get(i).equals(listaNomes2.get(j))) {
                                listaFinalNotas.set(i, listaNotas2.get(j) + listaFinalNotas.get(i));
                                valid++;
                            }
                        }
                    }
                    if (valid == 0) {
                        listaFinalNotas.add(listaFinalNotas.get(i));
                    }
                    valid = 0;
                }
                for (int i = 0; i < listaFinalNomes.size(); i++) {
                    if (listaFinalNomes.get(i).equals(listaNomes3.get(i))) {
                        listaFinalNotas.set(i, listaNotas3.get(i) + listaFinalNotas.get(i));
                        continue;
                    } else {
                        for (int j = 0; j < listaFinalNomes.size(); j++) {
                            if (listaFinalNomes.get(i).equals(listaNomes3.get(j))) {
                                listaFinalNotas.set(i, listaNotas3.get(j) + listaFinalNotas.get(i));
                                valid++;
                            }
                        }
                    }
                    if (valid == 0) {
                        listaFinalNotas.add(listaFinalNotas.get(i));
                    }
                    valid = 0;
                }
                for (int i = 0; i < listaFinalNomes.size(); i++) {
                    if (listaFinalNomes.get(i).equals(listaNomes4.get(i))) {
                        listaFinalNotas.set(i, listaNotas4.get(i) + listaFinalNotas.get(i));
                        continue;
                    } else {
                        for (int j = 0; j < listaFinalNomes.size(); j++) {
                            if (listaFinalNomes.get(i).equals(listaNomes4.get(j))) {
                                listaFinalNotas.set(i, listaNotas4.get(j) + listaFinalNotas.get(i));
                                valid++;
                            }
                        }
                    }
                    if (valid == 0) {
                        listaFinalNotas.add(listaFinalNotas.get(i));
                    }
                    valid = 0;
                }
                for (int i = 0; i < listaFinalNomes.size(); i++) {
                    System.out.print(listaFinalNomes.get(i));
                    System.out.print("  Notas: " + listaFinalNotas.get(i) + "\n");
                    numeroDeAlunos++;
                }
                System.out.println("\nNumero de alunos que respondeu ao exame: " + numeroDeAlunos+"\n\n");
            }
            todosAlunosTodasPlan.addAll(listaFinalNomes);
            todasNotasTodasPlan.addAll(listaFinalNotas);
            listaFinalNomes.clear();
            listaFinalNotas.clear();

            listaNomes1.clear();
            listaNomes2.clear();
            listaNomes3.clear();
            listaNomes4.clear();

            listaNotas1.clear();
            listaNotas2.clear();
            listaNotas3.clear();
            listaNotas4.clear();
            abas++;

        }
        System.out.println(";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;");
        System.out.println(";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;");
        System.out.println(";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;");
        System.out.println(";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;");
        System.out.println(";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;");
        System.out.println(";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;");
        System.out.println("\n\n\n\n");
        for (int i = 0; i < todosAlunosTodasPlan.size(); i++) {
            System.out.print("Nome: "+todosAlunosTodasPlan.get(i));
            System.out.print(" Notas: "+todasNotasTodasPlan.get(i)+"\n");
        }
    }
}
