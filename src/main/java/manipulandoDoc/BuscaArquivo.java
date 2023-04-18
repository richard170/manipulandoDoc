package manipulandoDoc;

public class BuscaArquivo {

    public BuscaArquivo() {
    }

    private static String caminho;
    private static int planilhas;

    public void Busca(){
        if (planilhas==0){
            caminho = "C:\\Users\\evolu\\Desktop\\manipulandoDoc\\test.xlsx";
            planilhas++;
        } else if (planilhas == 1){
            caminho = "C:\\Users\\evolu\\Desktop\\manipulandoDoc\\test2.xlsx";
            planilhas++;
        }else if (planilhas == 2){
            caminho = "C:\\Users\\evolu\\Desktop\\manipulandoDoc\\test3.xlsx";
            planilhas++;
        }else if (planilhas == 3){
            caminho = "C:\\Users\\evolu\\Desktop\\manipulandoDoc\\test4.xlsx";
            planilhas++;
        }else {
            planilhas = 0;
        }
    }

    public String getCaminho() {return caminho;}

    public void setCaminho(String caminho) {this.caminho = caminho;}

}
