using System;
using System.IO;
using NetOffice.ExcelApi;

namespace Clientes.classes
{
    public class Cliente
    {
     public string nome;
     public int idade;
     public DateTime dataNascimento;
     public Contato contato;
     public Endereco endereco;  

     public string cadastrar(){
        Application ex = new Application();
        FileInfo arquivo = new FileInfo(@"c:\Henrique\cliente.xlsx");
        if(arquivo.Exists){
            ex.Visible = true; //Abrir o microsoft excel
            //Vamos abrir o arquivo esxistente
            ex.Workbooks.Open(@"c:\Henrique\cliente.xlsx");
            /*
            Vamos fazer um laço com for para percorrer as linhas do excel e encontrar uma linha vazia.
            Quando encontrar essa linha ele deve parar e escrever a linha com os dados do cliente
             */

             for(int x=3; x <= 100; x++)
             {

                if(ex.Range("A"+x).Value== null )
                {
                   ex.Range("a"+x).Value = nome;
                   ex.Range("b"+x).Value = idade;
                   ex.Range("c"+x).Value = dataNascimento;
                   ex.Range("d"+x).Value = contato.telefone;
                   ex.Range("e"+x).Value = contato.celular;
                   ex.Range("f"+x).Value = contato.email;
                   ex.Range("g"+x).Value = endereco.logradouro;
                   ex.Range("h"+x).Value = endereco.numero;
                   ex.Range("i"+x).Value = endereco.complemento;
                   ex.Range("j"+x).Value = endereco.bairro;
                
                 break; 
                 }
                 //Vamos parar o laço for.

                 
                
                }
                ex.ActiveWorkbook.Save();
                ex.Quit();
             }
             else{
                ex.Visible = true;//abra o excel
                ex.Workbooks.Add();//adicionar uma planilha em branco
                /*
                Vamos montar o cabeçalho da colunas na linha 1

                 */
                 ex.Range("a1").Value= "Nome do Cliente";
                 ex.Range("b1").Value= "Idade";
                 ex.Range("c1").Value= "Data de Nascimento";
                 ex.Range("d1").Value= "Tel. Residencial";
                 ex.Range("e1").Value= "Celular";
                 ex.Range("f1").Value= "Email";
                 ex.Range("g1").Value= "Logradouro";
                 ex.Range("h1").Value= "Número";
                 ex.Range("i1").Value= "Complemento";
                 ex.Range("j1").Value= "Bairro";

             /*
             Vamos formatar aplicando negrito, tamanho de fonte diferente e letra diferente 
             */
             ex.Range("a1:j1").Font.Name="Tahoma";
             ex.Range("a1:j1").Font.Bold= true;
             ex.Range("a1:j1").Font.Size = 15;

            ex.Range("a2").Value = nome;
            ex.Range("b2").Value = idade;
            ex.Range("c2").Value = dataNascimento;
            ex.Range("d2").Value = contato.telefone;
            ex.Range("e2").Value = contato.celular;
            ex.Range("f2").Value = contato.email;
            ex.Range("g2").Value = endereco.logradouro;
            ex.Range("h2").Value = endereco.numero;
            ex.Range("i2").Value = endereco.complemento;
            ex.Range("j2").Value = endereco.bairro;

             //Fazer o auto ajuste    
                        
             ex.ActiveWorkbook.SaveAs(@"c:\Henrique\cliente.xlsx");
             ex.Quit();

            }

 
         
          return "Cliente salvo com sucesso";

        }

     
          public string [,] listar()
          {
              //Vamos construir uma matriz de string para
              //guardar os dados dos clientes 
              string[,] dados = new string[10,10];

              Application excel = new Application();
              excel.Visible =true;
              excel.Workbooks.Open(@"c:\Henrique\cliente.xlsx");
              for(int lin=1; lin<= 10; lin++){
                  for(int col =1; col<= 10; col++){
                     dados[lin-1,col-1] = excel.Cells[lin,col].Text.ToString();

                 
                    }

                 }
                 excel.Quit();

                 return dados;

          



            
       }
    


    }

}