using System;
using NetOffice.ExcelApi;
using Clientes.classes;

namespace Clientes
{
    class Program
    {
        static void Main(string[] args)
        {
            //Vamos instanciar a classe cliente para
            //realizar o cadastro dos clientes
                Cliente cli = new Cliente();
                //Vamos instanciar a classe contato
                Contato ct = new Contato();
                ct.telefone = "11-788878";
                ct.celular = "11-935421887";
                ct.email = "Ferzin@terra.com";

                //Vamos instanciar a calasse endereço
                Endereco end = new Endereco();
                end.logradouro = "lugar";
                end.numero = "D12";
                end.complemento = "D12";
                end.bairro = "jardim ania";


                cli.nome="Ferzin";
                cli.idade=23;
                cli.dataNascimento= DateTime.Parse("26/09/1992");
                cli.contato = ct;
                cli.endereco = end;

               // Console.WriteLine(cli.cadastrar());

               //passando todos os dados da matriz listar que está
               //carregando com os dados dos clientes para um novo 
               //array. Assim esses dados serão carregados de uma
               //só vez.
               string[,] info =cli.listar();

                for(int i = 0 ; i <10; i++){
                    for(int p = 0 ; p <10; p++){
                        Console.Write(info[i,p]+ "\t");
                    }
                    Console.WriteLine();
                }

        }
    }
}
