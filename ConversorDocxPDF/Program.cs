using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Spire.Doc;

namespace ConversorDocxPDF {
    internal class Program {
        static void Main(string[] args) {

            Document documento = new Document();

            try {

                documento.LoadFromFile(@"D:\Juliano Martins\Documentos\Trabalho\Curriculo.Docx");

                if (documento == null) {
                    throw new Exception("Erro na aplicação. Arquivo não foi carregado");
                }

                documento.SaveToFile(@"D:\Downloads", FileFormat.PDF);

                Console.WriteLine("Conversão concluída com sucesso");

                documento.Dispose();


            }
            catch (Exception e) {
                documento.Dispose();
                Console.Write("Erro na aplicação " + e.Message);
            }
        }
    }
}
