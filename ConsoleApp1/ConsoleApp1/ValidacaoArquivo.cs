using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class ValidacaoArquivoErro
    {
        public string MensagemErro { get; set; }
        public int Linha { get; set; }
    }

    public class ValidacaoArquivo
    {
        DateTime DataValidacao { get; set; } = DateTime.Now;
        public List<ValidacaoArquivoErro> Erros { get; set; }

        public bool ArquivoValido()
        {
            return Erros.Count() == 0;
        }

        public void AddErro(string Mensagem)
        {
            Erros.Add(new ValidacaoArquivoErro()
            {
                MensagemErro = Mensagem,
                Linha = 0
            });
        }

        public void AddErro(string Mensagem, int Linha)
        {
            Erros.Add(new ValidacaoArquivoErro()
            {
                MensagemErro = Mensagem,
                Linha = Linha
            });
        }

        public ValidacaoArquivo()
        {
            Erros = new List<ValidacaoArquivoErro>();
        }
    }
}
