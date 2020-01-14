using System;
using OpenQA.Selenium.Chrome;
using System.IO;
using System.Threading;
using System.Configuration;
using System.Linq;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using ConsoleApp1;

namespace Automacao_Recuperacao_Arquivo
{
    class Program
    {
        static void Main(string[] args)
        {
            string url = ConfigurationManager.AppSettings["path_print_laser_principal"];
            string url_upload = ConfigurationManager.AppSettings["path_print_laser_upload"];
            string url_main = ConfigurationManager.AppSettings["path_print_laser_main"] 
            string user = ConfigurationManager.AppSettings["keyUsuario"];
            string password = ConfigurationManager.AppSettings["keySenha"];
            string path_file;
            string grupo = ConfigurationManager.AppSettings["grupo"];
            string cliente = ConfigurationManager.AppSettings["cliente"];
            string unidade = ConfigurationManager.AppSettings["unidade"];
            string aplicacao = ConfigurationManager.AppSettings["aplicacao"];

            ChromeDriver drv;
            ValidacaoArquivo val = new ValidacaoArquivo();
            drv = new ChromeDriver();
            try
            {
                // INICIO
                //Antes de iniciar, codificar a conexão com a web service para verificar
                //se tem item a processar, caso não tenha....nem precesia iniciar os deamis 
                //itens do processo
                

                RealizarLogin(drv, url, user, password);
                Thread.Sleep(1000);

                ClicarSuporte(drv);
                Thread.Sleep(200);

                //Captura o endereço do arquivo
                path_file = @"C:\Boleto\RecuperacaoEletronica_20180413ka.csv";

                //Passo uma válidação no arquivo 
                Validacao_File(path_file, ref val);

                //Verifico se a validação caso o arquivo esteja OK, continuo, senão cancelo a demanda e vou para a próxima
                if (!val.ArquivoValido())
                {
                    Console.WriteLine("Arquivo não válido!!");
                }

                Thread.Sleep(200);
                ClicarUpload(drv);

                Thread.Sleep(200);
                Aguardar_Tela(drv, url_upload);

                Thread.Sleep(200);
                Selecionar_Grupo(drv, grupo, cliente, unidade, aplicacao, path_file);

                Thread.Sleep(200);
                ClicarEnviar(drv);
                Console.WriteLine("OK");

                Thread.Sleep(500);
                FechaTela(drv, url_upload, url_main);
                    
                drv.Quit();
                // FIM

            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        public static void Validacao_File(string path_file, ref ValidacaoArquivo val)
        {
            try
            {
                string linha = null;
                string[] Header = {"CONTA", "MÊS", "ANO", "CICLO"};
                string[] Formatos = {"FEBRABANV2", "FEBRABANV3", "VOL"};

                if (File.Exists(@path_file))
                {
                    StreamReader reader = new StreamReader(@path_file);
                    Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
                    linha = reader.ReadLine();
                    string[] linha_separada = linha.Split(';');
                    if (linha_separada.Count() == 5)
                    {
                        if (linha_separada[0].ToUpper().Trim() != "CONTA")
                        {
                            val.AddErro("Verificar o cabeçalho do arquivo o nome correto da 1º coluna é 'CONTA'", 1);
                        }

                        if (linha_separada[1].ToUpper().Replace("Ê","E").Replace("�", "E").Trim () != "MES")
                        {

                            val.AddErro("Verificar o cabeçalho do arquivo o nome correto da 2º coluna é 'MES'", 1);
                        }

                        if (linha_separada[2].ToUpper() != "ANO")
                        {
                           
                            val.AddErro("Verificar o cabeçalho do arquivo o nome correto da 3º coluna é 'ANO'", 1);
                        }

                        if (linha_separada[3].ToUpper() != "CICLO")
                        {
                           
                            val.AddErro("Verificar o cabeçalho do arquivo o nome correto da 4º coluna é 'CICLO'", 1);
                        }

                        if (linha_separada[4].ToUpper() != "FORMATO")
                        {
                           
                            val.AddErro("Verificar o cabeçalho do arquivo o nome correto da 4º coluna é 'FORMATO', preenchido com umas das opçoes ('FebrabanV2', FebrabanV3 ou 'VOL'", 1);
                        }

                        linha_separada[0] = null;
                        
                        while ((linha = reader.ReadLine()) !=null)
                        {

                            // array das celulas
                            linha_separada = linha.Split(';');

                            // validar so os inteiros
                            for (int cel = 0; cel < 3; cel++)
                            {
                                int teste = 0;

                                if (!int.TryParse(linha_separada[cel], out teste))
                                {
                                   
                                    val.AddErro(String.Format("Coluna {0} com formato inválido. Deve ser um número.", Header[cel]));
                                }

                            }

                            // Validar o texto
                            if (!String.IsNullOrEmpty(linha_separada[4]))
                            {
                                if (!(Formatos.Contains(linha_separada[4].ToUpper())))
                                {
                                   
                                    val.AddErro("Item da coluna FORMATO inválida deve ser 'FEBRABANV2, 'FEBRABANV3, 'VOL");
                                }

                            }
                            else
                            {
                                val.AddErro("Coluna FORMATO não pode ser vazio.");
                            }
                        }
                        reader.Close();

                    }
                    else
                    {
                      
                        val.AddErro("O arquivo contem " + linha_separada.Count() + " coluna, e o correto é ter 5 colunas, na seguinte ordem 'CONTA, MES, ANO, CICLO, FORMATO'");
                    }
                   


                }


            }
            catch (Exception ex)
            {

                throw;
            }


        }

        public static void RealizarLogin(ChromeDriver drv, string url, string user, string password)
        {
            //
            drv.Navigate().GoToUrl(url);
            drv.FindElementById("ASPxCallbackPanel1_pnlLogin_txtUsuario_I").SendKeys(user);
            drv.FindElementById("ASPxCallbackPanel1_pnlLogin_txtSenha_I").SendKeys(password);
            drv.FindElementById("ASPxCallbackPanel1_pnlLogin_cmdConfirmar_CD").Click();
            Thread.Sleep(2000);

            try
            {
                new WebDriverWait(drv, TimeSpan.FromSeconds(20))
                    .Until(ExpectedConditions.ElementExists(By.Id("ASPxCallbackPanel1_pnlAcesso_cmdAcesso_CD")))
                    .Click();
            }
            catch (WebDriverException ex)
            {

                Environment.Exit(0);
            }

            foreach (var aba in drv.WindowHandles)
            {
                drv.SwitchTo().Window(aba);
                if (!drv.Url.Equals("https://portal.printlaser.com/Main.aspx"))
                    drv.Close();
            }
        }

        public static void ClicarSuporte(ChromeDriver drv)
        {
            //
            try
            {
                new WebDriverWait(drv, TimeSpan.FromSeconds(20))
                    .Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.TagName("span")
                     ))
                     .Where(p => p.Text.ToLower().Equals("suporte"))
                     .First()
                     .Click();

            }
            catch (Exception ex)
            {

                Environment.Exit(1);
            }
        }

        public static void ClicarUpload(ChromeDriver drv)
        {
            //
            try
            {
                new WebDriverWait(drv, TimeSpan.FromSeconds(20))
                    .Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.TagName("span")))
                    .Where(p => p.Text.ToLower().Equals("upload de arquivos"))
                    .First()
                    .Click();

            }
            catch (Exception ex)
            {
                Environment.Exit(2);
            }
        }

        public static void Aguardar_Tela(ChromeDriver drv, string url_upload)
        {
            try
            {
                int i = 0;
                for (i = 0; i <= 10; i++)
                {
                    foreach (var aba in drv.WindowHandles)
                    {
                        drv.SwitchTo().Window(aba);
                        if (drv.Url.Equals(url_upload))
                        {
                            Thread.Sleep(1000);
                            return;
                        }
                    }
                    Console.WriteLine("Valor de i:", i);
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                Environment.Exit(3);
            }
        }

        public static void Selecionar_Grupo(ChromeDriver drv, string grupo, string cliente, string unidade, string aplicacao, string path_file)
        {
            try
            {
                //Grupo
                drv.FindElement(By.XPath(".//*[@id='acpMain_arpMain_acbEnvio_arpDestino_cboGrupos_B-1']")).Click();
                Thread.Sleep(500);
                drv.FindElement(By.Id("acpMain_arpMain_acbEnvio_arpDestino_cboGrupos_I")).SendKeys(grupo);
                Thread.Sleep(500);
                drv.FindElement(By.XPath(".//*[@id='acpMain_arpMain_acbEnvio_arpDestino_cboGrupos_B-1']")).Click();
                Thread.Sleep(500);
              
                //Cliente
                drv.FindElement(By.XPath(".//*[@id='acpMain_arpMain_acbEnvio_arpDestino_cboClientes_B-1']")).Click();
                Thread.Sleep(500);
                drv.FindElement(By.Id("acpMain_arpMain_acbEnvio_arpDestino_cboClientes_I")).SendKeys(cliente);
                Thread.Sleep(500);
                drv.FindElement(By.XPath(".//*[@id='acpMain_arpMain_acbEnvio_arpDestino_cboClientes_B-1']")).Click();
                Thread.Sleep(500);

                //Unidade
                drv.FindElement(By.XPath(".//*[@id='acpMain_arpMain_acbEnvio_arpDestino_cboUnidades_B-1']")).Click();
                Thread.Sleep(500);
                drv.FindElement(By.Id("acpMain_arpMain_acbEnvio_arpDestino_cboUnidades_I")).SendKeys(unidade);
                Thread.Sleep(500);
                drv.FindElement(By.XPath(".//*[@id='acpMain_arpMain_acbEnvio_arpDestino_cboUnidades_B-1']")).Click();

                //Aplicação
                drv.FindElement(By.XPath(".//*[@id='acpMain_arpMain_acbEnvio_arpDestino_cboAplicacoes_B-1']")).Click();
                Thread.Sleep(500);
                drv.FindElement(By.Id("acpMain_arpMain_acbEnvio_arpDestino_cboAplicacoes_I")).SendKeys(aplicacao);
                Thread.Sleep(500);
                drv.FindElement(By.XPath(".//*[@id='acpMain_arpMain_acbEnvio_arpDestino_cboAplicacoes_B-1']")).Click();

                //Passar o arquivo
                IWebElement elem = drv.FindElement (By.XPath ("//input[@type='file']"));
                Thread.Sleep(500);
                elem.SendKeys(@path_file);

            }
            catch (Exception ex)
            {
                Environment.Exit(4);
            }
        }

        public static void ClicarEnviar(ChromeDriver drv)
        {
            //
            try
            {
                new WebDriverWait(drv, TimeSpan.FromSeconds(20))
                    .Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.TagName("span")
                     ))
                     .Where(p => p.Text.ToLower().Equals("enviar"))
                     .First()
                     .Click();
                Thread.Sleep(500);
                IAlert a = drv.SwitchTo().Alert();
                if (a.Text.Equals("Confirma o envio dos arquivos?"))
                {
                    a.Accept();
                }

               
                

            }
            catch (Exception ex)
            {

                Environment.Exit(1);
            }
        }


        public static void FechaTela(ChromeDriver drv, string url_fechar, string url_setar)
        {
            //
          
            Thread.Sleep(1000);

            foreach (var aba in drv.WindowHandles)
            {
                drv.SwitchTo().Window(aba);
                if (drv.Url.Equals(url_fechar))
                    drv.Close();
            }

            foreach (var aba in drv.WindowHandles)
            {
                drv.SwitchTo().Window(aba);
                if (drv.Url.Equals("https://portal.printlaser.com/Main.aspx"))
                {
                    return;
                }
            }
        }
    }
}
