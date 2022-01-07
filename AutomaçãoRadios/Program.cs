// RsInstrument Exemplos para Analisadores de Espectro
// Procedimentos:
// - Instalar Rohde & Schwarz VISA 5.12.3+ pelo link https://www.rohde-schwarz.com/appnote/1dc02


using System;
using System.Threading;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using RohdeSchwarz.RsInstrument; // Biblioteca que providencia os comandos. Procure ela no www.nuget.org


namespace RsInstrument_FSW_Example
{
    
    class Radios
    {
        public void CriaPasta(string nomeArquivo, string nomePasta)
        {
            //---------------------------------------------------------------
            //Cria Uma pasta para salvar os valores
            //---------------------------------------------------------------

            // Cria uma subPasta na pasta teste Automação
            // Adicionando o nome Valores do Marker dentro da variavel Nome Pasta
            nomePasta = System.IO.Path.Combine(nomeArquivo, nomePasta);

            //Cria pasta
            System.IO.Directory.CreateDirectory(nomePasta);

            // Verifica o Caminho
        }

        public void SalvaValores(string nomeArquivo, string nomePasta, string valor, string freqC)
        {
            if (!System.IO.File.Exists(nomePasta))
            {
                // Combina o nome do arquivo ao caminho onde ta os prints
                CriaPasta(nomeArquivo, nomePasta);
                nomePasta = System.IO.Path.Combine(nomePasta, nomeArquivo);


                //Criando o arquivo e adicionando os Valores
                Console.WriteLine("Criando o arquivo \"{0}\" e adicionando os valores", nomeArquivo);
                File.AppendAllText(nomePasta, freqC.ToString() + ";");
                File.AppendAllText(nomePasta, valor.ToString() + ";");
            }
            else
            {
                Console.WriteLine("O arquivo \"{0}\" Ja existe. Apenas inserindo os valores", nomeArquivo);
                File.AppendAllText(nomePasta, freqC.ToString() + ";");
                File.AppendAllText(nomePasta, valor.ToString() + ";");
            }
        }
        void EstabilidadeFreq(RsInstrument instr, string freqC, string nomeArquivo, string nomePasta)
        {
            for (int i = 0; i < 4; i++)
            {

                //instr.WriteString("*RST;*CLS"); // Reinicialize o instrumento, limpe a fila de erros
                instr.WriteString("INST:SEL MREC"); //Configura o Analisador para o Spectrum mode
                instr.WriteString("INIT:CONT ON"); // Desliga a varredura contínua
                instr.WriteString("SYST:DISP:UPD ON"); // Display update ON - Desligar após a depuração

                instr.WriteString($"FREQ:CENT {freqC} MHz"); // Configura a Frequencia Central
                instr.WriteString("SWE:TIME:AUTO ON"); // Configura o sweep points
                instr.QueryOpc(); //Usando * OPC? consulta espera até que todas as configurações do instrumento sejam concluídas
                // -----------------------------------------------------------
                // SyncPoint 'SettingsApplied' - todas as configurações foram aplicadas
                // -----------------------------------------------------------
                instr.VisaTimeout = 2000; // tempo limite de varredura - defina ele mais alto do que o tempo de aquisição do instrumento
                instr.WriteString("INIT"); // Comece a varredura
                instr.QueryOpc(); // Usando * OPC? consulta espera até que o instrumento termine a aquisição
                string estabilidadeFreq = instr.QueryString("CALC:MARK:FUNC:ADEM:FERR?");
                Thread.Sleep(5000);
                Console.WriteLine(estabilidadeFreq);
                SalvaValores(nomeArquivo, nomePasta, estabilidadeFreq, freqC);
            }
        }

        void Potencia(RsInstrument instr, string freqC, string nomePasta, string nomeSubPasta)
        {
            //---------------------------------------------------------------------
            //Potencia
            //---------------------------------------------------------------------
            instr.WriteString("*RST;*CLS"); // Reinicialize o instrumento, limpe a fila de erros
            instr.WriteString("INIT:CONT ON"); // Desliga a varredura contínua
            instr.WriteString("SYST:DISP:UPD ON"); // Display update ON - Desligar após a depuração

            instr.WriteString("INST:SEL MREC"); //Configura o Analisador para o Spectrum mode
            instr.WriteString($"FREQ:CENT {freqC} MHz"); // Configura a Frequencia Central
            instr.WriteString("SWE:TIME:AUTO ON"); // Configura o sweep points
            instr.QueryOpc(); //Usando * OPC? consulta espera até que todas as configurações do instrumento sejam concluídas
            // -----------------------------------------------------------
            // SyncPoint 'SettingsApplied' - todas as configurações foram aplicadas
            // -----------------------------------------------------------
            instr.VisaTimeout = 2000; // tempo limite de varredura - defina ele mais alto do que o tempo de aquisição do instrumento
            instr.WriteString("INIT"); // Comece a varredura
            instr.QueryOpc(); // Usando * OPC? consulta espera até que o instrumento termine a aquisição
            Thread.Sleep(5000);
            string potencia = instr.QueryString("CALC:MARK:FUNC:ADEM:CARR?");
            Console.WriteLine(potencia);
            SalvaValores(nomePasta, nomeSubPasta, potencia, freqC);
        }

        void Mascara(RsInstrument instr, string[] freqC)
        {
            for (int i = 0; i < 21; i++)
            {
                instr.WriteString("INST:SEL SAN"); //Configura o Analisador para o Spectrum mode
                instr.WriteString("CALC:UNIT:POW DBM"); //Configura a unidade do reference Level
                instr.WriteString("INP:ATT 35dB"); //Configura o ATT
                instr.WriteString("DISP:TRAC:Y:RLEV 0dbm"); // Configura o Reference Level
                instr.WriteString($"FREQ:CENT {freqC[i]} MHz"); // Configura a Frequencia Central
                instr.WriteString("FREQ:SPAN 150 KHz"); // Configura o span
                instr.WriteString("BAND 1 kHz"); // Configura o RBW
                instr.WriteString("BAND:VID 1 kHz"); // Configura o VBW
                instr.WriteString("SWE:TIME:AUTO ON"); // Configura o sweep points
                //instr.WriteString("DISP:TRAC:MODE CLRWR"); //Configura o Trace
                instr.WriteString("DET POS"); //Configura o Trace
                instr.QueryOpc(); //Usando * OPC? consulta espera até que todas as configurações do instrumento sejam concluídas
                // -----------------------------------------------------------
                // SyncPoint 'SettingsApplied' - todas as configurações foram aplicadas
                // -----------------------------------------------------------
                instr.VisaTimeout = 2000; // tempo limite de varredura - defina ele mais alto do que o tempo de aquisição do instrumento
                instr.WriteString("INIT"); // Comece a varredura
                instr.QueryOpc(); // Usando * OPC? consulta espera até que o instrumento termine a aquisição
                // -----------------------------------------------------------
                // SyncPoint 'AcquisitionFinished' - os resultados estão prontos
                // -----------------------------------------------------------

                // Fazendo uma captura de tela do instrumento e transferindo o arquivo para o PC
                // -----------------------------------------------------------
                Thread.Sleep(5000);
                instr.WriteString("HCOP:DEV:LANG WMF");
                instr.WriteString($@"MMEM:NAME 'c:\temp\print{i}.WMF'");
                instr.WriteString("HCOP:IMM"); // Faça a captura de tela
                instr.QueryOpc(); // Espere a captura de tela ser salva
                instr.File.FromInstrumentToPc($@"c:\temp\print{i}.WMF", $@"C:\Users\80400197\Desktop\Prints radio\print{i}.WMF"); // envia o arquivo do instrumento para o PC
                Console.WriteLine(@"Arquivo de captura de tela salvo no PC");
                Console.WriteLine("Pressione qualquer tecla para continuar");
                Console.ReadKey();

            }
        }


        static void Main()
        {
            Radios Radio;
            
            RsInstrument instr; //Cria o Objeto instr como classe RsInstrument
            RsInstrument.AssertMinVersion("1.10.1");
            //Inicia as variaveis do marker, com valores padrao para entrar no While
            string[] freqC = {"148.41", "159.41", "173.95", "148.01", "152.61", "148.21", "152.81", "148.39", "152.99", "149.01", "153.61", "149.43", "154.03", "149.89", "154.49", "164.61", "169.21", "165.09", "169.69", "165.59", "170.19", "157.45875", "162.05875", "158.42125", "163.02125", "159,39625", "163.99625", "157.45625", "162.05625", "158.41875", "163.01875", "163.01875", "163.99375", "157.45625", "162.05625", "167.38125","171.98125", "169.18125", "173.78125"};
            

            try // Criar um Try-catch separado para inicialização impede o acesso a objetos não inicializados
            {
                //-----------------------------------------------------------
                // Inicialização:
                //-----------------------------------------------------------

                // Ajuste a string de recursos VISA para se adequar ao seu instrumento 
                instr = new RsInstrument("TCPIP0::FSMR3-200017::inst0::INSTR", false, false);
                Radio = new Radios();
                instr.VisaTimeout = 5000; // Tempo limite para operações de leitura VISA
                instr.OpcTimeout = 15000; // Tempo limite para operações sincronizadas com opc
                instr.InstrumentStatusChecking = true; // Verificação de erro após cada comando
            }
            catch (RsInstrumentException e)
            {
                Console.WriteLine($"Erro ao inicializar a sessão do instrumento:\n{e.Message}");
                Console.WriteLine("Pressione qualquer tecla para sair.");
                Console.ReadKey();
                return;
            }

            try // Try o bloco para capturar qualquer RsInstrumentException()
            {
                Console.WriteLine($"Versão do drive do instrumento: {instr.Identification.DriverVersion}, Versao do nucleo: {instr.Identification.CoreVersion}");
                instr.ClearStatus(); //Limpe os buffers de io do instrumento
                Console.WriteLine($"String de identificação do instrumento:\n{instr.Identification.IdnString}");
                instr.WriteString("*RST;*CLS"); // Reinicialize o instrumento, limpe a fila de erros
                instr.WriteString("INIT:CONT ON"); // Desliga a varredura contínua
                instr.WriteString("SYST:DISP:UPD ON"); // Display update ON - Desligar após a depuração
                //-----------------------------------------------------------
                // Basic Settings:
                //-----------------------------------------------------------
                Radio.Potencia(instr, "148.41", "Potencia.csv", @"C:\Users\80400197\Desktop\Prints radio\674");
                Radio.EstabilidadeFreq(instr, "148.41", "estabilidadeEfreq.csv", @"C:\Users\80400197\Desktop\Prints radio\674");
                Radio.Mascara(instr, freqC);
                

                
            }
            catch (RsInstrumentException e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                Console.WriteLine("Pressione qualquer tecla para terminar...");
                Console.ReadKey();
            }
        }
    }
}