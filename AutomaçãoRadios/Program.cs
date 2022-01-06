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
    
    class Program
    {
        static void Main()
        {
            RsInstrument instr; //Cria o Objeto instr como classe RsInstrument
            RsInstrument.AssertMinVersion("1.10.1");
            //Inicia as variaveis do marker, com valores padrao para entrar no While
            string[] freqC = {"148,41", "159,41", "173,95","148,41", "159,41","173,95", "148,01", "152,61", "148,21", "152,81", "148,39", "152,99", "149,01", "153,61", "149,43", "154,03", "149,89", "154,49", "164,61", "169,21", "165,09", "169,69", "165,59", "165,59", "170,19", "157,45875", "162,05875", "158,42125", "163,02125", "159,39625", "163,99625", "157,45625", "162,05625", "158,41875", "163,01875", "163,01875", "163,99375", "157,45625", "162,05625", "167,38125","171,98125", "169,18125", "173,78125"};

            //---------------------------------------------------------------
            //Cria Uma pasta para salvar os valores
            //---------------------------------------------------------------

            // Nome da pasta.
            string nomePasta = @"c:\teste Automação";

            // Cria uma subPasta na pasta teste Automação
            // Adicionando o nome Valores do Marker dentro da variavel Nome Pasta
            System.IO.Path.Combine(nomePasta, "valores do Marker");

            System.IO.Directory.CreateDirectory(nomePasta);

            //Cria uma variavel com o nome do arquivo que quer criar
            string fileName = "ValoresTeste.csv";

            // Combina o nome do arquivo ao caminho
            System.IO.Path.Combine(nomePasta, fileName);

            // Verifica o Caminho
            Console.WriteLine("Os valores vão ser salvos no caminho: {0}\n", nomePasta);




            try // Criar um Try-catch separado para inicialização impede o acesso a objetos não inicializados
            {
                //-----------------------------------------------------------
                // Inicialização:
                //-----------------------------------------------------------

                // Ajuste a string de recursos VISA para se adequar ao seu instrumento 
                instr = new RsInstrument("TCPIP0::FSMR3-200017::inst0::INSTR", false, false);
                instr.VisaTimeout = 3000; // Tempo limite para operações de leitura VISA
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
                //instr.WriteString("*RST;*CLS"); // Reinicialize o instrumento, limpe a fila de erros
                instr.WriteString("INIT:CONT ON"); // Desliga a varredura contínua
                instr.WriteString("SYST:DISP:UPD ON"); // Display update ON - Desligar após a depuração
                //-----------------------------------------------------------
                // Basic Settings:
                //-----------------------------------------------------------
                for (int i = 0; i < 42; i++)
                { 
                    instr.WriteString("INST:SEL SAN"); //Configura o Analisador para o Spectrum mode
                    instr.WriteString("CALC:UNIT:POW DBM"); //Configura a unidade do reference Level
                    instr.WriteString("INP:ATT 45dB"); //Configura o ATT
                    instr.WriteString("DISP:TRAC:Y:RLEV 0dbm"); // Configura o Reference Level
                    instr.WriteString($"FREQ:CENT {freqC[i]} MHz"); // Configura a Frequencia Central
                    instr.WriteString("FREQ:SPAN 150 MHz"); // Configura o span
                    instr.WriteString("BAND 1 kHz"); // Configura o RBW
                    instr.WriteString("BAND:VID 1 kHz"); // Configura o VBW
                    instr.WriteString("SWE:TIME:AUTO ON"); // Configura o sweep points
                    instr.WriteString("DISP:TRAC:MODE CLRWR"); //Configura o Trace
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
                    instr.WriteString("HCOP:DEV:LANG PNG");
                    instr.WriteString($@"MMEM:NAME 'c:\temp\Dev_Screenshot{freqC[i]}.png'");
                    instr.WriteString("HCOP:IMM"); // Faça a captura de tela
                    instr.QueryOpc(); // Espere a captura de tela ser salva
                    instr.File.FromInstrumentToPc($@"c:\temp\Dev_Screenshot{freqC[i]}.png", $@"C:\Windows\Temp\PC_Screenshot{freqC[i]}.png"); // envia o arquivo do instrumento para o PC
                    Console.WriteLine(@"Arquivo de captura de tela salvo no PC 'C:\Windows\Temp\PC_Screenshot.png'");
                    Thread.Sleep(5000);
                }

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