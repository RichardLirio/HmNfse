using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml;
using AracruzNfse;
using NfseWsService;
using System.ServiceModel.Channels;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Net;
using System.Net.NetworkInformation;
using System.ServiceModel;
using NotafiscalService;
using SchemaXmlAbrasf;
using Service;
using System.Security.Policy;

namespace HmNfse
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual),
  Guid("66EFA8BE-A461-4DD0-84E2-035D87DE07C5")]
    //Versão 1.23 - > Inclusão da prefeitura de Vitoria.
    //Versão 1.13 - > Inclusão da varivavel com o endereço do endpoint de requisição. Aracruz e São gabriel da palha
    //Versão 1.03 - > Correção caso o tomador do serviço não possua email
    //Versão 1.02 - > inclusão dos valores de impostos, alteração da função para retornar um Array com o numero da nota e com o codigo de autorização.
    public interface IHm_Nfse
    {
        string[] Geranfse(string dadosXML, string sdircert, string senhacert, string sdirxml, string url, string codmun);

        string[] CancelamentoNfse(string sdircert, string senhacert, string infosNfse, string sdirxml, string url, string codmun);

        string[] ConsultarNfse(string infosNfse, string url, string sdircert, string senhacert, string codmun);

        string[] Array();

    }

    [ClassInterface(ClassInterfaceType.None),
    Guid("9B014D4C-D2BA-4405-9709-2985D60CC681")]

    public class Hm_Nfse : IHm_Nfse
    {
        public Hm_Nfse() { }

        public static string ClasseParaXmlString<T>(T objeto)
        {
            System.Xml.Linq.XElement xml;
            var ser = new XmlSerializer(typeof(T));

            using (var memory = new MemoryStream())
            {
                using (TextReader tr = new StreamReader(memory, Encoding.UTF8))
                {
                    ser.Serialize(memory, objeto);
                    memory.Position = 0;
                    xml = XElement.Load(tr);
                    xml.Attributes().Where(x => x.Name.LocalName.Equals("xsd") || x.Name.LocalName.Equals("xsi")).Remove();
                }
            }
            return XElement.Parse(xml.ToString()).ToString(SaveOptions.DisableFormatting);
        }

        public static T XmlStringParaClasse<T>(string input) where T : class
        {
            var ser = new XmlSerializer(typeof(T));

            using (var sr = new StringReader(input))
                return (T)ser.Deserialize(sr);
        }

        /// <returns>Retorna a string contendo o node XML cujo nome foi passado no parâmetro nomeDoNode</returns>
        public static string ObterNodeDeArquivoXml(string nomeDoNode, string arquivoXml)
        {
            var xmlDoc = XDocument.Load(arquivoXml);
            var xmlString = (from d in xmlDoc.Descendants()
                             where d.Name.LocalName == nomeDoNode
                             select d).FirstOrDefault();

            if (xmlString == null)
                throw new Exception(String.Format("Nenhum objeto {0} encontrado no arquivo {1}!", nomeDoNode, arquivoXml));
            return xmlString.ToString();
        }

        // Criação do cabeçalho de requisição Soap
        public string montaCabec()
        {
            // Configuração do cabeçalho
            var cabecalho = new cabecalho { versaoDados = "2.04" };

            string scabec = ClasseParaXmlString(cabecalho);

            return scabec;
        }

        // Monta corpo da requisição de consulta de Nfse com as informações da empresa e da nota fiscal
        public string montaDadosConsultaNfse(string infosNfse)
        {
            var splitConsulta = infosNfse.Split(';');

            // Configuração do prestador
            var prestador = new tcIdentificacaoPessoaEmpresa
            {
                CpfCnpj = new tcCpfCnpj
                {
                    ItemElementName = ItemChoiceType.Cnpj,
                    Item = splitConsulta[0] // CNPJ do prestador
                },
                InscricaoMunicipal = splitConsulta[1] // Inscrição municipal do prestador
            };

            // Configuração da consulta
            var requestConsulta = new ConsultarNfseServicoPrestadoEnvio
            {
                Prestador = prestador,
                Item = splitConsulta[2], // Número da nota fiscal
                Pagina = "1"
            };

            string sConsultarNfse = ClasseParaXmlString(requestConsulta);

            return sConsultarNfse;

        }


        // Método para extrair os dados da nota do XML
        public static string[] splitDadosNfse(string dadosXML)
        {
            string sdadosXML = dadosXML.Replace("\\n", "\n");
            var txt = sdadosXML.Split(';');//array com os dados da nota separados por ;
            return txt;
        }

        // Método que constroi objeto GerarNfseEnvio
        public GerarNfseEnvio ExtrairDadosNota(string dadosXML)
        {

            var txt = splitDadosNfse(dadosXML);//array com os dados da nota separados por ;

            tcIdentificacaoRps identificacaoRps = new tcIdentificacaoRps();
            identificacaoRps.Numero = txt[0];//numero do rps
            identificacaoRps.Serie = "1";
            identificacaoRps.Tipo = 1;

            tcInfRps InfRps = new tcInfRps();
            InfRps.IdentificacaoRps = identificacaoRps;
            InfRps.DataEmissao = Convert.ToDateTime(txt[1]);//data emissão rps
            InfRps.Status = 1;//status rps
            InfRps.Id = Guid.NewGuid().ToString("N").Substring(0, 9);//id rps

            tcDadosServico servico = new tcDadosServico();
            decimal dval = Convert.ToDecimal(txt[3]);//valor do serviços
            decimal dissret = Convert.ToDecimal(txt[4]);//valor iss retido
            decimal daliq = Convert.ToDecimal(txt[5]);//valor aliquota
            decimal dinss = Convert.ToDecimal(txt[28]);//valor do inss
            decimal dir = Convert.ToDecimal(txt[29]);//valor do imposto de renda
            decimal dcsll = Convert.ToDecimal(txt[30]);//valor do csll
            decimal dpis = Convert.ToDecimal(txt[31]);//valor do pis
            decimal dcofins = Convert.ToDecimal(txt[32]);//valor do cofins
            decimal ddeducoes = Convert.ToDecimal(txt[33]);//valor das deduções
            decimal ddesconto = Convert.ToDecimal(txt[34]);//valor do desconto

            decimal dvaliss = dval * (daliq / 100);
            dvaliss = Math.Round(dvaliss, 2);

            //VALORES DO SERVIÇO
            tcValoresDeclaracaoServico valorserv = new tcValoresDeclaracaoServico();
            valorserv.ValorServicos = dval;
            valorserv.ValorIss = dvaliss;
            valorserv.Aliquota = daliq;
            valorserv.ValorInss = dinss;
            valorserv.ValorIr = dir;
            valorserv.ValorCsll = dcsll;
            valorserv.ValorPis = dpis;
            valorserv.ValorCofins = dcofins;
            valorserv.ValorDeducoes = ddeducoes;
            valorserv.DescontoCondicionado = ddesconto;

            //colocar para criar as classes dentro do xml
            valorserv.ValorIssSpecified = true;
            valorserv.AliquotaSpecified = true;
            valorserv.ValorInssSpecified = true;
            valorserv.ValorIrSpecified = true;
            valorserv.ValorCsllSpecified = true;
            valorserv.ValorPisSpecified = true;
            valorserv.ValorCofinsSpecified = true;
            valorserv.ValorDeducoesSpecified = true;
            valorserv.DescontoCondicionadoSpecified = true;


            //verifca de há retenção de iss
            decimal zero = Convert.ToDecimal(0);
            servico.Valores = valorserv;
            if (dissret == zero) { servico.IssRetido = Convert.ToSByte(2); } else { servico.IssRetido = Convert.ToSByte(1); }//retem iss - 1 = sim , 2 = não

            servico.ItemListaServico = tsItemListaServico.Item0107;
            servico.Discriminacao = txt[6];//descrição do serviço prestado
            servico.CodigoMunicipio = Convert.ToInt32(txt[7]);//codigo municipio
            servico.ExigibilidadeISS = Convert.ToSByte(txt[8]);//Exigibilidades possíveis 1 – Exigível;2 – Não incidência;3 – Isenção;4 – Exportação; 5 – Imunidade;6 – Exigibilidade Suspensa por Decisão Judicial;7 – Exigibilidade Suspensa por Processo Administrativo.
            servico.MunicipioIncidencia = Convert.ToInt32(txt[7]);//codigo municipio
            servico.CodigoTributacaoMunicipio = txt[9];//codigo referente ao serviço de acordo com a prefeitura


            tcIdentificacaoPessoaEmpresa pessoaEmpresaPre = new tcIdentificacaoPessoaEmpresa();
            tcCpfCnpj cpfcnpj = new tcCpfCnpj();
            cpfcnpj.ItemElementName = ItemChoiceType.Cnpj;
            cpfcnpj.Item = txt[10];//cnpj do prestador de serviço
            pessoaEmpresaPre.CpfCnpj = cpfcnpj;
            pessoaEmpresaPre.InscricaoMunicipal = txt[11];//inscrição municipal do prestador de serviço

            //verifica se o tomador é pessoa fisica ou juridica
            tcCpfCnpj cpfcnpjtom = new tcCpfCnpj();
            if (txt[12] == "F")//0 - cpf; 1 = cnpj
            {
                cpfcnpjtom.ItemElementName = ItemChoiceType.Cpf;//fisica
                cpfcnpjtom.Item = txt[13];
            }
            else
            {
                cpfcnpjtom.ItemElementName = ItemChoiceType.Cnpj;//juridica
                cpfcnpjtom.Item = txt[13];
            }

            tcDadosTomador tomador = new tcDadosTomador();
            tcIdentificacaoPessoaEmpresa pessoaEmpresaTom = new tcIdentificacaoPessoaEmpresa();
            pessoaEmpresaTom.CpfCnpj = cpfcnpjtom;
            tomador.IdentificacaoTomador = pessoaEmpresaTom;
            tomador.RazaoSocial = txt[14];//nome do tomador do serviço

            tcEndereco endtomador = new tcEndereco();
            endtomador.Endereco = txt[15];//endereço tomador
            endtomador.Numero = txt[16];//numero endereço
            endtomador.Bairro = txt[17];//bairro tomador
            endtomador.CodigoMunicipio = Convert.ToInt32(txt[18]);//codigo municipio tomador

            //verifica e seleciona a UF do tomador
            if (txt[19] == "AC") { endtomador.Uf = tsUf.AC; }
            if (txt[19] == "AL") { endtomador.Uf = tsUf.AL; }
            if (txt[19] == "AM") { endtomador.Uf = tsUf.AM; }
            if (txt[19] == "AP") { endtomador.Uf = tsUf.AP; }
            if (txt[19] == "BA") { endtomador.Uf = tsUf.BA; }
            if (txt[19] == "CE") { endtomador.Uf = tsUf.CE; }
            if (txt[19] == "DF") { endtomador.Uf = tsUf.DF; }
            if (txt[19] == "ES") { endtomador.Uf = tsUf.ES; }
            if (txt[19] == "GO") { endtomador.Uf = tsUf.GO; }
            if (txt[19] == "MA") { endtomador.Uf = tsUf.MA; }
            if (txt[19] == "MG") { endtomador.Uf = tsUf.MG; }
            if (txt[19] == "MS") { endtomador.Uf = tsUf.MS; }
            if (txt[19] == "MT") { endtomador.Uf = tsUf.MT; }
            if (txt[19] == "PA") { endtomador.Uf = tsUf.PA; }
            if (txt[19] == "PB") { endtomador.Uf = tsUf.PB; }
            if (txt[19] == "PE") { endtomador.Uf = tsUf.PE; }
            if (txt[19] == "PI") { endtomador.Uf = tsUf.PI; }
            if (txt[19] == "PR") { endtomador.Uf = tsUf.PR; }
            if (txt[19] == "RJ") { endtomador.Uf = tsUf.RJ; }
            if (txt[19] == "RN") { endtomador.Uf = tsUf.RN; }
            if (txt[19] == "RO") { endtomador.Uf = tsUf.RO; }
            if (txt[19] == "RR") { endtomador.Uf = tsUf.RR; }
            if (txt[19] == "RS") { endtomador.Uf = tsUf.RS; }
            if (txt[19] == "SC") { endtomador.Uf = tsUf.SC; }
            if (txt[19] == "SE") { endtomador.Uf = tsUf.SE; }
            if (txt[19] == "SP") { endtomador.Uf = tsUf.SP; }
            if (txt[19] == "TO") { endtomador.Uf = tsUf.TO; }

            //cep tomador
            endtomador.Cep = txt[20];//cep tomador

            tcContato contatoTomador = new tcContato();

            if (txt[22] == "")
            {
                string[] itens = { txt[21] };//telefone e email tomador
                ItemsChoiceType[] itenschoice = new ItemsChoiceType[] { ItemsChoiceType.Telefone };
                contatoTomador.Items = itens;
                contatoTomador.ItemsElementName = itenschoice;

            }
            else
            {
                string[] itens = { txt[21], txt[22] };//telefone e email tomador
                ItemsChoiceType[] itenschoice = new ItemsChoiceType[] { ItemsChoiceType.Telefone, ItemsChoiceType.Email };
                contatoTomador.Items = itens;
                contatoTomador.ItemsElementName = itenschoice;
            }

            tomador.Item = endtomador;
            tomador.Contato = contatoTomador;


            tcInfDeclaracaoPrestacaoServico declara = new tcInfDeclaracaoPrestacaoServico();
            declara.Id = "Id_" + Guid.NewGuid().ToString("N");//id do rps precisa ter letra e numero
            declara.Rps = InfRps;
            declara.Competencia = Convert.ToDateTime(txt[23]);//competencia na nfse
            declara.TomadorServico = tomador;
            declara.Prestador = pessoaEmpresaPre;
            declara.Servico = servico;

            //Se o prestador é optante pelo simples nacional tsSimNao -> 1 = Sim , 2 = Não
            if (txt[24] == "") { declara.OptanteSimplesNacional = Convert.ToSByte("2"); } else { declara.OptanteSimplesNacional = Convert.ToSByte("1"); }

            //paremtros referente ao prestador de serviço
            declara.IncentivoFiscal = Convert.ToSByte(txt[25]);
            if (txt[26] == "\n") { declara.InformacoesComplementares = null; } else { declara.InformacoesComplementares = txt[26]; }
            declara.RegimeEspecialTributacao = Convert.ToSByte(txt[27]);
            declara.RegimeEspecialTributacaoSpecified = true;


            tcDeclaracaoPrestacaoServico tcDeclaracao = new tcDeclaracaoPrestacaoServico();
            tcDeclaracao.InfDeclaracaoPrestacaoServico = declara;


            GerarNfseEnvio gerarNfseEnvio = new GerarNfseEnvio();
            gerarNfseEnvio.Rps = tcDeclaracao;

            return gerarNfseEnvio;
        }

        // Método para assinar o XML
        public string AssinarXML(GerarNfseEnvio gerarNfseEnvio, string sdircert, string senhacert)
        {
            // Lógica para assinar o XML da nota fiscal

            string sgerarNfseEnvio = ClasseParaXmlString(gerarNfseEnvio);

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(sgerarNfseEnvio);

            xmlDocument.PreserveWhitespace = false;

            SignedXml signedXml;

            XmlNodeList ListInfNFe = xmlDocument.GetElementsByTagName("InfDeclaracaoPrestacaoServico");
            X509Certificate2 xCert;
            xCert = new X509Certificate2(sdircert, senhacert);

            foreach (XmlElement InfDeclaracaoPrestacaoServico in ListInfNFe)

            {

                string id = InfDeclaracaoPrestacaoServico.Attributes.GetNamedItem("Id").InnerText;
                signedXml = new SignedXml(InfDeclaracaoPrestacaoServico);
                signedXml.SigningKey = xCert.PrivateKey;

                System.Security.Cryptography.Xml.Reference reference = new System.Security.Cryptography.Xml.Reference("#" + id);
                reference.AddTransform(new XmlDsigEnvelopedSignatureTransform());
                reference.AddTransform(new XmlDsigC14NTransform());
                signedXml.AddReference(reference);

                System.Security.Cryptography.Xml.KeyInfo keyInfo = new System.Security.Cryptography.Xml.KeyInfo();
                keyInfo.AddClause(new KeyInfoX509Data(xCert));

                signedXml.KeyInfo = keyInfo;

                signedXml.ComputeSignature();

                XmlElement xmlSignature = xmlDocument.CreateElement("Signature", "http://www.w3.org/2000/09/xmldsig#");
                XmlElement xmlSignedInfo = signedXml.SignedInfo.GetXml();
                XmlElement xmlKeyInfo = signedXml.KeyInfo.GetXml();

                XmlElement xmlSignatureValue = xmlDocument.CreateElement("SignatureValue", xmlSignature.NamespaceURI);
                string signBase64 = Convert.ToBase64String(signedXml.Signature.SignatureValue);
                XmlText text = xmlDocument.CreateTextNode(signBase64);
                xmlSignatureValue.AppendChild(text);

                xmlSignature.AppendChild(xmlDocument.ImportNode(xmlSignedInfo, true));
                xmlSignature.AppendChild(xmlSignatureValue);
                xmlSignature.AppendChild(xmlDocument.ImportNode(xmlKeyInfo, true));

                var evento = xmlDocument.GetElementsByTagName("GerarNfseEnvio");
                evento[0].AppendChild(xmlSignature);

            }

            xmlDocument.Save("c:/temp/nfseRequest.xml");
            var xmlNfse = xmlDocument.InnerXml.ToString();

            return sgerarNfseEnvio;
        }

        /////////////////////////////////////////////////////////////////////////////////////////////



        // testar acrescenter o modelo de xml da pmv



        // SERVIÇOS //////////////////////////////////////////////////////////////////////////////////

        // Gera uma nova Nfse de serviço prestado
        public string[] Geranfse(string dadosXML, string sdircert, string senhacert, string sdirxml, string url, string codmun)
        {
            try
            {
                

                if (codmun == "3205309")// Prefeitura de Vitoria
                {
                    // Organizar os dados do XML em uma classe
                    var notaFiscal = ServicePmv.ExtrairDadosNotaPmv(dadosXML);
                    // Assinar o XML
                    var notaFiscalAssinada = ServicePmv.AssinarXmlPmv(notaFiscal, sdircert, senhacert);

                    //requisição PMV
                    return ServicePmv.resquestGerarNfsePMV(notaFiscalAssinada, url, sdirxml, sdircert, senhacert, dadosXML);
                }
                else
                {
                    // Organizar os dados do XML em uma classe
                    var notaFiscal = ExtrairDadosNota(dadosXML);

                    // Assinar o XML
                    var notaFiscalAssinada = AssinarXML(notaFiscal, sdircert, senhacert);

                    //Requisição abrasf
                    return resquestGerarNfseAbrasf(notaFiscalAssinada, url, sdirxml, dadosXML);
                }

            }
            catch (Exception ex)
            {
                return new string[] { ex.ToString(), "Erro" };
            }
        }

        // Cancela Nfse
        public string[] CancelamentoNfse(string sdircert, string senhacert, string infosNfse, string sdirxml, string url, string codmun)
        {
            try
            {


                if (codmun == "3205309")// Prefeitura de Vitoria
                {

                    //requisição PMV
                    return ServicePmv.CancelamentoNfsePmv(sdircert, senhacert, infosNfse, sdirxml, url);
                }
                else
                {

                    //Requisição abrasf
                    return requestCancelaNfseAbrasf(sdircert, senhacert, infosNfse, sdirxml, url);
                }

            }
            catch (Exception ex)
            {
                return new string[] { ex.ToString(), "Erro" };
            }
        }

        // Consulta se a Nfse de serviço prestado já foi autorizada
        public string[] ConsultarNfse(string infosNfse, string url, string sdircert, string senhacert, string codmun)
        {
            var splitConsulta = infosNfse.Split(';');

            try
            {
                // Serialização dos objetos para XML
                var cabecalho = montaCabec();
                var dadosConsultaNfse = montaDadosConsultaNfse(infosNfse);

                //requisição de acodo com a prefeitura
                if (codmun == "3205309")//prefeitura de Vitoria - ES
                {
                    string[] requisicaoPMV = ServicePmv.requestConsultaNfsePMV(sdircert, senhacert, url, dadosConsultaNfse);

                    return requisicaoPMV;
                }
                else
                {
                    string[] requisicaoAbrasf = requestConsultaNfseAbrasf(url, cabecalho, dadosConsultaNfse);

                    return requisicaoAbrasf;
                }
               
            }
            catch(Exception ex)
            {
                return new string[] { ex.ToString(), "CATCH" }; 
            }
        
        
        }

        /////////////////////////////////////////////////////////////////////////////////////////////






        // REQUISIÇÕES ///////////////////////////////////////////////////////////////////////////


        // Requisição de consulta na prefeituras com sistema ABRASF
        public string[] requestConsultaNfseAbrasf(string url, string scabec, string sConsultarNfse)
        {
            try
            {
                var input = new NfseWsService.input();
                input.nfseCabecMsg = scabec;
                input.nfseDadosMsg = sConsultarNfse;

                string remoteAddress = url;//endpoint da prefeitura de acordo com o codigo de municipio

                var endpointConfiguration = new NfseWsService.nfseClient.EndpointConfiguration();

                var wse = new NfseWsService.nfseClient(endpointConfiguration, remoteAddress);

                var request = new NfseWsService.ConsultarNfseServicoPrestado();
                request.ConsultarNfseServicoPrestadoRequest = input;

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;//PROTOCOLO DE SEGURANÇA

                string sxmlrequest = ClasseParaXmlString(sConsultarNfse);

                XmlDocument xmlresquest = new XmlDocument();
                xmlresquest.LoadXml(sxmlrequest);
                xmlresquest.PreserveWhitespace = true;
                xmlresquest.Save("c:/temp/RequestConsulta.xml");

                // wse.ClientCredentials.ClientCertificate.Certificate = cert;
                var consulta = wse.ConsultarNfseServicoPrestadoAsync(request);
                consulta.Wait();


                //CAPTURANDO RETORNO DA CHAMADA
                var resp1 = consulta.Result.ConsultarNfseServicoPrestadoResponse;
                var sInnerText = ((System.Xml.XmlElement)((System.Xml.XmlNode[])resp1)[1]).InnerText;
                var InnerText = XmlStringParaClasse<ConsultarNfseServicoPrestadoResposta>(sInnerText);

                try
                {


                    ConsultarNfseServicoPrestadoRespostaListaNfse listaNfse = ((ConsultarNfseServicoPrestadoRespostaListaNfse)InnerText.Item);
                    var nNota = listaNfse.CompNfse[0].Nfse.InfNfse.Numero.ToString();
                    var nCodVerificação = listaNfse.CompNfse[0].Nfse.InfNfse.CodigoVerificacao.ToString();

                    string slistaNfse = ClasseParaXmlString(listaNfse);

                    XmlDocument XMLlistaNfse = new XmlDocument();
                    XMLlistaNfse.LoadXml(slistaNfse);
                    XMLlistaNfse.PreserveWhitespace = true;
                    XMLlistaNfse.Save("c:/temp/RetornoConsulta.xml");

                    return new string[] { nNota, nCodVerificação };//NOTA JA AUTORIZADA
                }
                catch
                {
                    ListaMensagemRetorno listamsgret = ((ListaMensagemRetorno)InnerText.Item);
                    tcMensagemRetorno[] arrayMsgret = listamsgret.MensagemRetorno;
                    string erromsg = arrayMsgret[0].Mensagem;

                    return new string[] { erromsg, "ER504" };//NOTA NAO LOCALIZADA ERRO 504

                }

            }
            catch (Exception ex)
            {
                return new string[] { ex.ToString(), "CATCH" };
            }

        }

        // Requisição para gerar nova nota fiscal de serviço no sistema ABRASF
        public string[] resquestGerarNfseAbrasf(string notaFiscal, string url, string sdirxml, string dadosXML)
        {

            var txt = splitDadosNfse(dadosXML);//array com os dados da nota separados por ;
            string snota = txt[0];

            // Lógica para enviar a requisição para a prefeitura e retornar a resposta
            var scabec = montaCabec();

            NfseWsService.input input = new NfseWsService.input();
            input.nfseCabecMsg = scabec;//cab.ToString();
            input.nfseDadosMsg = notaFiscal;//declara.ToString();

            string remoteAddress = url;//endpoint da prefeitura de acordo com o codigo de municipio

            var endpointConfiguration = new NfseWsService.nfseClient.EndpointConfiguration();

            var wse = new NfseWsService.nfseClient(endpointConfiguration, remoteAddress);

            var request = new NfseWsService.GerarNfse();
            request.GerarNfseRequest = input;

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;//PROTOCOLO DE SEGURANÇA

            var consulta = wse.GerarNfseAsync(request);
            consulta.Wait();

            //capturando o retorno da chamada para a prefeitura
            var response = consulta.Result.GerarNfseResponse;
            var sInnerText = ((System.Xml.XmlElement)((System.Xml.XmlNode[])response)[1]).InnerText;
            var InnerText = XmlStringParaClasse<GerarNfseResposta>(sInnerText);

            try
            {

                GerarNfseRespostaListaNfse listaNfse = ((GerarNfseRespostaListaNfse)InnerText.Item);
                var nNota = listaNfse.CompNfse.Nfse.InfNfse.Numero.ToString();
                var nCodVerificação = listaNfse.CompNfse.Nfse.InfNfse.CodigoVerificacao.ToString();

                string slistaNfse = ClasseParaXmlString(listaNfse);

                XmlDocument XMLlistaNfse = new XmlDocument();
                XMLlistaNfse.LoadXml(slistaNfse);
                XMLlistaNfse.PreserveWhitespace = true;
                XMLlistaNfse.Save($"{sdirxml}.xml");//salva o xml da nota dentro da pasta nfse do sistema
                var sret = "";
                if (nNota != snota) { sret = "ER206"; } else { sret = $"{nNota}"; }//caso gere uma nota fiscal com numero diferente do enviado pelo sistema
                return new string[] { sret, nCodVerificação };//retorna o numero da nota e seu codigo de verificação.
            }
            catch
            {
                XmlDocument xmlErro = new XmlDocument();
                xmlErro.LoadXml(sInnerText);
                xmlErro.Save("c:/temp/RetornoErro.xml");
                var sErro = "Erro";

                ListaMensagemRetorno listamsgret = ((ListaMensagemRetorno)InnerText.Item);
                tcMensagemRetorno[] arrayMsgret = listamsgret.MensagemRetorno;
                int iCont = arrayMsgret.Count() - 1;
                if (iCont >= 1)
                {
                    for (int i = 0; i <= iCont; i++)
                    {

                        return new string[] { listamsgret.MensagemRetorno[i].Mensagem.ToString(), sErro };
                    }
                }//if iCont < 1
                else
                {
                    return new string[] { listamsgret.MensagemRetorno[0].Mensagem.ToString(), sErro };

                }//else iCont < 1

            }
            return new string[] { "Erro ao tratar o retorno da chamada.", "Erro" };
        }

        public string[] requestCancelaNfseAbrasf(string sdircert, string senhacert, string infosNfse, string sdirxml, string url)
        {

            var txt = infosNfse.Split(';');

            try
            {
                cabecalho cab = new cabecalho();
                cab.versao = "2.04";
                cab.versaoDados = "2.04";

                tcCpfCnpj cpfcnpj = new tcCpfCnpj();
                cpfcnpj.ItemElementName = ItemChoiceType.Cnpj;
                cpfcnpj.Item = txt[0];//cnpj do prestador de serviço

                tcIdentificacaoNfse identificacaoNfse = new tcIdentificacaoNfse();
                identificacaoNfse.Numero = txt[1];//numero da nfse
                identificacaoNfse.CpfCnpj = cpfcnpj;
                identificacaoNfse.InscricaoMunicipal = txt[2];//inscrição municipal prestador
                identificacaoNfse.CodigoMunicipio = Convert.ToInt32(txt[3]);//codigo_municipio_prestação

                tcInfPedidoCancelamento InfoPedidoCancelamento = new tcInfPedidoCancelamento();
                InfoPedidoCancelamento.IdentificacaoNfse = identificacaoNfse;
                InfoPedidoCancelamento.CodigoCancelamento = Convert.ToSByte("1");
                InfoPedidoCancelamento.Id = "Cancelamento_nfse_" + txt[1];

                tcPedidoCancelamento pedidoCancelamento = new tcPedidoCancelamento();
                pedidoCancelamento.InfPedidoCancelamento = InfoPedidoCancelamento;

                CancelarNfseEnvio cancelarNfse = new CancelarNfseEnvio();
                cancelarNfse.Pedido = pedidoCancelamento;

                string scabec = ClasseParaXmlString(cab);
                string sCancelarNfse = ClasseParaXmlString(cancelarNfse);

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(sCancelarNfse);


                ////ASSINAR
                xmlDocument.PreserveWhitespace = false;
                SignedXml signedXml;

                XmlNodeList ListInfNFe = xmlDocument.GetElementsByTagName("InfPedidoCancelamento");
                X509Certificate2 xCert;
                xCert = new X509Certificate2(sdircert, senhacert);

                foreach (XmlElement InfPedidoCancelamento in ListInfNFe)

                {

                    string id = InfPedidoCancelamento.Attributes.GetNamedItem("Id").InnerText;
                    signedXml = new SignedXml(InfPedidoCancelamento);
                    signedXml.SigningKey = xCert.PrivateKey;

                    System.Security.Cryptography.Xml.Reference reference = new System.Security.Cryptography.Xml.Reference("#" + id);
                    reference.AddTransform(new XmlDsigEnvelopedSignatureTransform());
                    reference.AddTransform(new XmlDsigC14NTransform());
                    signedXml.AddReference(reference);

                    System.Security.Cryptography.Xml.KeyInfo keyInfo = new System.Security.Cryptography.Xml.KeyInfo();
                    keyInfo.AddClause(new KeyInfoX509Data(xCert));

                    signedXml.KeyInfo = keyInfo;

                    signedXml.ComputeSignature();

                    XmlElement xmlSignature = xmlDocument.CreateElement("Signature", "http://www.w3.org/2000/09/xmldsig#");
                    XmlElement xmlSignedInfo = signedXml.SignedInfo.GetXml();
                    XmlElement xmlKeyInfo = signedXml.KeyInfo.GetXml();

                    XmlElement xmlSignatureValue = xmlDocument.CreateElement("SignatureValue", xmlSignature.NamespaceURI);
                    string signBase64 = Convert.ToBase64String(signedXml.Signature.SignatureValue);
                    XmlText text = xmlDocument.CreateTextNode(signBase64);
                    xmlSignatureValue.AppendChild(text);

                    xmlSignature.AppendChild(xmlDocument.ImportNode(xmlSignedInfo, true));
                    xmlSignature.AppendChild(xmlSignatureValue);
                    xmlSignature.AppendChild(xmlDocument.ImportNode(xmlKeyInfo, true));

                    var evento = xmlDocument.GetElementsByTagName("Pedido");
                    evento[0].AppendChild(xmlSignature);

                }

                xmlDocument.Save("c:/temp/RequestCancelamento.xml");

                NfseWsService.input input = new NfseWsService.input();
                input.nfseCabecMsg = scabec;//cab.ToString();
                input.nfseDadosMsg = sCancelarNfse;//CancelarNfseEnvio.ToString();

                string remoteAddress = url;//endpoint da prefeitura de acordo com o codigo de municipio

                var endpointConfiguration = new NfseWsService.nfseClient.EndpointConfiguration();

                var wse = new NfseWsService.nfseClient(endpointConfiguration, remoteAddress);

                NfseWsService.CancelarNfse cancNfse = new NfseWsService.CancelarNfse();
                cancNfse.CancelarNfseRequest = input;

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;//PROTOCOLO DE SEGURANÇA

                var consulta = wse.CancelarNfseAsync(cancNfse);
                consulta.Wait();

                var response = consulta.Result.CancelarNfseResponse;

                var sInnerText = ((System.Xml.XmlElement)((System.Xml.XmlNode[])response)[1]).InnerText;
                var InnerText = XmlStringParaClasse<CancelarNfseResposta>(sInnerText);

                try
                {

                    tcRetCancelamento retCancelamento = ((tcRetCancelamento)InnerText.Item);
                    var confirmaCancelamento = retCancelamento.NfseCancelamento.Confirmacao.ToString();

                    string sretCancelamento = ClasseParaXmlString(retCancelamento);

                    XmlDocument xmlRetCancelamento = new XmlDocument();
                    xmlRetCancelamento.LoadXml(sretCancelamento);
                    xmlRetCancelamento.Save($"{sdirxml}_Cancelamento.xml");
                    return new string[] { "CN200" };

                }
                catch
                {

                    XmlDocument xmlErro = new XmlDocument();
                    xmlErro.LoadXml(sInnerText);
                    xmlErro.Save("c:/temp/RetornoErro.xml");

                    ListaMensagemRetorno listamsgret = ((ListaMensagemRetorno)InnerText.Item);
                    tcMensagemRetorno[] arrayMsgret = listamsgret.MensagemRetorno;
                    int iCont = arrayMsgret.Count() - 1;
                    if (iCont >= 1)
                    {
                        for (int i = 0; i <= iCont; i++)
                        {

                            return new string[] { listamsgret.MensagemRetorno[i].Codigo.ToString(), listamsgret.MensagemRetorno[i].Mensagem.ToString() };
                        }
                    }//if iCont < 1
                    else
                    {
                        return new string[] { listamsgret.MensagemRetorno[0].Codigo.ToString(), listamsgret.MensagemRetorno[0].Mensagem.ToString() };

                    }//else iCont < 1
                }

                return new string[] { "Err", "Erro no tratamento do retorno" };
            }
            catch (Exception ex)
            {
                return new string[] { "ERR", ex.ToString() };
            }
        }
 
        /////////////////////////////////////////////////////////////////////////////////////////////






        //referencia de como retornar 
        public string[] Array() { return new string[] { "Index 0", "Index 1" }; }


    }
}
