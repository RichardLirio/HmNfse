// Método que constroi objeto GerarNfseEnvioPmv
using SchemaXmlPmv;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;
using NotafiscalService;
using System.ServiceModel;
using HmNfse;
using System;
using System.Linq;
using System.Net;

namespace Service
{
    public class ServicePmv
    {
        // Criação do Bind utilizado pela prefeitura de Vitoria-ES
        public static BasicHttpsBinding bindPMV()
        {
            var binding = new BasicHttpsBinding(BasicHttpsSecurityMode.Transport);
            binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.Certificate;

            return binding;
        }


        // Método para extrair dados e montar xml da Nota fiscal de serviço
        public static GerarNfseEnvio ExtrairDadosNotaPmv(string dadosXML)
        {

            var txt = Hm_Nfse.splitDadosNfse(dadosXML);//array com os dados da nota separados por ;

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

            servico.ItemListaServico = txt[9];
            servico.Discriminacao = txt[6];//descrição do serviço prestado
            servico.CodigoMunicipio = Convert.ToInt32(txt[7]);//codigo municipio
            servico.ExigibilidadeISS = Convert.ToSByte(txt[8]);//Exigibilidades possíveis 1 – Exigível;2 – Não incidência;3 – Isenção;4 – Exportação; 5 – Imunidade;6 – Exigibilidade Suspensa por Decisão Judicial;7 – Exigibilidade Suspensa por Processo Administrativo.
            servico.MunicipioIncidencia = Convert.ToInt32(txt[7]);//codigo municipio
            servico.CodigoTributacaoMunicipio = txt[9];//codigo referente ao serviço de acordo com a prefeitura
            servico.MunicipioIncidencia = Convert.ToInt32(txt[7]);//codigo municipio
            servico.MunicipioIncidenciaSpecified = true;
            if (txt[9] == "13.04") { servico.CodigoCnae = Convert.ToInt32("8219901"); } else { servico.CodigoCnae = Convert.ToInt32(txt[9]); }
            
            servico.CodigoCnaeSpecified = true;


            tcIdentificacaoPrestador pessoaEmpresaPre = new tcIdentificacaoPrestador();
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
            tcIdentificacaoTomador pessoaEmpresaTom = new tcIdentificacaoTomador();
            pessoaEmpresaTom.CpfCnpj = cpfcnpjtom;
            tomador.IdentificacaoTomador = pessoaEmpresaTom;
            tomador.RazaoSocial = txt[14];//nome do tomador do serviço

            tcEndereco endtomador = new tcEndereco();
            endtomador.Endereco = txt[15];//endereço tomador
            endtomador.Numero = txt[16];//numero endereço
            endtomador.Bairro = txt[17];//bairro tomador
            endtomador.CodigoMunicipio = Convert.ToInt32(txt[18]);//codigo municipio tomador

            //verifica e seleciona a UF do tomador
            if (txt[19] == "AC") { endtomador.Uf = "AC"; }
            if (txt[19] == "AL") { endtomador.Uf = "AL"; }
            if (txt[19] == "AM") { endtomador.Uf = "AM"; }
            if (txt[19] == "AP") { endtomador.Uf = "AP"; }
            if (txt[19] == "BA") { endtomador.Uf = "BA"; }
            if (txt[19] == "CE") { endtomador.Uf = "CE"; }
            if (txt[19] == "DF") { endtomador.Uf = "DF"; }
            if (txt[19] == "ES") { endtomador.Uf = "ES"; }
            if (txt[19] == "GO") { endtomador.Uf = "GO"; }
            if (txt[19] == "MA") { endtomador.Uf = "MA"; }
            if (txt[19] == "MG") { endtomador.Uf = "MG"; }
            if (txt[19] == "MS") { endtomador.Uf = "MS"; }
            if (txt[19] == "MT") { endtomador.Uf = "MT"; }
            if (txt[19] == "PA") { endtomador.Uf = "PA"; }
            if (txt[19] == "PB") { endtomador.Uf = "PB"; }
            if (txt[19] == "PE") { endtomador.Uf = "PE"; }
            if (txt[19] == "PI") { endtomador.Uf = "PI"; }
            if (txt[19] == "PR") { endtomador.Uf = "PR"; }
            if (txt[19] == "RJ") { endtomador.Uf = "RJ"; }
            if (txt[19] == "RN") { endtomador.Uf = "RN"; }
            if (txt[19] == "RO") { endtomador.Uf = "RO"; }
            if (txt[19] == "RR") { endtomador.Uf = "RR"; }
            if (txt[19] == "RS") { endtomador.Uf = "RS"; }
            if (txt[19] == "SC") { endtomador.Uf = "SC"; }
            if (txt[19] == "SE") { endtomador.Uf = "SE"; }
            if (txt[19] == "SP") { endtomador.Uf = "SP"; }
            if (txt[19] == "TO") { endtomador.Uf = "TO"; }

            //cep tomador
            endtomador.Cep = txt[20];//cep tomador

            tcContato contatoTomador = new tcContato();

            if (txt[22] == "")
            {


                contatoTomador.Telefone = txt[21];

            }
            else
            {
                contatoTomador.Telefone = txt[21];
                contatoTomador.Email = txt[22];

            }

            tomador.Endereco = endtomador;
            tomador.Contato = contatoTomador;


            tcInfDeclaracaoPrestacaoServico declara = new tcInfDeclaracaoPrestacaoServico();
            declara.Id = "Id_" + Guid.NewGuid().ToString("N");//id do rps precisa ter letra e numero
            declara.Rps = InfRps;
            declara.Competencia = Convert.ToDateTime(txt[23]);//competencia na nfse
            declara.Tomador = tomador;
            declara.Prestador = pessoaEmpresaPre;
            declara.Servico = servico;

            //Se o prestador é optante pelo simples nacional tsSimNao -> 1 = Sim , 2 = Não
            if (txt[24] == "") { declara.OptanteSimplesNacional = Convert.ToSByte("2"); } else { declara.OptanteSimplesNacional = Convert.ToSByte("1"); }

            //paremtros referente ao prestador de serviço
            declara.IncentivoFiscal = Convert.ToSByte(txt[25]);
            //if (txt[26] == "\n") { declara.InformacoesComplementares = null; } else { declara.InformacoesComplementares = txt[26]; }
            declara.RegimeEspecialTributacao = Convert.ToSByte(txt[27]);
            declara.RegimeEspecialTributacaoSpecified = true;


            tcDeclaracaoPrestacaoServico tcDeclaracao = new tcDeclaracaoPrestacaoServico();
            tcDeclaracao.InfDeclaracaoPrestacaoServico = declara;


            GerarNfseEnvio gerarNfseEnvio = new GerarNfseEnvio();
            gerarNfseEnvio.Rps = tcDeclaracao;

            return gerarNfseEnvio;
        }


        // Método para assinar o XML
        public static string AssinarXmlPmv(GerarNfseEnvio gerarNfseEnvio, string sdircert, string senhacert)
        {
            // Lógica para assinar o XML da nota fiscal

            string sgerarNfseEnvio = Hm_Nfse.ClasseParaXmlString(gerarNfseEnvio);

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

                var evento = xmlDocument.GetElementsByTagName("Rps");
                evento[0].AppendChild(xmlSignature);

            }

            xmlDocument.Save("c:/temp/nfseRequest.xml");
            var xmlNfse = xmlDocument.InnerXml.ToString();

            return xmlNfse;
        }


        // Requisição para gerar nova nota fiscal de serviço no sistema da PMV
        public static string[] resquestGerarNfsePMV(string notaFiscal, string url, string sdirxml, string certificadoPath, string certificadoSenha, string dadosXML)
        {
            try
            {
                var txt = Hm_Nfse.splitDadosNfse(dadosXML);//array com os dados da nota separados por ;
                string snota = txt[0];

                var xmlNfse = new GerarNfseRequestBody();
                xmlNfse.mensagemXML = notaFiscal;

                // Carregar o certificado
                var cert = new X509Certificate2(certificadoPath, certificadoSenha);

                // Configuração da comunicação com o serviço
                var binding = bindPMV();
                var endpoint = new EndpointAddress(url);

                using (var client = new NotaFiscalServiceSoapClient(binding, endpoint))
                {
                    client.ClientCredentials.ClientCertificate.Certificate = cert;

                    // Salva xml da requisição para o caso de suporte
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.LoadXml(notaFiscal);
                    xmlDocument.Save("c:/temp/nfseRequest.xml");
                    xmlDocument.PreserveWhitespace = true;


                    // Realiza o envio da nota fiscal
                    var response = client.GerarNfseAsync(notaFiscal);

                    var resultado = Hm_Nfse.XmlStringParaClasse<GerarNfseResposta>(response.Result.Body.GerarNfseResult);

                    if (resultado.Item is GerarNfseRespostaListaNfse listaNfse)
                    {
                        var nNota = listaNfse.CompNfse.Nfse.InfNfse.Numero.ToString();
                        var nCodVerificação = listaNfse.CompNfse.Nfse.InfNfse.CodigoVerificacao.ToString();
                        nCodVerificação = nCodVerificação.Remove(nCodVerificação.Length - 1,1);


                        string slistaNfse = Hm_Nfse.ClasseParaXmlString(listaNfse);

                        XmlDocument XMLlistaNfse = new XmlDocument();
                        XMLlistaNfse.LoadXml(slistaNfse);
                        XMLlistaNfse.PreserveWhitespace = true;
                        XMLlistaNfse.Save($"{sdirxml}.xml");//salva o xml da nota dentro da pasta nfse do sistema
                        var sret = "";
                        if (nNota != snota) { sret = "ER206"; } else { sret = $"{nNota}"; }//caso gere uma nota fiscal com numero diferente do enviado pelo sistema
                        return new string[] { sret, nCodVerificação };//retorna o numero da nota e seu codigo de verificação.
                    }

                    else if (resultado.Item is ListaMensagemRetorno listaMensagemRetorno)//Caso retorne algum erro nas informações enviadas
                    {
                        string slistaMensagemRetorno = Hm_Nfse.ClasseParaXmlString(listaMensagemRetorno);
                        XmlDocument xmlErro = new XmlDocument();
                        xmlErro.LoadXml(slistaMensagemRetorno);
                        xmlErro.Save("c:/temp/RetornoErro.xml");
                        var sErro = "Erro";

                        tcMensagemRetorno[] arrayMsgret = listaMensagemRetorno.MensagemRetorno;
                        int iCont = arrayMsgret.Count() - 1;
                        if (iCont >= 1)
                        {
                            for (int i = 0; i <= iCont; i++)
                            {

                                return new string[] { listaMensagemRetorno.MensagemRetorno[i].Mensagem.ToString(), sErro };
                            }
                        }//if iCont < 1
                        else
                        {
                            return new string[] { listaMensagemRetorno.MensagemRetorno[0].Mensagem.ToString(), sErro };

                        }//else iCont < 1
                    }
                    return new string[] { "Erro ao tratar o retorno da chamada.", "Erro" };
                }

            }
            catch (Exception ex)
            {
                return new string[] { ex.ToString() };

            };

        }


        // Requisição de consulta na prefeitura de Vitoria-ES
        public static string[] requestConsultaNfsePMV(string certificadoPath, string certificadoSenha, string url, string sConsultarNfse)
        {
            try
            {
                var xmlConsultaNfse = new ConsultarNfseServicoPrestadoRequestBody();
                xmlConsultaNfse.mensagemXML = sConsultarNfse;

                // Carregar o certificado
                var cert = new X509Certificate2(certificadoPath, certificadoSenha);

                // Configuração da comunicação com o serviço
                var binding = bindPMV();
                var endpoint = new EndpointAddress(url);

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;//PROTOCOLO DE SEGURANÇA

                // Criação do cliente do serviço
                using (var client = new NotaFiscalServiceSoapClient(binding, endpoint))
                {
                    client.ClientCredentials.ClientCertificate.Certificate = cert;

                    // Realiza a consulta
                    var response = client.ConsultarNfseServicoPrestadoAsync(sConsultarNfse);

                    // Salva o xml utilizado para a requisição para vizualização em caso de suporte,
                    XmlDocument xmlresquest = new XmlDocument();
                    xmlresquest.LoadXml(sConsultarNfse);
                    xmlresquest.PreserveWhitespace = true;
                    xmlresquest.Save("c:/temp/RequestConsulta.xml");

                    // Captura o retorno da chamada
                    var resultado = Hm_Nfse.XmlStringParaClasse<ConsultarNfseServicoPrestadoResposta>(response.Result.Body.ConsultarNfseServicoPrestadoResult);

                    if (resultado.Item is ConsultarNfseServicoPrestadoRespostaListaNfse listaNfse)
                    {
                        var nNota = listaNfse.CompNfse[0].Nfse.InfNfse.Numero.ToString();
                        var nCodVerificação = listaNfse.CompNfse[0].Nfse.InfNfse.CodigoVerificacao.ToString();

                        // Salva o xml de retorno para a requisição para vizualização em caso de suporte,
                        string slistaNfse = Hm_Nfse.ClasseParaXmlString(listaNfse);
                        XmlDocument XMLlistaNfse = new XmlDocument();
                        XMLlistaNfse.LoadXml(slistaNfse);
                        XMLlistaNfse.PreserveWhitespace = true;
                        XMLlistaNfse.Save("c:/temp/RetornoConsulta.xml");

                        return new string[] { nNota, nCodVerificação }; // Nota já autorizada, código de autenticação da nota
                    }
                    else if (resultado.Item is ListaMensagemRetorno listaMensagemRetorno)
                    {
                        var mensagemRetorno = listaMensagemRetorno.MensagemRetorno[0].Mensagem;
                        return new string[] { mensagemRetorno, "ER504" }; // Nota não localizada, erro 504
                    }
                    else
                    {
                        return new string[] { "Erro desconhecido ao processar resposta", "ER500" }; // Erro desconhecido, código 500
                    }
                }
            }
            catch (Exception ex)
            {
                return new string[] { ex.ToString() };

            };

        }

        // Cancela Nfse
        public static string[] CancelamentoNfsePmv(string sdircert, string senhacert, string infosNfse, string sdirxml, string url)
        {
            

            try
            {
                var sCancelaNfse = MontaDadosCancelamentoNfse(infosNfse);

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(sCancelaNfse);

                var xmlCancelamentoNfse = new CancelarNfseRequestBody();
                xmlCancelamentoNfse.mensagemXML = sCancelaNfse;

                // Carregar o certificado
                var cert = new X509Certificate2(sdircert, senhacert);

                // Configuração da comunicação com o serviço
                var binding = bindPMV();
                var endpoint = new EndpointAddress(url);


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

                var xmlNfse = xmlDocument.InnerXml.ToString();

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;//PROTOCOLO DE SEGURANÇA
                string remoteAddress = url;//endpoint da prefeitura de acordo com o codigo de municipio

                // Criação do cliente do serviço
                using (var client = new NotaFiscalServiceSoapClient(binding, endpoint))
                {
                    client.ClientCredentials.ClientCertificate.Certificate = cert;

                    // Realiza a consulta
                    var response = client.CancelarNfseAsync(xmlNfse);

                    xmlDocument.Save("c:/temp/RequestCancelamento.xml");

                    // Captura o retorno da chamada
                    var resultado = Hm_Nfse.XmlStringParaClasse<CancelarNfseResposta>(response.Result.Body.CancelarNfseResult);


                    if (resultado.Item is tcRetCancelamento retCancelamento)
                    {

                        var confirmaCancelamento = retCancelamento.NfseCancelamento.ConfirmacaoCancelamento.ToString();

                        string sretCancelamento = Hm_Nfse.ClasseParaXmlString(retCancelamento);

                        XmlDocument xmlRetCancelamento = new XmlDocument();
                        xmlRetCancelamento.LoadXml(sretCancelamento);
                        xmlRetCancelamento.Save($"{sdirxml}_Cancelamento.xml");
                        return new string[] { "CN200" };
                    }
                    else if (resultado.Item is ListaMensagemRetorno listaMensagemRetorno)//Caso retorne algum erro nas informações enviadas
                    {
                        string slistaMensagemRetorno = Hm_Nfse.ClasseParaXmlString(listaMensagemRetorno);
                        XmlDocument xmlErro = new XmlDocument();
                        xmlErro.LoadXml(slistaMensagemRetorno);
                        xmlErro.Save("c:/temp/RetornoErro.xml");
                        var sErro = "Erro";

                        tcMensagemRetorno[] arrayMsgret = listaMensagemRetorno.MensagemRetorno;
                        int iCont = arrayMsgret.Count() - 1;
                        if (iCont >= 1)
                        {
                            for (int i = 0; i <= iCont; i++)
                            {

                                return new string[] { listaMensagemRetorno.MensagemRetorno[i].Mensagem.ToString(), sErro };
                            }
                        }//if iCont < 1
                        else
                        {
                            return new string[] { listaMensagemRetorno.MensagemRetorno[0].Mensagem.ToString(), sErro };

                        }//else iCont < 1
                    }
                    return new string[] { "Erro ao tratar o retorno da chamada.", "Erro" };
                }
            }
            catch (Exception ex)
            {
                return new string[] { "ERR", ex.ToString() };
            }
        }


        // Monta corpo da requisição de cancelamento da Nfse
        public static string MontaDadosCancelamentoNfse(string infosNfse)
        {
            var txt = infosNfse.Split(';');

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
            InfoPedidoCancelamento.CodigoCancelamentoSpecified = true;
            InfoPedidoCancelamento.Id = "Cancelamento_nfse_" + txt[1];

            tcPedidoCancelamento pedidoCancelamento = new tcPedidoCancelamento();
            pedidoCancelamento.InfPedidoCancelamento = InfoPedidoCancelamento;

            CancelarNfseEnvio cancelarNfse = new CancelarNfseEnvio();
            cancelarNfse.Pedido = pedidoCancelamento;

            string sCancelarNfse = Hm_Nfse.ClasseParaXmlString(cancelarNfse);

            

            return sCancelarNfse;
        }
    }


}

