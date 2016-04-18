unit Principalu;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XPMenu, Menus, jpeg, ExtCtrls, DrLabel, StdCtrls, ComCtrls,
  ToolWin, funcoes, IEButton, D7ComboBoxStringsGetPatch, shellapi, Gradpan,
  IdAntiFreezeBase, IdAntiFreeze, IdMessage, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, comOBj;

function conecta_dados_web(banco: string):boolean;
function doc_declaracao(aluno: integer; tipo: string; data_hoje: tdatetime):boolean;
function doc_contrato(aluno: integer):boolean;
function ano_letivo():integer;
function envia_email(mensagem, email, de, email_dest, assunto, email_cc, host,
   username, passw, anexo: string; porta: integer): boolean;
function nota(Cod_Aluno, Disciplina: integer; Curso: String; Unidade: integer): double;
function disciplina(Matricula, Disciplina: integer): String;
function yWallPapper(Image1:TImage):boolean;
function yyear(data:tDateTime):integer;
function yday(data:tDateTime):integer;
function yTabd ( mes:integer; ano:integer ):tdatetime;
function ymonth(data:tDateTime):integer;
function yright(vr:string;n:integer): string;
function yiif(condicao:boolean; valor1:variant; valor2:variant):variant;
function alfzeros(vr:string;n:integer): string;
function allbnull(vr:string): string;
function acesso(Usuario: string; OpcaoMenu: String):boolean;

type
  TPrincipal = class(TForm)
    StatusBar1: TStatusBar;
    Menu: TMainMenu;
    Cadastros1: TMenuItem;
    Departamentos1: TMenuItem;
    Funcionrios1: TMenuItem;
    Funcionrios2: TMenuItem;
    N1: TMenuItem;
    Cursos1: TMenuItem;
    urmas1: TMenuItem;
    Disciplinas1: TMenuItem;
    N2: TMenuItem;
    Professores1: TMenuItem;
    Alunos1: TMenuItem;
    N3: TMenuItem;
    Movimentao1: TMenuItem;
    ContasaPagaer1: TMenuItem;
    Matrculas1: TMenuItem;
    Relatrios1: TMenuItem;
    Configuraes1: TMenuItem;
    N6: TMenuItem;
    Sair1: TMenuItem;
    Departamentos2: TMenuItem;
    Funcionrios3: TMenuItem;
    Fornecedores1: TMenuItem;
    Crusos1: TMenuItem;
    urmas2: TMenuItem;
    Disciplinas2: TMenuItem;
    Professores2: TMenuItem;
    Alunos2: TMenuItem;
    Eventos2: TMenuItem;
    Matrculas2: TMenuItem;
    Escalas2: TMenuItem;
    XPMenu1: TXPMenu;
    ToolBar1: TToolBar;
    Image1: TImage;
    Timer1: TTimer;
    Label1: TLabel;
    IEButton2: TIEButton;
    IEButton1: TIEButton;
    IEButton3: TIEButton;
    IEButton5: TIEButton;
    ServidorSMTPEmail1: TMenuItem;
    SobreoSistema1: TMenuItem;
    IEButton7: TIEButton;
    IEButton8: TIEButton;
    IEButton9: TIEButton;
    IEButton10: TIEButton;
    IEButton11: TIEButton;
    Ferramentas1: TMenuItem;
    Enviodeemail1: TMenuItem;
    EnviodeSMS1: TMenuItem;
    IEButton15: TIEButton;
    IEButton18: TIEButton;
    Mensalidades1: TMenuItem;
    Usurios1: TMenuItem;
    DadosdaEscola1: TMenuItem;
    MensalidadesemAberto1: TMenuItem;
    GradPan1: TGradPan;
    N10: TMenuItem;
    N11: TMenuItem;
    AnoLetivo1: TMenuItem;
    ContasBancrias1: TMenuItem;
    N12: TMenuItem;
    LanamentoBancrio1: TMenuItem;
    HistricosPadres1: TMenuItem;
    ContasBancrias2: TMenuItem;
    HistricosExtratoBancrio1: TMenuItem;
    Memomensagem: TMemo;
    ListBoxanexos: TListBox;
    IdSMTP1: TIdSMTP;
    msg: TIdMessage;
    IdAntiFreeze1: TIdAntiFreeze;
    Image3: TImage;
    DRLabel5: TDRLabel;
    DRLabel6: TDRLabel;
    LanarNotas1: TMenuItem;
    DRLabel1: TDRLabel;
    DRLabel2: TDRLabel;
    Cont1: TMenuItem;
    HistricosPadro1: TMenuItem;
    Declaraie1: TMenuItem;
    Mat1: TMenuItem;
    Aprovao1: TMenuItem;
    FrequenciaRegular1: TMenuItem;
    Contasapagar1: TMenuItem;
    ExtratoBancrio1: TMenuItem;
    N4: TMenuItem;
    ProcessamentoDirio1: TMenuItem;
    IEButton16: TIEButton;
    IEButton20: TIEButton;
    Produtos1: TMenuItem;
    SaveDialog1: TSaveDialog;
    Event1: TMenuItem;
    VendasdeProdutos1: TMenuItem;
    ContasaReceber1: TMenuItem;
    ContasaReceber2: TMenuItem;
    Renda1: TMenuItem;
    Quitao1: TMenuItem;
    Coordenadores1: TMenuItem;
    Coordenadores2: TMenuItem;
    Consult1: TMenuItem;
    AlunosAtivos1: TMenuItem;
    VendasnoPerodo1: TMenuItem;
    Responsaveis1: TMenuItem;
    GeraArqpSMS1: TMenuItem;
    procedure Sair1Click(Sender: TObject);
    procedure Departamentos1Click(Sender: TObject);
    procedure Funcionrios1Click(Sender: TObject);
    procedure Funcionrios2Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure DRLabel5Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Cursos1Click(Sender: TObject);
    procedure Unidades1Click(Sender: TObject);
    procedure urmas1Click(Sender: TObject);
    procedure Disciplinas1Click(Sender: TObject);
    procedure Professores1Click(Sender: TObject);
    procedure LivrosMdulos1Click(Sender: TObject);
    procedure Alunos1Click(Sender: TObject);
    procedure Produtos1Click(Sender: TObject);
    procedure ContasFluxo1Click(Sender: TObject);
    procedure Departamentos2Click(Sender: TObject);
    procedure ServidorSMTPEmail1Click(Sender: TObject);
    procedure SobreoSistema1Click(Sender: TObject);
    procedure Entradas1Click(Sender: TObject);
    procedure Enviodeemail1Click(Sender: TObject);
    procedure EnviodeSMS1Click(Sender: TObject);
    procedure Matrculas1Click(Sender: TObject);
    procedure ContasaPagaer1Click(Sender: TObject);
    procedure ContasaReceber1Click(Sender: TObject);
    procedure FluxodeCaixa1Click(Sender: TObject);
    procedure Eventos1Click(Sender: TObject);
    procedure Funcionrios3Click(Sender: TObject);
    procedure Fornecedores1Click(Sender: TObject);
    procedure Crusos1Click(Sender: TObject);
    procedure Unidades2Click(Sender: TObject);
    procedure urmas2Click(Sender: TObject);
    procedure Disciplinas2Click(Sender: TObject);
    procedure Professores2Click(Sender: TObject);
    procedure Alunos2Click(Sender: TObject);
    procedure LivrosMdulos2Click(Sender: TObject);
    procedure Produtos2Click(Sender: TObject);
    procedure ContasFluxo2Click(Sender: TObject);
    procedure Mensalidades1Click(Sender: TObject);
    procedure BaixadeMensalidades1Click(Sender: TObject);
    procedure Usurios1Click(Sender: TObject);
    procedure Entradas2Click(Sender: TObject);
    procedure DadosdaEscola1Click(Sender: TObject);
    procedure ContasaPagar1Click(Sender: TObject);
    procedure ContasaReceber2Click(Sender: TObject);
    procedure Eventos2Click(Sender: TObject);
    procedure Matrculas2Click(Sender: TObject);
    procedure ParametrosFinanceiroVendas1Click(Sender: TObject);
    procedure AnoLetivo1Click(Sender: TObject);
    procedure ContasBancrias1Click(Sender: TObject);
    procedure LanamentoBancrio1Click(Sender: TObject);
    procedure HistricosPadres1Click(Sender: TObject);
    procedure HistricosExtratoBancrio1Click(Sender: TObject);
    procedure ContasBancrias2Click(Sender: TObject);
    procedure LanarNotas1Click(Sender: TObject);
    procedure ProcessamentoDirio1Click(Sender: TObject);
    procedure Escalas2Click(Sender: TObject);
    procedure MensalidadesemAberto1Click(Sender: TObject);
    procedure Event1Click(Sender: TObject);
    procedure VendasdeProdutos1Click(Sender: TObject);
    procedure Cont1Click(Sender: TObject);
    procedure HistricosPadro1Click(Sender: TObject);
    procedure ExtratoBancrio1Click(Sender: TObject);
    procedure Mat1Click(Sender: TObject);
    procedure Aprovao1Click(Sender: TObject);
    procedure FrequenciaRegular1Click(Sender: TObject);
    procedure Renda1Click(Sender: TObject);
    procedure Quitao1Click(Sender: TObject);
    procedure Coordenadores1Click(Sender: TObject);
    procedure Coordenadores2Click(Sender: TObject);
    procedure AlunosAtivos1Click(Sender: TObject);
    procedure VendasnoPerodo1Click(Sender: TObject);
    procedure Responsaveis1Click(Sender: TObject);
    procedure GeraArqpSMS1Click(Sender: TObject);
  private
    { Private declarations }
    procedure ManipulaExcecoes(Sender: TObject; E: Exception);
  public
    { Public declarations }
  end;
var
  Principal: TPrincipal;
  Found, i, vcodigo, vvendedor, vano_letivo, vporta_smtp, vempr_cloud,
  vDias_sem_juros: integer;
  data_sys, vdata_atual: tdatetime;
  vlogin, cnpj_lib, vadministrador, Empresa, CgcEmp, cidemp, endemp, baiemp,
  vsimples_nac, numemp, vinscrmuni, inscr_muni, vfaxemp, estadoemp, s0, s1, s,
  s2, s3, s7, vgera_email, vversao, vcaracter_decimal, vsms_login, cepemp,
  vsms_password, vsms_remetente, vhost, vusuario_email, vsenha_email,
  vsms_de, vmens_cadastro, vmens_aniversariantes, sms_celular, foneemp, ve_mail,
  vbobina, vmatricial_usb, complemp, vsms_ponto, vbaixa_parcial, nome_fantasia,
  email_contador, email_relatorios, vlinha, vimprime_recibo, vcnpjlib, vita,
  vmar, vmes, tlinha, vparametro, vbanco_web: String;
  vpede_senha, vsenha_ok, avaliacao, ok_Sistem, vfecha, vconfigurado_ok,
  vcloud, vestacao, vacesso_ok: boolean;
  vjurosaomes: double;
  vcampos_contrato: tmemo;

implementation

uses Departamentosu, Funcionariosu, Fornecedoresu, FRegistrou,
  DataModule1u, FLocacaou, Fsenhau, Cursosu, Unidadesu,
  Turmasu, Disciplinasu, Professoresu, LivrosModulosu, Alunosu, Produtosu,
  ContasFluxou, FParametrosu, ImpDepartamentosu, Entradasu, Unit1,
  FSaldoSMSu, Matriculasu, FPagaru, FReceberu, FFluxou,
  FEventosu, ImpFuncionariosu, ImpFornecedoresu,
  ImpCursosu, ImpTurmasu, ImpDisciplinasu, ImpProfessoresu, ImpAlunosu,
  ImpProdutosu, ImpContasFluxou, ImpLivrosu, FMensalidadesu,
  FBaixaMensalidadesu, FCadastroUsuarios, ImpEntradasu,
  FDadosEmpresau, ImpContasPagaru, ImpContasReceberu,
  ImpEventosu, ImpMatriculasu, FParametrosFinanceirau,
  FAno_Letivou, FSelecao_Ano_Letivou,
  FContasBancariasu, FExtratou, FHistoricosPadroesu, ImpHistoricosu,
  ImpContasBacariau, FLancar_notasu, FCadastrou, FSenha_liberacaou,
  FConfiguraPastau, FProcessamento_Diariou, ImpFichadeAlunosu,
  ImpMensalidadesu, ImpExtratou, ImpDeclaracaoMatriculau, FVendasProdutosu,
  ImpDecl_Aprovacaou, ImpDecl_Frequenciau, ImpDecl_Rendau,
  ImpDecl_Quitacaou, FCoordenadoresu, ImpCoordenadores,
  FConsultaAlunos_Matriculasu, DataModule2u, ImpVendasu, ImpResponsaveisu,
  FarqSMSu;

{$R *.dfm}

function conecta_dados_web(banco: string):boolean;
begin
   with datamodule2 do begin
      ZConnection1.Database:='procedim_'+lowercase(banco);
   end;
   result:=true;
end;

Function campos_contrato(Campo: string):boolean;
var
vok: boolean;
i: integer;

begin
   vok:=false;
   with principal do begin
      for i:=0 to vcampos_contrato.Lines.Count-1 do begin
         if campo=vcampos_contrato.Lines.Strings[i] then vok:=true;
      end;
   end;
   result:=vok;
end;

function doc_contrato(aluno: integer):boolean;
var
MSWord, winword, docs, doc: variant;
vok: boolean;
vendereco_aluno, iextensao, varq, vresponsavel, vnacional_responsavel,
vnatural_de_responsavel, vcivil_responsavel, vfiliacao_responsavel,
vnascto_responsavel, vendereco_responsavel, vbairro_responsavel,
vcidade_responsavel, vuf_responsavel, vfone_responsavel,
vprofissao_responsavel, vrg_responsavel, vcpf_responsavel, vcep_responsavel: string;
vvalor, ventrada, vmensalidade: double;
vparcelas: integer;

begin
   MSWord:= CreateOleObject ('Word.Application'); //cria o objeto
   MSWord.Visible := True;
   vok:=false;
   if fileexists(s0+'documentos\contrato.docx') then begin
      Doc:= MSWord.Documents.Open(s0+'documentos\contrato.docx'); //caminho do documento padrao
      iextensao:='docx';
      vok:=true;
   end
   else
   begin
      Doc:= MSWord.Documents.Open(s0+'documentos\contrato.doc'); //caminho do documento padrao
      iextensao:='doc';
      vok:=true;
   end;
   if vok then begin
      with datamodule1 do begin
         vcampos_contrato.Lines.Add(ibdataset_parametrosCAMPOS_CONTRATO.Value);
         ibdataset_alunos.close;
         ibdataset_alunos.SelectSQL.Clear;
         ibdataset_alunos.SelectSQL.Add('Select * from alunos where codigo=:codigo order by codigo');
         ibdataset_alunos.Params[0].AsInteger:=aluno;
         ibdataset_alunos.Open;
         ibdataset_matriculas.Close;
         ibdataset_matriculas.SelectSQL.Clear;
         ibdataset_matriculas.SelectSQL.Add('Select * from matriculas where cod_aluno=:cod_aluno and ano=:ano order by cod_aluno');
         ibdataset_matriculas.Params[0].AsInteger:=ibdataset_alunosCODIGO.Value;
         ibdataset_matriculas.Params[1].AsInteger:=principalu.vano_letivo;
         ibdataset_matriculas.Open;
         ibdataset_mensalidades.Close;
         ibdataset_mensalidades.SelectSQL.Clear;
         ibdataset_mensalidades.SelectSQL.Add('Select * from mensalidades where matricula=:matricula');
         ibdataset_mensalidades.params[0].asinteger:=ibdataset_matriculasN_MATRICULA.Value;
         ibdataset_mensalidades.Open;
         vparcelas:=0;
         vmensalidade:=ibdataset_matriculasVALOR_MENSALIDADE.Value;
         if ibdataset_mensalidadesVALOR12.Value>0 then vparcelas:=1;
         if ibdataset_mensalidadesVALOR11.Value>0 then vparcelas:=2;
         if ibdataset_mensalidadesVALOR10.Value>0 then vparcelas:=3;
         if ibdataset_mensalidadesVALOR09.Value>0 then vparcelas:=4;
         if ibdataset_mensalidadesVALOR08.Value>0 then vparcelas:=5;
         if ibdataset_mensalidadesVALOR07.Value>0 then vparcelas:=6;
         if ibdataset_mensalidadesVALOR06.Value>0 then vparcelas:=7;
         if ibdataset_mensalidadesVALOR05.Value>0 then vparcelas:=8;
         if ibdataset_mensalidadesVALOR04.Value>0 then vparcelas:=9;
         if ibdataset_mensalidadesVALOR03.Value>0 then vparcelas:=10;
         if ibdataset_mensalidadesVALOR02.Value>0 then vparcelas:=11;
         if ibdataset_mensalidadesVALOR01.Value>0 then vparcelas:=12;
         //
         if ibdataset_mensalidadesVALOR01.Value>0 then
            ventrada:=ibdataset_mensalidadesVALOR01.Value
         else
         begin
            if ibdataset_mensalidadesVALOR02.Value>0 then
               ventrada:=ibdataset_mensalidadesVALOR02.Value
            else
            begin
               if ibdataset_mensalidadesVALOR03.Value>0 then
                  ventrada:=ibdataset_mensalidadesVALOR03.Value
               else
               begin
                  if ibdataset_mensalidadesVALOR04.Value>0 then
                     ventrada:=ibdataset_mensalidadesVALOR04.Value
                  else
                  begin
                     if ibdataset_mensalidadesVALOR05.Value>0 then
                        ventrada:=ibdataset_mensalidadesVALOR05.Value
                     else
                     begin
                        if ibdataset_mensalidadesVALOR06.Value>0 then
                           ventrada:=ibdataset_mensalidadesVALOR06.Value
                        else
                        begin
                           if ibdataset_mensalidadesVALOR07.Value>0 then
                              ventrada:=ibdataset_mensalidadesVALOR07.Value
                           else
                           begin
                              if ibdataset_mensalidadesVALOR08.Value>0 then
                                 ventrada:=ibdataset_mensalidadesVALOR08.Value
                              else
                              begin
                                 if ibdataset_mensalidadesVALOR09.Value>0 then
                                    ventrada:=ibdataset_mensalidadesVALOR09.Value
                                 else
                                 begin
                                    if ibdataset_mensalidadesVALOR10.Value>0 then
                                       ventrada:=ibdataset_mensalidadesVALOR10.Value
                                    else
                                    begin
                                       if ibdataset_mensalidadesVALOR11.Value>0 then
                                          ventrada:=ibdataset_mensalidadesVALOR11.Value
                                       else
                                       begin
                                          if ibdataset_mensalidadesVALOR12.Value>0 then
                                             ventrada:=ibdataset_mensalidadesVALOR12.Value
                                       end;
                                    end;
                                 end;
                              end;
                           end;
                        end;
                     end;
                  end;
               end;
            end;
         end;
         vvalor:=ibdataset_mensalidadesVALOR01.Value+ibdataset_mensalidadesVALOR02.Value+
                 ibdataset_mensalidadesVALOR03.Value+ibdataset_mensalidadesVALOR04.Value+
                 ibdataset_mensalidadesVALOR05.Value+ibdataset_mensalidadesVALOR06.Value+
                 ibdataset_mensalidadesVALOR07.Value+ibdataset_mensalidadesVALOR08.Value+
                 ibdataset_mensalidadesVALOR09.Value+ibdataset_mensalidadesVALOR10.Value+
                 ibdataset_mensalidadesVALOR11.Value+ibdataset_mensalidadesVALOR12.Value;
         //dados escola
         // dados do aluno
         vcep_responsavel:=' ';
         vnatural_de_responsavel:=' ';
         if ibdataset_alunosRESP_PREFERENCIAL.Value='P' then begin
            vresponsavel:=ibdataset_alunosNOME_RESPONSAVEL01.Value;
            vnacional_responsavel:='BRASILEIRA';
            vcivil_responsavel:='CASADO';
            vfiliacao_responsavel:=' ';
            if ibdataset_alunosNASCIMENTO_RESPONSAVEL01.Value>strtodate('01/01/1910') then
               vnascto_responsavel:=datetostr(ibdataset_alunosNASCIMENTO_RESPONSAVEL01.Value)
            else
               vnascto_responsavel:=' ';
            vcep_responsavel:=ibdataset_alunosCEP_RESPONSAVEL01.Value;
            vendereco_responsavel:=ibdataset_alunosEND_RESPONSAVEL01.Value+', '+
                                   ibdataset_alunosCOMPL_RESPONSAVEL01.Value;
            vbairro_responsavel:=ibdataset_alunosBAIRRO_RESPONSAVEL01.Value;
            vnatural_de_responsavel:=ibdataset_alunosNATURAL_RESPONSAVEL01.Value;
            vcidade_responsavel:=ibdataset_alunosCIDADE_RESPONSAVEL01.Value;
            vuf_responsavel:=ibdataset_alunosESTADO_RESPONSAVEL01.Value;
            vfone_responsavel:=sonumero(copy(ibdataset_alunosTEL_RESPONSAVEL01.Value,1,14));
            vprofissao_responsavel:=ibdataset_alunosPROFISSAO_RESPONSAVEL01.Value;
            vrg_responsavel:=ibdataset_alunosRG_RESPONSAVEL01.Value;
            vcpf_responsavel:=ibdataset_alunosCPF_RESPONSAVEL01.Value;
         end
         else
         begin
            if ibdataset_alunosRESP_PREFERENCIAL.Value='M' then begin
               vresponsavel:=ibdataset_alunosNOME_RESPONSAVEL02.Value;
               vnacional_responsavel:='BRASILEIRA';
               vcivil_responsavel:='CASADA';
               vfiliacao_responsavel:=' ';
               if ibdataset_alunosNASCIMENTO_RESPONSAVEL02.Value>strtodate('01/01/1910') then
                  vnascto_responsavel:=datetostr(ibdataset_alunosNASCIMENTO_RESPONSAVEL02.Value)
               else
                  vnascto_responsavel:=' ';
               vendereco_responsavel:=ibdataset_alunosEND_RESPONSAVEL02.Value+', '+
                                      ibdataset_alunosCOMPL_RESPONSAVEL02.Value;
               vbairro_responsavel:=ibdataset_alunosBAIRRO_RESPONSAVEL02.Value;
               vcidade_responsavel:=ibdataset_alunosCIDADE_RESPONSAVEL02.Value;
               vnatural_de_responsavel:=ibdataset_alunosNATURAL_RESPONSAVEL02.Value;
               vcep_responsavel:=ibdataset_alunosCEP_RESPONSAVEL02.Value;
               vuf_responsavel:=ibdataset_alunosESTADO_RESPONSAVEL02.Value;
               vfone_responsavel:=sonumero(copy(ibdataset_alunosTEL_RESPONSAVEL02.Value,1,14));
               vprofissao_responsavel:=ibdataset_alunosPROFISSAO_RESPONSAVEL02.Value;
               vrg_responsavel:=ibdataset_alunosRG_RESPONSAVEL02.Value;
               vcpf_responsavel:=ibdataset_alunosCPF_RESPONSAVEL02.Value;
            end
            else
            begin
               vresponsavel:=ibdataset_alunosNOME_RESPONSAVEL03.Value;
               vnacional_responsavel:='BRASILEIRA';
               vcivil_responsavel:=' ';
               vfiliacao_responsavel:=' ';
               if ibdataset_alunosNASCIMENTO_RESPONSAVEL03.Value>strtodate('01/01/1910') then
                  vnascto_responsavel:=datetostr(ibdataset_alunosNASCIMENTO_RESPONSAVEL03.Value)
               else
                  vnascto_responsavel:=' ';
               vendereco_responsavel:=ibdataset_alunosEND_RESPONSAVEL03.Value+', '+
                                      ibdataset_alunosCOMPL_RESPONSAVEL03.Value;
               vbairro_responsavel:=ibdataset_alunosBAIRRO_RESPONSAVEL03.Value;
               vcep_responsavel:=ibdataset_alunosCEP_RESPONSAVEL03.Value;
               vcidade_responsavel:=ibdataset_alunosCIDADE_RESPONSAVEL03.Value;
               vuf_responsavel:=ibdataset_alunosESTADO_RESPONSAVEL03.Value;
               vnatural_de_responsavel:=ibdataset_alunosNATURAL_RESPONSAVEL03.Value;
               vfone_responsavel:=sonumero(copy(ibdataset_alunosTEL_RESPONSAVEL03.Value,1,14));
               vprofissao_responsavel:=ibdataset_alunosPROFISSAO_RESPONSAVEL03.Value;
               vrg_responsavel:=ibdataset_alunosRG_RESPONSAVEL03.Value;
               vcpf_responsavel:=ibdataset_alunosCPF_RESPONSAVEL03.Value;
            end;
         end;
         vendereco_aluno:=ibdataset_alunosENDERECO.Value+', '+ibdataset_alunosNUMERO.Value+' '+
                          ibdataset_alunosCOMPLEMENTO.Value+' '+ibdataset_alunosBAIRRO.Value+
                          ' CEP: '+ibdataset_alunosCEP.Value+' - '+ibdataset_alunosCIDADE.Value+
                          '-'+ibdataset_alunosESTADO.Value;
         if campos_contrato('@escola@') then
            Doc.Content.Find.Execute(FindText:='@escola@', ReplaceWith:= principalu.empresa);
         if campos_contrato('@cnpj@') then
            Doc.Content.Find.Execute(FindText:='@cnpj@', ReplaceWith:= principalu.CgcEmp);
         if campos_contrato('@endereco@') then
            Doc.Content.Find.Execute(FindText:='@endereco@', ReplaceWith:= principalu.endemp);
         if campos_contrato('@numero@') then
            Doc.Content.Find.Execute(FindText:='@numero@', ReplaceWith:= principalu.numemp);
         if campos_contrato('@complemento@') then
            Doc.Content.Find.Execute(FindText:='@complemento@', ReplaceWith:= principalu.complemp);
         if campos_contrato('@bairro@') then
            Doc.Content.Find.Execute(FindText:='@bairro@', ReplaceWith:= principalu.baiemp);
         if campos_contrato('@cidade@') then
            Doc.Content.Find.Execute(FindText:='@cidade@', ReplaceWith:= principalu.cidemp);
         if campos_contrato('@cep@') then
            Doc.Content.Find.Execute(FindText:='@cep@', ReplaceWith:= principalu.cepemp);
         if campos_contrato('@estado@') then
            Doc.Content.Find.Execute(FindText:='@estado@', ReplaceWith:= principalu.estadoemp);
         if campos_contrato('@fone@') then
            Doc.Content.Find.Execute(FindText:='@fone@', ReplaceWith:= principalu.foneemp);
         if campos_contrato('@responsavel@') then
            Doc.Content.Find.Execute(FindText:='@responsavel@', ReplaceWith:=vresponsavel);
         if campos_contrato('@nac_resp@') then
            Doc.Content.Find.Execute(FindText:='@nac_resp@', ReplaceWith:=vnacional_responsavel);
         if campos_contrato('@nat_de_resp@') then
            Doc.Content.Find.Execute(FindText:='@nat_de_resp@', ReplaceWith:=vnatural_de_responsavel);
         if campos_contrato('@civil_resp@') then
            Doc.Content.Find.Execute(FindText:='@civil_resp@', ReplaceWith:=vcivil_responsavel);
         if campos_contrato('@filiacao_resp@') then
            Doc.Content.Find.Execute(FindText:='@filiacao_resp@', ReplaceWith:=vfiliacao_responsavel);
         if campos_contrato('@nasc_resp@') then
            Doc.Content.Find.Execute(FindText:='@nasc_resp@', ReplaceWith:=vnascto_responsavel);
         if campos_contrato('@end_resp@') then
            Doc.Content.Find.Execute(FindText:='@end_resp@', ReplaceWith:=vendereco_responsavel);
         if campos_contrato('@bairro_resp@') then
            Doc.Content.Find.Execute(FindText:='@bairro_resp@', ReplaceWith:=vbairro_responsavel);
         if campos_contrato('@cep_resp@') then
            Doc.Content.Find.Execute(FindText:='@cep_resp@', ReplaceWith:=vcep_responsavel);
         if campos_contrato('@cidade_resp@') then
            Doc.Content.Find.Execute(FindText:='@cidade_resp@', ReplaceWith:=vcidade_responsavel);
         if campos_contrato('@uf_resp@') then
            Doc.Content.Find.Execute(FindText:='@uf_resp@', ReplaceWith:=vuf_responsavel);
         if campos_contrato('@fone_resp@') then
            Doc.Content.Find.Execute(FindText:='@fone_resp@', ReplaceWith:=vfone_responsavel);
         if campos_contrato('@rg_resp@') then
            Doc.Content.Find.Execute(FindText:='@rg_resp@', ReplaceWith:=vrg_responsavel);
         if campos_contrato('@cpf_resp@') then
            Doc.Content.Find.Execute(FindText:='@cpf_resp@', ReplaceWith:=vcpf_responsavel);
         if campos_contrato('@prof_resp@') then
            Doc.Content.Find.Execute(FindText:='@prof_resp@', ReplaceWith:=vprofissao_responsavel);
         if campos_contrato('@aluno@') then
            Doc.Content.Find.Execute(FindText:='@aluno@', ReplaceWith:=ibdataset_alunosNOME.Value);
         if campos_contrato('@pai@') then
            Doc.Content.Find.Execute(FindText:='@pai@', ReplaceWith:=ibdataset_alunosPAI.Value);
         if campos_contrato('@mae@') then
            Doc.Content.Find.Execute(FindText:='@mae@', ReplaceWith:=ibdataset_alunosMAE.Value);
         if campos_contrato('@rg_aluno@') then
            Doc.Content.Find.Execute(FindText:='@rg_aluno@', ReplaceWith:=ibdataset_alunosRG.Value);
         if campos_contrato('@nat_de_aluno@') then
            Doc.Content.Find.Execute(FindText:='@nat_de_aluno@', ReplaceWith:=ibdataset_alunosNATURAL_DE.Value);
         if campos_contrato('@cpf_aluno@') then
            Doc.Content.Find.Execute(FindText:='@cpf_aluno@', ReplaceWith:=ibdataset_alunosCPF.Value);
         if campos_contrato('@fone_aluno@') then
            Doc.Content.Find.Execute(FindText:='@fone_aluno@', ReplaceWith:=ibdataset_alunosTELEFONE.Value);
         if campos_contrato('@nasc_aluno@') then begin
            if ibdataset_alunosNASCIMENTO.Value>strtodate('01/01/1910') then
               Doc.Content.Find.Execute(FindText:='@nasc_aluno@', ReplaceWith:=datetostr(ibdataset_alunosNASCIMENTO.Value))
            else
               Doc.Content.Find.Execute(FindText:='@nasc_aluno@', ReplaceWith:=' ');
         end;
         if campos_contrato('@nac_aluno@') then
            Doc.Content.Find.Execute(FindText:='@nac_aluno@', ReplaceWith:=ibdataset_alunosNACIONALIDADE.Value);
         if campos_contrato('@filiacao_aluno@') then
            Doc.Content.Find.Execute(FindText:='@filiacao_aluno@', ReplaceWith:=ibdataset_alunosPAI.Value+' / '+
                                                                  ibdataset_alunosMAE.Value);
         if campos_contrato('@valor@') then
            Doc.Content.Find.Execute(FindText:='@valor@', ReplaceWith:= formatfloat('##,##0.00',vmensalidade));
         if campos_contrato('@celular_aluno@') then
            Doc.Content.Find.Execute(FindText:='@celular_aluno@', ReplaceWith:=ibdataset_alunosCELULAR.Value);
         if campos_contrato('@endereco_aluno@') then
            Doc.Content.Find.Execute(FindText:='@endereco_aluno@', ReplaceWith:=ibdataset_alunosCELULAR.Value);
         if campos_contrato('@curso_aluno@') then
            Doc.Content.Find.Execute(FindText:='@curso_aluno@', ReplaceWith:=ibdataset_matriculasCURSO.Value);
         if campos_contrato('@serie_aluno@') then
            Doc.Content.Find.Execute(FindText:='@serie_aluno@', ReplaceWith:=ibdataset_matriculasSERIE.Value);
         if campos_contrato('@turma_aluno@') then
            Doc.Content.Find.Execute(FindText:='@turma_aluno@', ReplaceWith:=ibdataset_matriculasTURMA.Value);
         if campos_contrato('@turno@') then
            Doc.Content.Find.Execute(FindText:='@turno@', ReplaceWith:=ibdataset_matriculasTURMA.Value);
         // outros
         if campos_contrato('@ano_letivo@') then
            Doc.Content.Find.Execute(FindText:='@ano_letivo@', ReplaceWith:=inttostr(principalu.vano_letivo));
         if campos_contrato('@responsavel_aluno@') then begin
            if ibdataset_alunosRESP_PREFERENCIAL.Value='P' then
               Doc.Content.Find.Execute(FindText:='@responsavel_aluno@', ReplaceWith:=ibdataset_alunosNOME_RESPONSAVEL01.Value)
            else
            begin
               if ibdataset_alunosRESP_PREFERENCIAL.Value='M' then
                  Doc.Content.Find.Execute(FindText:='@responsavel_aluno@', ReplaceWith:=ibdataset_alunosNOME_RESPONSAVEL02.Value)
               else
                  Doc.Content.Find.Execute(FindText:='@responsavel_aluno@', ReplaceWith:=ibdataset_alunosNOME_RESPONSAVEL03.Value);
            end;
         end;
         if campos_contrato('@valor_total@') then
            Doc.Content.Find.Execute(FindText:='@valor_total@', ReplaceWith:=formatfloat('###,##0.00',vvalor));
         if campos_contrato('@valor_por_extenso@') then begin
            if vvalor>0 then
               Doc.Content.Find.Execute(FindText:='@valor_por_extenso@', ReplaceWith:=extenso(vvalor))
            else
               Doc.Content.Find.Execute(FindText:='@valor_por_extenso@', ReplaceWith:='');
         end;
         if campos_contrato('@parcelas@') then
            Doc.Content.Find.Execute(FindText:='@parcelas@', ReplaceWith:=inttostr(vparcelas));
         if campos_contrato('@parcelas_extenso@') then
            Doc.Content.Find.Execute(FindText:='@parcelas_extenso@', ReplaceWith:= numeroextenso1(vparcelas));
         if campos_contrato('@entrada@') then begin
            Doc.Content.Find.Execute(FindText:='@entrada@', ReplaceWith:= formatfloat('###,##0.00',ventrada));
            if ventrada>0 then
               Doc.Content.Find.Execute(FindText:='@entrada_extenso@', ReplaceWith:=extenso(ventrada))
            else
               Doc.Content.Find.Execute(FindText:='@entrada_extenso@', ReplaceWith:='');
         end;
         if campos_contrato('@parcelas_restante@') then
            Doc.Content.Find.Execute(FindText:='@parcelas_restante@', ReplaceWith:= inttostr(vparcelas-1));
         if campos_contrato('@restante_extenso@') then begin
            if vparcelas-1>0 then
               Doc.Content.Find.Execute(FindText:='@restante_extenso@', ReplaceWith:= numeroextenso1(vparcelas-1))
            else
               Doc.Content.Find.Execute(FindText:='@restante_extenso@', ReplaceWith:= '');
         end;
         if campos_contrato('@dia_vencimento@') then
            Doc.Content.Find.Execute(FindText:='@dia_vencimento@', ReplaceWith:= inttostr(ibdataset_matriculasDIA_MENSALIDADE.Value));
         if campos_contrato('@juros_dia@') then
            Doc.Content.Find.Execute(FindText:='@juros_dia@', ReplaceWith:= formatfloat('##0.00',ibdataset_parametrosJUROS.Value)+'%');
         if campos_contrato('@multa@') then
            Doc.Content.Find.Execute(FindText:='@multa@', ReplaceWith:= formatfloat('##0.00',ibdataset_parametrosMULTA.Value)+'%');
         if campos_contrato('@cidade_data@') then
            Doc.Content.Find.Execute(FindText:='@cidade_data@', ReplaceWith:= principalu.cidemp);
         if campos_contrato('@data_extenso@') then
            Doc.Content.Find.Execute(FindText:='@data_extenso@', ReplaceWith:= dataporextenso(ibdataset_matriculasDATA_MATRICULA.Value));
         //salvar
{         varq:='CONTRATO_'+buscatroca(ibdataset_alunosNOME.Value,' ','_')+'.'+iextensao;
         with principal do begin
            SaveDialog1.FileName:=varq;
            with SaveDialog1 do if execute then
               varq:=filename;
         end;
         MSWord.ActiveDocument.SaveAs(varq);
         Docs:=MSWord.Documents;
         SHOWMESSAGE('DOIS');
         Docs.Open(varq);
         SHOWMESSAGE('DOIS'); }
      end;
   end;
   result:=true;
end;


function doc_declaracao(aluno: integer; tipo: string; data_hoje: tdatetime):boolean;
var
MSWord, winword, docs, doc: variant;
vok: boolean;
vendereco_aluno, iextensao, varq, vresponsavel, vnacional_responsavel,
vnatural_de_responsavel, vcivil_responsavel, vfiliacao_responsavel,
vnascto_responsavel, vendereco_responsavel, vbairro_responsavel,
vcidade_responsavel, vuf_responsavel, vfone_responsavel, varqv1, varqv2,
vprofissao_responsavel, vrg_responsavel, vcpf_responsavel, vcep_responsavel: string;
vvalor, ventrada, vmensalidade: double;
vparcelas: integer;

begin
   MSWord:= CreateOleObject ('Word.Application'); //cria o objeto
   MSWord.Visible := True;
   vok:=false;
   varqv1:='decl_matricula.docx';
   varqv2:='decl_matricula.doc';
   if tipo='FREQUENCIA' then begin
      varqv1:='decl_frequencia.docx';
      varqv2:='decl_frequencia.doc';
   end;
   if tipo='APROVACAO' then begin
      varqv1:='decl_aprovacao.docx';
      varqv2:='decl_aprovacao.doc';
   end;
   if tipo='RENDA' then begin
      varqv1:='decl_renda.docx';
      varqv2:='decl_renda.doc';
   end;
   if tipo='QUITACAO' then begin
      varqv1:='decl_quitacao.docx';
      varqv2:='decl_quitacao.doc';
   end;
   if fileexists(s0+'documentos\'+varqv1) then begin
      Doc:= MSWord.Documents.Open(s0+'documentos\'+varqv1); //caminho do documento padrao
      iextensao:='docx';
      vok:=true;
   end
   else
   begin
      if fileexists(s0+'documentos\'+varqv2) then begin
         Doc:= MSWord.Documents.Open(s0+'documentos\'+varqv2); //caminho do documento padrao
         iextensao:='doc';
         vok:=true;
      end;
   end;
   if vok then begin
      with datamodule1 do begin
         if tipo='MATRICULA' then
            vcampos_contrato.Lines.Add(ibdataset_parametrosCAMPOS_DECL_MATRICULA.Value);
         if tipo='APROVACAO' then
            vcampos_contrato.Lines.Add(ibdataset_parametrosCAMPOS_DECL_APROVACAO.Value);
         if tipo='FREQUENCIA' then
            vcampos_contrato.Lines.Add(ibdataset_parametrosCAMPOS_DECL_FREQUENCIA.Value);
         if tipo='RENDA' then
            vcampos_contrato.Lines.Add(ibdataset_parametrosCAMPOS_DECL_RENDA.Value);
         if tipo='QUITACAO' then
            vcampos_contrato.Lines.Add(ibdataset_parametrosCAMPOS_DECL_QUITACAO.Value);
         ibdataset_alunos.close;
         ibdataset_alunos.SelectSQL.Clear;
         ibdataset_alunos.SelectSQL.Add('Select * from alunos where codigo=:codigo order by codigo');
         ibdataset_alunos.Params[0].AsInteger:=aluno;
         ibdataset_alunos.Open;
         ibdataset_matriculas.Close;
         ibdataset_matriculas.SelectSQL.Clear;
         ibdataset_matriculas.SelectSQL.Add('Select * from matriculas where cod_aluno=:cod_aluno and ano=:ano order by cod_aluno');
         ibdataset_matriculas.Params[0].AsInteger:=ibdataset_alunosCODIGO.Value;
         ibdataset_matriculas.Params[1].AsInteger:=principalu.vano_letivo;
         ibdataset_matriculas.Open;
         ibdataset_mensalidades.Close;
         ibdataset_mensalidades.SelectSQL.Clear;
         ibdataset_mensalidades.SelectSQL.Add('Select * from mensalidades where matricula=:matricula');
         ibdataset_mensalidades.params[0].asinteger:=ibdataset_matriculasN_MATRICULA.Value;
         ibdataset_mensalidades.Open;
         vparcelas:=0;
         vmensalidade:=ibdataset_matriculasVALOR_MENSALIDADE.Value;
         if ibdataset_mensalidadesVALOR12.Value>0 then vparcelas:=1;
         if ibdataset_mensalidadesVALOR11.Value>0 then vparcelas:=2;
         if ibdataset_mensalidadesVALOR10.Value>0 then vparcelas:=3;
         if ibdataset_mensalidadesVALOR09.Value>0 then vparcelas:=4;
         if ibdataset_mensalidadesVALOR08.Value>0 then vparcelas:=5;
         if ibdataset_mensalidadesVALOR07.Value>0 then vparcelas:=6;
         if ibdataset_mensalidadesVALOR06.Value>0 then vparcelas:=7;
         if ibdataset_mensalidadesVALOR05.Value>0 then vparcelas:=8;
         if ibdataset_mensalidadesVALOR04.Value>0 then vparcelas:=9;
         if ibdataset_mensalidadesVALOR03.Value>0 then vparcelas:=10;
         if ibdataset_mensalidadesVALOR02.Value>0 then vparcelas:=11;
         if ibdataset_mensalidadesVALOR01.Value>0 then vparcelas:=12;
         //
         if ibdataset_mensalidadesVALOR01.Value>0 then
            ventrada:=ibdataset_mensalidadesVALOR01.Value
         else
         begin
            if ibdataset_mensalidadesVALOR02.Value>0 then
               ventrada:=ibdataset_mensalidadesVALOR02.Value
            else
            begin
               if ibdataset_mensalidadesVALOR03.Value>0 then
                  ventrada:=ibdataset_mensalidadesVALOR03.Value
               else
               begin
                  if ibdataset_mensalidadesVALOR04.Value>0 then
                     ventrada:=ibdataset_mensalidadesVALOR04.Value
                  else
                  begin
                     if ibdataset_mensalidadesVALOR05.Value>0 then
                        ventrada:=ibdataset_mensalidadesVALOR05.Value
                     else
                     begin
                        if ibdataset_mensalidadesVALOR06.Value>0 then
                           ventrada:=ibdataset_mensalidadesVALOR06.Value
                        else
                        begin
                           if ibdataset_mensalidadesVALOR07.Value>0 then
                              ventrada:=ibdataset_mensalidadesVALOR07.Value
                           else
                           begin
                              if ibdataset_mensalidadesVALOR08.Value>0 then
                                 ventrada:=ibdataset_mensalidadesVALOR08.Value
                              else
                              begin
                                 if ibdataset_mensalidadesVALOR09.Value>0 then
                                    ventrada:=ibdataset_mensalidadesVALOR09.Value
                                 else
                                 begin
                                    if ibdataset_mensalidadesVALOR10.Value>0 then
                                       ventrada:=ibdataset_mensalidadesVALOR10.Value
                                    else
                                    begin
                                       if ibdataset_mensalidadesVALOR11.Value>0 then
                                          ventrada:=ibdataset_mensalidadesVALOR11.Value
                                       else
                                       begin
                                          if ibdataset_mensalidadesVALOR12.Value>0 then
                                             ventrada:=ibdataset_mensalidadesVALOR12.Value
                                       end;
                                    end;
                                 end;
                              end;
                           end;
                        end;
                     end;
                  end;
               end;
            end;
         end;
         vvalor:=ibdataset_mensalidadesVALOR01.Value+ibdataset_mensalidadesVALOR02.Value+
                 ibdataset_mensalidadesVALOR03.Value+ibdataset_mensalidadesVALOR04.Value+
                 ibdataset_mensalidadesVALOR05.Value+ibdataset_mensalidadesVALOR06.Value+
                 ibdataset_mensalidadesVALOR07.Value+ibdataset_mensalidadesVALOR08.Value+
                 ibdataset_mensalidadesVALOR09.Value+ibdataset_mensalidadesVALOR10.Value+
                 ibdataset_mensalidadesVALOR11.Value+ibdataset_mensalidadesVALOR12.Value;
         //dados escola
         // dados do aluno
         vcep_responsavel:=' ';
         vnatural_de_responsavel:=' ';
         if alltrim(ibdataset_alunosNOME_RESPONSAVEL01.Value)<>'' then begin
            vresponsavel:=ibdataset_alunosNOME_RESPONSAVEL01.Value;
            vnacional_responsavel:='BRASILEIRA';
            vcivil_responsavel:='CASADO';
            vfiliacao_responsavel:=' ';
            if ibdataset_alunosNASCIMENTO_RESPONSAVEL01.Value>strtodate('01/01/1910') then
               vnascto_responsavel:=datetostr(ibdataset_alunosNASCIMENTO_RESPONSAVEL01.Value)
            else
               vnascto_responsavel:=' ';
            vcep_responsavel:=ibdataset_alunosCEP_RESPONSAVEL01.Value;
            vendereco_responsavel:=ibdataset_alunosEND_RESPONSAVEL01.Value+', '+
                                   ibdataset_alunosCOMPL_RESPONSAVEL01.Value;
            vbairro_responsavel:=ibdataset_alunosBAIRRO_RESPONSAVEL01.Value;
            vnatural_de_responsavel:=ibdataset_alunosNATURAL_RESPONSAVEL01.Value;
            vcidade_responsavel:=ibdataset_alunosCIDADE_RESPONSAVEL01.Value;
            vuf_responsavel:=ibdataset_alunosESTADO_RESPONSAVEL01.Value;
            vfone_responsavel:=sonumero(copy(ibdataset_alunosTEL_RESPONSAVEL01.Value,1,14));
            vprofissao_responsavel:=ibdataset_alunosPROFISSAO_RESPONSAVEL01.Value;
            vrg_responsavel:=ibdataset_alunosRG_RESPONSAVEL01.Value;
            vcpf_responsavel:=ibdataset_alunosCPF_RESPONSAVEL01.Value;
         end
         else
         begin
            if alltrim(ibdataset_alunosNOME_RESPONSAVEL02.Value)<>'' then begin
               vresponsavel:=ibdataset_alunosNOME_RESPONSAVEL02.Value;
               vnacional_responsavel:='BRASILEIRA';
               vcivil_responsavel:='CASADA';
               vfiliacao_responsavel:=' ';
               if ibdataset_alunosNASCIMENTO_RESPONSAVEL02.Value>strtodate('01/01/1910') then
                  vnascto_responsavel:=datetostr(ibdataset_alunosNASCIMENTO_RESPONSAVEL02.Value)
               else
                  vnascto_responsavel:=' ';
               vendereco_responsavel:=ibdataset_alunosEND_RESPONSAVEL02.Value+', '+
                                      ibdataset_alunosCOMPL_RESPONSAVEL02.Value;
               vbairro_responsavel:=ibdataset_alunosBAIRRO_RESPONSAVEL02.Value;
               vcidade_responsavel:=ibdataset_alunosCIDADE_RESPONSAVEL02.Value;
               vnatural_de_responsavel:=ibdataset_alunosNATURAL_RESPONSAVEL02.Value;
               vcep_responsavel:=ibdataset_alunosCEP_RESPONSAVEL02.Value;
               vuf_responsavel:=ibdataset_alunosESTADO_RESPONSAVEL02.Value;
               vfone_responsavel:=sonumero(copy(ibdataset_alunosTEL_RESPONSAVEL02.Value,1,14));
               vprofissao_responsavel:=ibdataset_alunosPROFISSAO_RESPONSAVEL02.Value;
               vrg_responsavel:=ibdataset_alunosRG_RESPONSAVEL02.Value;
               vcpf_responsavel:=ibdataset_alunosCPF_RESPONSAVEL02.Value;
            end
            else
            begin
               vresponsavel:=ibdataset_alunosNOME_RESPONSAVEL03.Value;
               vnacional_responsavel:='BRASILEIRA';
               vcivil_responsavel:=' ';
               vfiliacao_responsavel:=' ';
               if ibdataset_alunosNASCIMENTO_RESPONSAVEL03.Value>strtodate('01/01/1910') then
                  vnascto_responsavel:=datetostr(ibdataset_alunosNASCIMENTO_RESPONSAVEL03.Value)
               else
                  vnascto_responsavel:=' ';
               vendereco_responsavel:=ibdataset_alunosEND_RESPONSAVEL03.Value+', '+
                                      ibdataset_alunosCOMPL_RESPONSAVEL03.Value;
               vbairro_responsavel:=ibdataset_alunosBAIRRO_RESPONSAVEL03.Value;
               vcep_responsavel:=ibdataset_alunosCEP_RESPONSAVEL03.Value;
               vcidade_responsavel:=ibdataset_alunosCIDADE_RESPONSAVEL03.Value;
               vuf_responsavel:=ibdataset_alunosESTADO_RESPONSAVEL03.Value;
               vnatural_de_responsavel:=ibdataset_alunosNATURAL_RESPONSAVEL03.Value;
               vfone_responsavel:=sonumero(copy(ibdataset_alunosTEL_RESPONSAVEL03.Value,1,14));
               vprofissao_responsavel:=ibdataset_alunosPROFISSAO_RESPONSAVEL03.Value;
               vrg_responsavel:=ibdataset_alunosRG_RESPONSAVEL03.Value;
               vcpf_responsavel:=ibdataset_alunosCPF_RESPONSAVEL03.Value;
            end;
         end;
         vendereco_aluno:=ibdataset_alunosENDERECO.Value+', '+ibdataset_alunosNUMERO.Value+' '+
                          ibdataset_alunosCOMPLEMENTO.Value+' '+ibdataset_alunosBAIRRO.Value+
                          ' CEP: '+ibdataset_alunosCEP.Value+' - '+ibdataset_alunosCIDADE.Value+
                          '-'+ibdataset_alunosESTADO.Value;
         if campos_contrato('@escola@') then
            Doc.Content.Find.Execute(FindText:='@escola@', ReplaceWith:= principalu.empresa);
         if campos_contrato('@cnpj@') then
            Doc.Content.Find.Execute(FindText:='@cnpj@', ReplaceWith:= principalu.CgcEmp);
         if campos_contrato('@endereco@') then
            Doc.Content.Find.Execute(FindText:='@endereco@', ReplaceWith:= principalu.endemp);
         if campos_contrato('@numero@') then
            Doc.Content.Find.Execute(FindText:='@numero@', ReplaceWith:= principalu.numemp);
         if campos_contrato('@complemento@') then
            Doc.Content.Find.Execute(FindText:='@complemento@', ReplaceWith:= principalu.complemp);
         if campos_contrato('@bairro@') then
            Doc.Content.Find.Execute(FindText:='@bairro@', ReplaceWith:= principalu.baiemp);
         if campos_contrato('@cidade@') then
            Doc.Content.Find.Execute(FindText:='@cidade@', ReplaceWith:= principalu.cidemp);
         if campos_contrato('@cep@') then
            Doc.Content.Find.Execute(FindText:='@cep@', ReplaceWith:= principalu.cepemp);
         if campos_contrato('@estado@') then
            Doc.Content.Find.Execute(FindText:='@estado@', ReplaceWith:= principalu.estadoemp);
         if campos_contrato('@fone@') then
            Doc.Content.Find.Execute(FindText:='@fone@', ReplaceWith:= principalu.foneemp);
         if campos_contrato('@responsavel@') then
            Doc.Content.Find.Execute(FindText:='@responsavel@', ReplaceWith:=vresponsavel);
         if campos_contrato('@nac_resp@') then
            Doc.Content.Find.Execute(FindText:='@nac_resp@', ReplaceWith:=vnacional_responsavel);
         if campos_contrato('@nat_de_resp@') then
            Doc.Content.Find.Execute(FindText:='@nat_de_resp@', ReplaceWith:=vnatural_de_responsavel);
         if campos_contrato('@civil_resp@') then
            Doc.Content.Find.Execute(FindText:='@civil_resp@', ReplaceWith:=vcivil_responsavel);
         if campos_contrato('@filiacao_resp@') then
            Doc.Content.Find.Execute(FindText:='@filiacao_resp@', ReplaceWith:=vfiliacao_responsavel);
         if campos_contrato('@nasc_resp@') then
            Doc.Content.Find.Execute(FindText:='@nasc_resp@', ReplaceWith:=vnascto_responsavel);
         if campos_contrato('@end_resp@') then
            Doc.Content.Find.Execute(FindText:='@end_resp@', ReplaceWith:=vendereco_responsavel);
         if campos_contrato('@bairro_resp@') then
            Doc.Content.Find.Execute(FindText:='@bairro_resp@', ReplaceWith:=vbairro_responsavel);
         if campos_contrato('@cep_resp@') then
            Doc.Content.Find.Execute(FindText:='@cep_resp@', ReplaceWith:=vcep_responsavel);
         if campos_contrato('@cidade_resp@') then
            Doc.Content.Find.Execute(FindText:='@cidade_resp@', ReplaceWith:=vcidade_responsavel);
         if campos_contrato('@uf_resp@') then
            Doc.Content.Find.Execute(FindText:='@uf_resp@', ReplaceWith:=vuf_responsavel);
         if campos_contrato('@fone_resp@') then
            Doc.Content.Find.Execute(FindText:='@fone_resp@', ReplaceWith:=vfone_responsavel);
         if campos_contrato('@rg_resp@') then
            Doc.Content.Find.Execute(FindText:='@rg_resp@', ReplaceWith:=vrg_responsavel);
         if campos_contrato('@cpf_resp@') then
            Doc.Content.Find.Execute(FindText:='@cpf_resp@', ReplaceWith:=vcpf_responsavel);
         if campos_contrato('@prof_resp@') then
            Doc.Content.Find.Execute(FindText:='@prof_resp@', ReplaceWith:=vprofissao_responsavel);
         if campos_contrato('@aluno@') then
            Doc.Content.Find.Execute(FindText:='@aluno@', ReplaceWith:=ibdataset_alunosNOME.Value);
         if campos_contrato('@rg_aluno@') then
            Doc.Content.Find.Execute(FindText:='@rg_aluno@', ReplaceWith:=ibdataset_alunosRG.Value);
         if campos_contrato('@nat_de_aluno@') then
            Doc.Content.Find.Execute(FindText:='@nat_de_aluno@', ReplaceWith:=ibdataset_alunosNATURAL_DE.Value);
         if campos_contrato('@pai@') then
            Doc.Content.Find.Execute(FindText:='@pai@', ReplaceWith:=ibdataset_alunosPAI.Value);
         if campos_contrato('@mae@') then
            Doc.Content.Find.Execute(FindText:='@mae@', ReplaceWith:=ibdataset_alunosMAE.Value);
         if campos_contrato('@cpf_aluno@') then
            Doc.Content.Find.Execute(FindText:='@cpf_aluno@', ReplaceWith:=ibdataset_alunosCPF.Value);
         if campos_contrato('@fone_aluno@') then
            Doc.Content.Find.Execute(FindText:='@fone_aluno@', ReplaceWith:=ibdataset_alunosTELEFONE.Value);
         if campos_contrato('@nasc_aluno@') then begin
            if ibdataset_alunosNASCIMENTO.Value>strtodate('01/01/1910') then
               Doc.Content.Find.Execute(FindText:='@nasc_aluno@', ReplaceWith:=datetostr(ibdataset_alunosNASCIMENTO.Value))
            else
               Doc.Content.Find.Execute(FindText:='@nasc_aluno@', ReplaceWith:=' ');
         end;
         if campos_contrato('@nac_aluno@') then
            Doc.Content.Find.Execute(FindText:='@nac_aluno@', ReplaceWith:=ibdataset_alunosNACIONALIDADE.Value);
         if campos_contrato('@filiacao_aluno@') then
            Doc.Content.Find.Execute(FindText:='@filiacao_aluno@', ReplaceWith:=ibdataset_alunosPAI.Value+' / '+
                                                                  ibdataset_alunosMAE.Value);
         if campos_contrato('@valor@') then
            Doc.Content.Find.Execute(FindText:='@valor@', ReplaceWith:= formatfloat('##,##0.00',vmensalidade));
         if campos_contrato('@celular_aluno@') then
            Doc.Content.Find.Execute(FindText:='@celular_aluno@', ReplaceWith:=ibdataset_alunosCELULAR.Value);
         if campos_contrato('@endereco_aluno@') then
            Doc.Content.Find.Execute(FindText:='@endereco_aluno@', ReplaceWith:=ibdataset_alunosCELULAR.Value);
         if campos_contrato('@curso_aluno@') then
            Doc.Content.Find.Execute(FindText:='@curso_aluno@', ReplaceWith:=ibdataset_matriculasCURSO.Value);
         if campos_contrato('@serie_aluno@') then
            Doc.Content.Find.Execute(FindText:='@serie_aluno@', ReplaceWith:=ibdataset_matriculasSERIE.Value);
         if campos_contrato('@turma_aluno@') then
            Doc.Content.Find.Execute(FindText:='@turma_aluno@', ReplaceWith:=ibdataset_matriculasTURMA.Value);
         if campos_contrato('@turno@') then
            Doc.Content.Find.Execute(FindText:='@turno@', ReplaceWith:=ibdataset_matriculasTURMA.Value);
         // outros
         if campos_contrato('@ano_letivo@') then
            Doc.Content.Find.Execute(FindText:='@ano_letivo@', ReplaceWith:=inttostr(principalu.vano_letivo));
         if campos_contrato('@responsavel_aluno@') then begin
            if ibdataset_alunosRESP_PREFERENCIAL.Value='P' then
               Doc.Content.Find.Execute(FindText:='@responsavel_aluno@', ReplaceWith:=ibdataset_alunosNOME_RESPONSAVEL01.Value)
            else
            begin
               if ibdataset_alunosRESP_PREFERENCIAL.Value='M' then
                  Doc.Content.Find.Execute(FindText:='@responsavel_aluno@', ReplaceWith:=ibdataset_alunosNOME_RESPONSAVEL02.Value)
               else
                  Doc.Content.Find.Execute(FindText:='@responsavel_aluno@', ReplaceWith:=ibdataset_alunosNOME_RESPONSAVEL03.Value)
            end;
         end;
         if campos_contrato('@valor_total@') then
            Doc.Content.Find.Execute(FindText:='@valor_total@', ReplaceWith:=formatfloat('###,##0.00',vvalor));
         if campos_contrato('@valor_por_extenso@') then begin
            if vvalor>0 then
               Doc.Content.Find.Execute(FindText:='@valor_por_extenso@', ReplaceWith:=extenso(vvalor))
            else
               Doc.Content.Find.Execute(FindText:='@valor_por_extenso@', ReplaceWith:='');
         end;
         if campos_contrato('@parcelas@') then
            Doc.Content.Find.Execute(FindText:='@parcelas@', ReplaceWith:=inttostr(vparcelas));
         if campos_contrato('@parcelas_extenso@') then
            Doc.Content.Find.Execute(FindText:='@parcelas_extenso@', ReplaceWith:= numeroextenso1(vparcelas));
         if campos_contrato('@entrada@') then begin
            Doc.Content.Find.Execute(FindText:='@entrada@', ReplaceWith:= formatfloat('###,##0.00',ventrada));
            if ventrada>0 then
               Doc.Content.Find.Execute(FindText:='@entrada_extenso@', ReplaceWith:=extenso(ventrada))
            else
               Doc.Content.Find.Execute(FindText:='@entrada_extenso@', ReplaceWith:='');
         end;
         if campos_contrato('@parcelas_restante@') then
            Doc.Content.Find.Execute(FindText:='@parcelas_restante@', ReplaceWith:= inttostr(vparcelas-1));
         if campos_contrato('@restante_extenso@') then begin
            if vparcelas-1>0 then
               Doc.Content.Find.Execute(FindText:='@restante_extenso@', ReplaceWith:= numeroextenso1(vparcelas-1))
            else
               Doc.Content.Find.Execute(FindText:='@restante_extenso@', ReplaceWith:= '');
         end;
         if campos_contrato('@dia_vencimento@') then
            Doc.Content.Find.Execute(FindText:='@dia_vencimento@', ReplaceWith:= inttostr(ibdataset_matriculasDIA_MENSALIDADE.Value));
         if campos_contrato('@juros_dia@') then
            Doc.Content.Find.Execute(FindText:='@juros_dia@', ReplaceWith:= formatfloat('##0.00',ibdataset_parametrosJUROS.Value)+'%');
         if campos_contrato('@multa@') then
            Doc.Content.Find.Execute(FindText:='@multa@', ReplaceWith:= formatfloat('##0.00',ibdataset_parametrosMULTA.Value)+'%');
         if campos_contrato('@cidade_data@') then
            Doc.Content.Find.Execute(FindText:='@cidade_data@', ReplaceWith:= principalu.cidemp);
         if campos_contrato('@data_extenso@') then
            Doc.Content.Find.Execute(FindText:='@data_extenso@', ReplaceWith:= dataporextenso(ibdataset_matriculasDATA_MATRICULA.Value));
         if campos_contrato('@data_hoje@') then begin
            if data_hoje>strtodate('01/01/1910') then
               Doc.Content.Find.Execute(FindText:='@data_hoje@', ReplaceWith:= dataporextenso(data_hoje))
            else
               Doc.Content.Find.Execute(FindText:='@data_hoje@', ReplaceWith:= dataporextenso(date()));
         end;
         //salvar
      end;
   end
   else
      messagedlg('Arquivo '+varqv1+' não encontrado',mtinformation,[mbok],0);
   result:=true;
end;

function ano_letivo():integer;
var
vano: integer;

begin
   with datamodule1 do begin
      ibdataset_ano_letivo.Close;
      ibdataset_ano_letivo.SelectSQL.Clear;
      ibdataset_ano_letivo.SelectSQL.Add('Select * from ano_letivo order by ano');
      ibdataset_ano_letivo.Open;
      vano:=0;
      while ibdataset_ano_letivo.eof=false do begin
         if (ibdataset_ano_letivoFECHADO.Value='N') and
            (vano=0) then
            vano:=ibdataset_ano_letivoANO.Value;
         ibdataset_ano_letivo.Next;
      end;
      if vano=0 then vano:=strtointdef(copy(datetostr(date()),7,4),0);
   end;
   result:=vano;
end;

function envia_email(mensagem, email, de, email_dest, assunto, email_cc, host,
   username, passw, anexo: string; porta: integer): boolean;
var xanexo: integer;

begin
with principal do begin
  memomensagem.Text:=mensagem;
  listboxanexos.Items.Clear;
  if alltrim(anexo)<>'' then
     listboxanexos.Items.Add(anexo);
  msg.body.Assign(memomensagem.Lines); // corpo da mensagem
  msg.From.Text := email; // e-mail de origem
  msg.From.Name := de; //nome que aparecerá no provedor quando o destinatário verificar o e-mail
  msg.Recipients.EMailAddresses := email_cc;// endereço que também receberá o e-mail;
  msg.Subject:= assunto; // assunto do e-mail
  msg.CCList.EMailAddresses := email_dest; //e-mail destinatário

  idsmtp1.Host := host;// seu provedor ex: terra
  idsmtp1.Port := porta;  // 25; //porta padrão para o envio de e-mail (SMTP) - Não mudar
  idsmtp1.Username := username;// Seu username
  idsmtp1.Password := passw;//Sua senha

  //Prioridade da mensagem
  msg.Priority := mpHigh;  // Alta
   //Envio de arquivos anexos
  for xAnexo := 0 to ListBoxanexos.Items.Count -1 do
  tidattachment.Create(msg.MessageParts , Tfilename(listboxanexos.Items.Strings [xanexo]));
    idsmtp1.Connect;
  try
    idsmtp1.Send(msg);
  finally
    idsmtp1.Disconnect;
  end;
//  Application.MessageBox('Email enviado com sucesso!', 'Confirmação', MB_ICONINFORMATION + MB_OK);
  result:=true;
end;
end;

function nota(Cod_Aluno, Disciplina: integer; Curso: String; Unidade: integer): double;
var
vnota: double;

begin
   with datamodule1 do begin
      vnota:=0;
      if disciplina=1 then begin
         ibdataset_Disciplina01.Close;
         ibdataset_Disciplina01.SelectSQL.Clear;
         ibdataset_Disciplina01.SelectSQL.Add('Select * from Disciplina01 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina01.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina01.Params[1].AsString:=Curso;
         ibdataset_Disciplina01.Open;
         if ibdataset_disciplina01.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina01UNI1AVA1.Value+ibdataset_disciplina01UNI1AVA2.Value+
                      ibdataset_disciplina01UNI1AVA3.Value+ibdataset_disciplina01UNI1AVA4.Value+
                      ibdataset_disciplina01UNI1AVA5.Value+ibdataset_disciplina01UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina01UNI2AVA1.Value+ibdataset_disciplina01UNI2AVA2.Value+
                      ibdataset_disciplina01UNI2AVA3.Value+ibdataset_disciplina01UNI2AVA4.Value+
                      ibdataset_disciplina01UNI2AVA5.Value+ibdataset_disciplina01UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina01UNI3AVA1.Value+ibdataset_disciplina01UNI3AVA2.Value+
                      ibdataset_disciplina01UNI3AVA3.Value+ibdataset_disciplina01UNI3AVA4.Value+
                      ibdataset_disciplina01UNI3AVA5.Value+ibdataset_disciplina01UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina01UNI4AVA1.Value+ibdataset_disciplina01UNI4AVA2.Value+
                      ibdataset_disciplina01UNI4AVA3.Value+ibdataset_disciplina01UNI4AVA4.Value+
                      ibdataset_disciplina01UNI4AVA5.Value+ibdataset_disciplina01UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=2 then begin
         ibdataset_Disciplina02.Close;
         ibdataset_Disciplina02.SelectSQL.Clear;
         ibdataset_Disciplina02.SelectSQL.Add('Select * from Disciplina02 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina02.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina02.Params[1].AsString:=Curso;
         ibdataset_Disciplina02.Open;
         if ibdataset_disciplina02.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina02UNI1AVA1.Value+ibdataset_disciplina02UNI1AVA2.Value+
                      ibdataset_disciplina02UNI1AVA3.Value+ibdataset_disciplina02UNI1AVA4.Value+
                      ibdataset_disciplina02UNI1AVA5.Value+ibdataset_disciplina02UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina02UNI2AVA1.Value+ibdataset_disciplina02UNI2AVA2.Value+
                      ibdataset_disciplina02UNI2AVA3.Value+ibdataset_disciplina02UNI2AVA4.Value+
                      ibdataset_disciplina02UNI2AVA5.Value+ibdataset_disciplina02UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina02UNI3AVA1.Value+ibdataset_disciplina02UNI3AVA2.Value+
                      ibdataset_disciplina02UNI3AVA3.Value+ibdataset_disciplina02UNI3AVA4.Value+
                      ibdataset_disciplina02UNI3AVA5.Value+ibdataset_disciplina02UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina02UNI4AVA1.Value+ibdataset_disciplina02UNI4AVA2.Value+
                      ibdataset_disciplina02UNI4AVA3.Value+ibdataset_disciplina02UNI4AVA4.Value+
                      ibdataset_disciplina02UNI4AVA5.Value+ibdataset_disciplina02UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=3 then begin
         ibdataset_Disciplina03.Close;
         ibdataset_Disciplina03.SelectSQL.Clear;
         ibdataset_Disciplina03.SelectSQL.Add('Select * from Disciplina03 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina03.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina03.Params[1].AsString:=Curso;
         ibdataset_Disciplina03.Open;
         if ibdataset_disciplina03.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina03UNI1AVA1.Value+ibdataset_disciplina03UNI1AVA2.Value+
                      ibdataset_disciplina03UNI1AVA3.Value+ibdataset_disciplina03UNI1AVA4.Value+
                      ibdataset_disciplina03UNI1AVA5.Value+ibdataset_disciplina03UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina03UNI2AVA1.Value+ibdataset_disciplina03UNI2AVA2.Value+
                      ibdataset_disciplina03UNI2AVA3.Value+ibdataset_disciplina03UNI2AVA4.Value+
                      ibdataset_disciplina03UNI2AVA5.Value+ibdataset_disciplina03UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina03UNI3AVA1.Value+ibdataset_disciplina03UNI3AVA2.Value+
                      ibdataset_disciplina03UNI3AVA3.Value+ibdataset_disciplina03UNI3AVA4.Value+
                      ibdataset_disciplina03UNI3AVA5.Value+ibdataset_disciplina03UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina03UNI4AVA1.Value+ibdataset_disciplina03UNI4AVA2.Value+
                      ibdataset_disciplina03UNI4AVA3.Value+ibdataset_disciplina03UNI4AVA4.Value+
                      ibdataset_disciplina03UNI4AVA5.Value+ibdataset_disciplina03UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=4 then begin
         ibdataset_Disciplina04.Close;
         ibdataset_Disciplina04.SelectSQL.Clear;
         ibdataset_Disciplina04.SelectSQL.Add('Select * from Disciplina04 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina04.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina04.Params[1].AsString:=Curso;
         ibdataset_Disciplina04.Open;
         if ibdataset_disciplina04.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina04UNI1AVA1.Value+ibdataset_disciplina04UNI1AVA2.Value+
                      ibdataset_disciplina04UNI1AVA3.Value+ibdataset_disciplina04UNI1AVA4.Value+
                      ibdataset_disciplina04UNI1AVA5.Value+ibdataset_disciplina04UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina04UNI2AVA1.Value+ibdataset_disciplina04UNI2AVA2.Value+
                      ibdataset_disciplina04UNI2AVA3.Value+ibdataset_disciplina04UNI2AVA4.Value+
                      ibdataset_disciplina04UNI2AVA5.Value+ibdataset_disciplina04UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina04UNI3AVA1.Value+ibdataset_disciplina04UNI3AVA2.Value+
                      ibdataset_disciplina04UNI3AVA3.Value+ibdataset_disciplina04UNI3AVA4.Value+
                      ibdataset_disciplina04UNI3AVA5.Value+ibdataset_disciplina04UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina04UNI4AVA1.Value+ibdataset_disciplina04UNI4AVA2.Value+
                      ibdataset_disciplina04UNI4AVA3.Value+ibdataset_disciplina04UNI4AVA4.Value+
                      ibdataset_disciplina04UNI4AVA5.Value+ibdataset_disciplina04UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=5 then begin
         ibdataset_Disciplina05.Close;
         ibdataset_Disciplina05.SelectSQL.Clear;
         ibdataset_Disciplina05.SelectSQL.Add('Select * from Disciplina05 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina05.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina05.Params[1].AsString:=Curso;
         ibdataset_Disciplina05.Open;
         if ibdataset_disciplina05.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina05UNI1AVA1.Value+ibdataset_disciplina05UNI1AVA2.Value+
                      ibdataset_disciplina05UNI1AVA3.Value+ibdataset_disciplina05UNI1AVA4.Value+
                      ibdataset_disciplina05UNI1AVA5.Value+ibdataset_disciplina05UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina05UNI2AVA1.Value+ibdataset_disciplina05UNI2AVA2.Value+
                      ibdataset_disciplina05UNI2AVA3.Value+ibdataset_disciplina05UNI2AVA4.Value+
                      ibdataset_disciplina05UNI2AVA5.Value+ibdataset_disciplina05UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina05UNI3AVA1.Value+ibdataset_disciplina05UNI3AVA2.Value+
                      ibdataset_disciplina05UNI3AVA3.Value+ibdataset_disciplina05UNI3AVA4.Value+
                      ibdataset_disciplina05UNI3AVA5.Value+ibdataset_disciplina05UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina05UNI4AVA1.Value+ibdataset_disciplina05UNI4AVA2.Value+
                      ibdataset_disciplina05UNI4AVA3.Value+ibdataset_disciplina05UNI4AVA4.Value+
                      ibdataset_disciplina05UNI4AVA5.Value+ibdataset_disciplina05UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=6 then begin
         ibdataset_Disciplina06.Close;
         ibdataset_Disciplina06.SelectSQL.Clear;
         ibdataset_Disciplina06.SelectSQL.Add('Select * from Disciplina06 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina06.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina06.Params[1].AsString:=Curso;
         ibdataset_Disciplina06.Open;
         if ibdataset_disciplina06.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina06UNI1AVA1.Value+ibdataset_disciplina06UNI1AVA2.Value+
                      ibdataset_disciplina06UNI1AVA3.Value+ibdataset_disciplina06UNI1AVA4.Value+
                      ibdataset_disciplina06UNI1AVA5.Value+ibdataset_disciplina06UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina06UNI2AVA1.Value+ibdataset_disciplina06UNI2AVA2.Value+
                      ibdataset_disciplina06UNI2AVA3.Value+ibdataset_disciplina06UNI2AVA4.Value+
                      ibdataset_disciplina06UNI2AVA5.Value+ibdataset_disciplina06UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina06UNI3AVA1.Value+ibdataset_disciplina06UNI3AVA2.Value+
                      ibdataset_disciplina06UNI3AVA3.Value+ibdataset_disciplina06UNI3AVA4.Value+
                      ibdataset_disciplina06UNI3AVA5.Value+ibdataset_disciplina06UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina06UNI4AVA1.Value+ibdataset_disciplina06UNI4AVA2.Value+
                      ibdataset_disciplina06UNI4AVA3.Value+ibdataset_disciplina06UNI4AVA4.Value+
                      ibdataset_disciplina06UNI4AVA5.Value+ibdataset_disciplina06UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=7 then begin
         ibdataset_Disciplina07.Close;
         ibdataset_Disciplina07.SelectSQL.Clear;
         ibdataset_Disciplina07.SelectSQL.Add('Select * from Disciplina07 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina07.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina07.Params[1].AsString:=Curso;
         ibdataset_Disciplina07.Open;
         if ibdataset_disciplina07.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina07UNI1AVA1.Value+ibdataset_disciplina07UNI1AVA2.Value+
                      ibdataset_disciplina07UNI1AVA3.Value+ibdataset_disciplina07UNI1AVA4.Value+
                      ibdataset_disciplina07UNI1AVA5.Value+ibdataset_disciplina07UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina07UNI2AVA1.Value+ibdataset_disciplina07UNI2AVA2.Value+
                      ibdataset_disciplina07UNI2AVA3.Value+ibdataset_disciplina07UNI2AVA4.Value+
                      ibdataset_disciplina07UNI2AVA5.Value+ibdataset_disciplina07UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina07UNI3AVA1.Value+ibdataset_disciplina07UNI3AVA2.Value+
                      ibdataset_disciplina07UNI3AVA3.Value+ibdataset_disciplina07UNI3AVA4.Value+
                      ibdataset_disciplina07UNI3AVA5.Value+ibdataset_disciplina07UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina07UNI4AVA1.Value+ibdataset_disciplina07UNI4AVA2.Value+
                      ibdataset_disciplina07UNI4AVA3.Value+ibdataset_disciplina07UNI4AVA4.Value+
                      ibdataset_disciplina07UNI4AVA5.Value+ibdataset_disciplina07UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=8 then begin
         ibdataset_Disciplina08.Close;
         ibdataset_Disciplina08.SelectSQL.Clear;
         ibdataset_Disciplina08.SelectSQL.Add('Select * from Disciplina08 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina08.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina08.Params[1].AsString:=Curso;
         ibdataset_Disciplina08.Open;
         if ibdataset_disciplina08.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina08UNI1AVA1.Value+ibdataset_disciplina08UNI1AVA2.Value+
                      ibdataset_disciplina08UNI1AVA3.Value+ibdataset_disciplina08UNI1AVA4.Value+
                      ibdataset_disciplina08UNI1AVA5.Value+ibdataset_disciplina08UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina08UNI2AVA1.Value+ibdataset_disciplina08UNI2AVA2.Value+
                      ibdataset_disciplina08UNI2AVA3.Value+ibdataset_disciplina08UNI2AVA4.Value+
                      ibdataset_disciplina08UNI2AVA5.Value+ibdataset_disciplina08UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina08UNI3AVA1.Value+ibdataset_disciplina08UNI3AVA2.Value+
                      ibdataset_disciplina08UNI3AVA3.Value+ibdataset_disciplina08UNI3AVA4.Value+
                      ibdataset_disciplina08UNI3AVA5.Value+ibdataset_disciplina08UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina08UNI4AVA1.Value+ibdataset_disciplina08UNI4AVA2.Value+
                      ibdataset_disciplina08UNI4AVA3.Value+ibdataset_disciplina08UNI4AVA4.Value+
                      ibdataset_disciplina08UNI4AVA5.Value+ibdataset_disciplina08UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=9 then begin
         ibdataset_Disciplina09.Close;
         ibdataset_Disciplina09.SelectSQL.Clear;
         ibdataset_Disciplina09.SelectSQL.Add('Select * from Disciplina09 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina09.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina09.Params[1].AsString:=Curso;
         ibdataset_Disciplina09.Open;
         if ibdataset_disciplina09.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina09UNI1AVA1.Value+ibdataset_disciplina09UNI1AVA2.Value+
                      ibdataset_disciplina09UNI1AVA3.Value+ibdataset_disciplina09UNI1AVA4.Value+
                      ibdataset_disciplina09UNI1AVA5.Value+ibdataset_disciplina09UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina09UNI2AVA1.Value+ibdataset_disciplina09UNI2AVA2.Value+
                      ibdataset_disciplina09UNI2AVA3.Value+ibdataset_disciplina09UNI2AVA4.Value+
                      ibdataset_disciplina09UNI2AVA5.Value+ibdataset_disciplina09UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina09UNI3AVA1.Value+ibdataset_disciplina09UNI3AVA2.Value+
                      ibdataset_disciplina09UNI3AVA3.Value+ibdataset_disciplina09UNI3AVA4.Value+
                      ibdataset_disciplina09UNI3AVA5.Value+ibdataset_disciplina09UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina09UNI4AVA1.Value+ibdataset_disciplina09UNI4AVA2.Value+
                      ibdataset_disciplina09UNI4AVA3.Value+ibdataset_disciplina09UNI4AVA4.Value+
                      ibdataset_disciplina09UNI4AVA5.Value+ibdataset_disciplina09UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=10 then begin
         ibdataset_Disciplina10.Close;
         ibdataset_Disciplina10.SelectSQL.Clear;
         ibdataset_Disciplina10.SelectSQL.Add('Select * from Disciplina10 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina10.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina10.Params[1].AsString:=Curso;
         ibdataset_Disciplina10.Open;
         if ibdataset_disciplina10.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina10UNI1AVA1.Value+ibdataset_disciplina10UNI1AVA2.Value+
                      ibdataset_disciplina10UNI1AVA3.Value+ibdataset_disciplina10UNI1AVA4.Value+
                      ibdataset_disciplina10UNI1AVA5.Value+ibdataset_disciplina10UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina10UNI2AVA1.Value+ibdataset_disciplina10UNI2AVA2.Value+
                      ibdataset_disciplina10UNI2AVA3.Value+ibdataset_disciplina10UNI2AVA4.Value+
                      ibdataset_disciplina10UNI2AVA5.Value+ibdataset_disciplina10UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina10UNI3AVA1.Value+ibdataset_disciplina10UNI3AVA2.Value+
                      ibdataset_disciplina10UNI3AVA3.Value+ibdataset_disciplina10UNI3AVA4.Value+
                      ibdataset_disciplina10UNI3AVA5.Value+ibdataset_disciplina10UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina10UNI4AVA1.Value+ibdataset_disciplina10UNI4AVA2.Value+
                      ibdataset_disciplina10UNI4AVA3.Value+ibdataset_disciplina10UNI4AVA4.Value+
                      ibdataset_disciplina10UNI4AVA5.Value+ibdataset_disciplina10UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=11 then begin
         ibdataset_Disciplina11.Close;
         ibdataset_Disciplina11.SelectSQL.Clear;
         ibdataset_Disciplina11.SelectSQL.Add('Select * from Disciplina11 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina11.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina11.Params[1].AsString:=Curso;
         ibdataset_Disciplina11.Open;
         if ibdataset_disciplina11.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina11UNI1AVA1.Value+ibdataset_disciplina11UNI1AVA2.Value+
                      ibdataset_disciplina11UNI1AVA3.Value+ibdataset_disciplina11UNI1AVA4.Value+
                      ibdataset_disciplina11UNI1AVA5.Value+ibdataset_disciplina11UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina11UNI2AVA1.Value+ibdataset_disciplina11UNI2AVA2.Value+
                      ibdataset_disciplina11UNI2AVA3.Value+ibdataset_disciplina11UNI2AVA4.Value+
                      ibdataset_disciplina11UNI2AVA5.Value+ibdataset_disciplina11UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina11UNI3AVA1.Value+ibdataset_disciplina11UNI3AVA2.Value+
                      ibdataset_disciplina11UNI3AVA3.Value+ibdataset_disciplina11UNI3AVA4.Value+
                      ibdataset_disciplina11UNI3AVA5.Value+ibdataset_disciplina11UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina11UNI4AVA1.Value+ibdataset_disciplina11UNI4AVA2.Value+
                      ibdataset_disciplina11UNI4AVA3.Value+ibdataset_disciplina11UNI4AVA4.Value+
                      ibdataset_disciplina11UNI4AVA5.Value+ibdataset_disciplina11UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=12 then begin
         ibdataset_Disciplina12.Close;
         ibdataset_Disciplina12.SelectSQL.Clear;
         ibdataset_Disciplina12.SelectSQL.Add('Select * from Disciplina12 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina12.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina12.Params[1].AsString:=Curso;
         ibdataset_Disciplina12.Open;
         if ibdataset_disciplina12.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina12UNI1AVA1.Value+ibdataset_disciplina12UNI1AVA2.Value+
                      ibdataset_disciplina12UNI1AVA3.Value+ibdataset_disciplina12UNI1AVA4.Value+
                      ibdataset_disciplina12UNI1AVA5.Value+ibdataset_disciplina12UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina12UNI2AVA1.Value+ibdataset_disciplina12UNI2AVA2.Value+
                      ibdataset_disciplina12UNI2AVA3.Value+ibdataset_disciplina12UNI2AVA4.Value+
                      ibdataset_disciplina12UNI2AVA5.Value+ibdataset_disciplina12UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina12UNI3AVA1.Value+ibdataset_disciplina12UNI3AVA2.Value+
                      ibdataset_disciplina12UNI3AVA3.Value+ibdataset_disciplina12UNI3AVA4.Value+
                      ibdataset_disciplina12UNI3AVA5.Value+ibdataset_disciplina12UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina12UNI4AVA1.Value+ibdataset_disciplina12UNI4AVA2.Value+
                      ibdataset_disciplina12UNI4AVA3.Value+ibdataset_disciplina12UNI4AVA4.Value+
                      ibdataset_disciplina12UNI4AVA5.Value+ibdataset_disciplina12UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=13 then begin
         ibdataset_Disciplina13.Close;
         ibdataset_Disciplina13.SelectSQL.Clear;
         ibdataset_Disciplina13.SelectSQL.Add('Select * from Disciplina13 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina13.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina13.Params[1].AsString:=Curso;
         ibdataset_Disciplina13.Open;
         if ibdataset_disciplina13.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina13UNI1AVA1.Value+ibdataset_disciplina13UNI1AVA2.Value+
                      ibdataset_disciplina13UNI1AVA3.Value+ibdataset_disciplina13UNI1AVA4.Value+
                      ibdataset_disciplina13UNI1AVA5.Value+ibdataset_disciplina13UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina13UNI2AVA1.Value+ibdataset_disciplina13UNI2AVA2.Value+
                      ibdataset_disciplina13UNI2AVA3.Value+ibdataset_disciplina13UNI2AVA4.Value+
                      ibdataset_disciplina13UNI2AVA5.Value+ibdataset_disciplina13UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina13UNI3AVA1.Value+ibdataset_disciplina13UNI3AVA2.Value+
                      ibdataset_disciplina13UNI3AVA3.Value+ibdataset_disciplina13UNI3AVA4.Value+
                      ibdataset_disciplina13UNI3AVA5.Value+ibdataset_disciplina13UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina13UNI4AVA1.Value+ibdataset_disciplina13UNI4AVA2.Value+
                      ibdataset_disciplina13UNI4AVA3.Value+ibdataset_disciplina13UNI4AVA4.Value+
                      ibdataset_disciplina13UNI4AVA5.Value+ibdataset_disciplina13UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=14 then begin
         ibdataset_Disciplina14.Close;
         ibdataset_Disciplina14.SelectSQL.Clear;
         ibdataset_Disciplina14.SelectSQL.Add('Select * from Disciplina14 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina14.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina14.Params[1].AsString:=Curso;
         ibdataset_Disciplina14.Open;
         if ibdataset_disciplina14.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina14UNI1AVA1.Value+ibdataset_disciplina14UNI1AVA2.Value+
                      ibdataset_disciplina14UNI1AVA3.Value+ibdataset_disciplina14UNI1AVA4.Value+
                      ibdataset_disciplina14UNI1AVA5.Value+ibdataset_disciplina14UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina14UNI2AVA1.Value+ibdataset_disciplina14UNI2AVA2.Value+
                      ibdataset_disciplina14UNI2AVA3.Value+ibdataset_disciplina14UNI2AVA4.Value+
                      ibdataset_disciplina14UNI2AVA5.Value+ibdataset_disciplina14UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina14UNI3AVA1.Value+ibdataset_disciplina14UNI3AVA2.Value+
                      ibdataset_disciplina14UNI3AVA3.Value+ibdataset_disciplina14UNI3AVA4.Value+
                      ibdataset_disciplina14UNI3AVA5.Value+ibdataset_disciplina14UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina14UNI4AVA1.Value+ibdataset_disciplina14UNI4AVA2.Value+
                      ibdataset_disciplina14UNI4AVA3.Value+ibdataset_disciplina14UNI4AVA4.Value+
                      ibdataset_disciplina14UNI4AVA5.Value+ibdataset_disciplina14UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      //
      if disciplina=15 then begin
         ibdataset_Disciplina15.Close;
         ibdataset_Disciplina15.SelectSQL.Clear;
         ibdataset_Disciplina15.SelectSQL.Add('Select * from Disciplina15 where cd_aluno=:cd_aluno and curso=:curso');
         ibdataset_Disciplina15.Params[0].AsInteger:=Cod_Aluno;
         ibdataset_Disciplina15.Params[1].AsString:=Curso;
         ibdataset_Disciplina15.Open;
         if ibdataset_disciplina15.RecordCount>0 then begin
            if unidade=1 then
               vnota:=ibdataset_disciplina15UNI1AVA1.Value+ibdataset_disciplina15UNI1AVA2.Value+
                      ibdataset_disciplina15UNI1AVA3.Value+ibdataset_disciplina15UNI1AVA4.Value+
                      ibdataset_disciplina15UNI1AVA5.Value+ibdataset_disciplina15UNI1PROVA.Value;
            if unidade=2 then
               vnota:=ibdataset_disciplina15UNI2AVA1.Value+ibdataset_disciplina15UNI2AVA2.Value+
                      ibdataset_disciplina15UNI2AVA3.Value+ibdataset_disciplina15UNI2AVA4.Value+
                      ibdataset_disciplina15UNI2AVA5.Value+ibdataset_disciplina15UNI2PROVA.Value;
            if unidade=3 then
               vnota:=ibdataset_disciplina15UNI3AVA1.Value+ibdataset_disciplina15UNI3AVA2.Value+
                      ibdataset_disciplina15UNI3AVA3.Value+ibdataset_disciplina15UNI3AVA4.Value+
                      ibdataset_disciplina15UNI3AVA5.Value+ibdataset_disciplina15UNI3PROVA.Value;
            if unidade=4 then
               vnota:=ibdataset_disciplina15UNI4AVA1.Value+ibdataset_disciplina15UNI4AVA2.Value+
                      ibdataset_disciplina15UNI4AVA3.Value+ibdataset_disciplina15UNI4AVA4.Value+
                      ibdataset_disciplina15UNI4AVA5.Value+ibdataset_disciplina15UNI4PROVA.Value;
            result:=vnota;
         end;
      end;
      result:=vnota;
   end;
end;

function disciplina(Matricula, Disciplina: integer): String;
begin
   with datamodule1 do begin
      if disciplina=1 then begin
         ibdataset_Disciplina01.Close;
         ibdataset_Disciplina01.SelectSQL.Clear;
         ibdataset_Disciplina01.SelectSQL.Add('Select * from Disciplina01 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina01.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina01.Open;
         if ibdataset_disciplina01.RecordCount>0 then
            result:=ibdataset_disciplina01DISCIPLINA.Value;
      end;
      //
      if disciplina=2 then begin
         ibdataset_Disciplina02.Close;
         ibdataset_Disciplina02.SelectSQL.Clear;
         ibdataset_Disciplina02.SelectSQL.Add('Select * from Disciplina02 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina02.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina02.Open;
         if ibdataset_disciplina02.RecordCount>0 then
            result:=ibdataset_disciplina02DISCIPLINA.Value;
      end;
      //
      if disciplina=3 then begin
         ibdataset_Disciplina03.Close;
         ibdataset_Disciplina03.SelectSQL.Clear;
         ibdataset_Disciplina03.SelectSQL.Add('Select * from Disciplina03 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina03.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina03.Open;
         if ibdataset_disciplina03.RecordCount>0 then
            result:=ibdataset_disciplina03DISCIPLINA.Value;
      end;
      //
      if disciplina=4 then begin
         ibdataset_Disciplina04.Close;
         ibdataset_Disciplina04.SelectSQL.Clear;
         ibdataset_Disciplina04.SelectSQL.Add('Select * from Disciplina04 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina04.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina04.Open;
         if ibdataset_disciplina04.RecordCount>0 then
            result:=ibdataset_disciplina04DISCIPLINA.Value;
      end;
      //
      if disciplina=5 then begin
         ibdataset_Disciplina05.Close;
         ibdataset_Disciplina05.SelectSQL.Clear;
         ibdataset_Disciplina05.SelectSQL.Add('Select * from Disciplina05 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina05.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina05.Open;
         if ibdataset_disciplina05.RecordCount>0 then
            result:=ibdataset_disciplina05DISCIPLINA.Value;
      end;
      //
      if disciplina=6 then begin
         ibdataset_Disciplina06.Close;
         ibdataset_Disciplina06.SelectSQL.Clear;
         ibdataset_Disciplina06.SelectSQL.Add('Select * from Disciplina06 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina06.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina06.Open;
         if ibdataset_disciplina06.RecordCount>0 then
            result:=ibdataset_disciplina06DISCIPLINA.Value;
      end;
      //
      if disciplina=7 then begin
         ibdataset_Disciplina07.Close;
         ibdataset_Disciplina07.SelectSQL.Clear;
         ibdataset_Disciplina07.SelectSQL.Add('Select * from Disciplina07 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina07.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina07.Open;
         if ibdataset_disciplina07.RecordCount>0 then
            result:=ibdataset_disciplina07DISCIPLINA.Value;
      end;
      //
      if disciplina=8 then begin
         ibdataset_Disciplina08.Close;
         ibdataset_Disciplina08.SelectSQL.Clear;
         ibdataset_Disciplina08.SelectSQL.Add('Select * from Disciplina08 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina08.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina08.Open;
         if ibdataset_disciplina08.RecordCount>0 then
            result:=ibdataset_disciplina08DISCIPLINA.Value;
      end;
      //
      if disciplina=9 then begin
         ibdataset_Disciplina09.Close;
         ibdataset_Disciplina09.SelectSQL.Clear;
         ibdataset_Disciplina09.SelectSQL.Add('Select * from Disciplina09 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina09.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina09.Open;
         if ibdataset_disciplina09.RecordCount>0 then
            result:=ibdataset_disciplina09DISCIPLINA.Value;
      end;
      //
      if disciplina=10 then begin
         ibdataset_Disciplina10.Close;
         ibdataset_Disciplina10.SelectSQL.Clear;
         ibdataset_Disciplina10.SelectSQL.Add('Select * from Disciplina10 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina10.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina10.Open;
         if ibdataset_disciplina10.RecordCount>0 then
            result:=ibdataset_disciplina10DISCIPLINA.Value;
      end;
      //
      if disciplina=11 then begin
         ibdataset_Disciplina11.Close;
         ibdataset_Disciplina11.SelectSQL.Clear;
         ibdataset_Disciplina11.SelectSQL.Add('Select * from Disciplina11 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina11.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina11.Open;
         if ibdataset_disciplina11.RecordCount>0 then
            result:=ibdataset_disciplina11DISCIPLINA.Value;
      end;
      //
      if disciplina=12 then begin
         ibdataset_Disciplina12.Close;
         ibdataset_Disciplina12.SelectSQL.Clear;
         ibdataset_Disciplina12.SelectSQL.Add('Select * from Disciplina12 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina12.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina12.Open;
         if ibdataset_disciplina12.RecordCount>0 then
            result:=ibdataset_disciplina12DISCIPLINA.Value;
      end;
      //
      if disciplina=13 then begin
         ibdataset_Disciplina13.Close;
         ibdataset_Disciplina13.SelectSQL.Clear;
         ibdataset_Disciplina13.SelectSQL.Add('Select * from Disciplina13 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina13.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina13.Open;
         if ibdataset_disciplina13.RecordCount>0 then
            result:=ibdataset_disciplina13DISCIPLINA.Value;
      end;
      //
      if disciplina=14 then begin
         ibdataset_Disciplina14.Close;
         ibdataset_Disciplina14.SelectSQL.Clear;
         ibdataset_Disciplina14.SelectSQL.Add('Select * from Disciplina14 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina14.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina14.Open;
         if ibdataset_disciplina14.RecordCount>0 then
            result:=ibdataset_disciplina14DISCIPLINA.Value;
      end;
      //
      if disciplina=15 then begin
         ibdataset_Disciplina15.Close;
         ibdataset_Disciplina15.SelectSQL.Clear;
         ibdataset_Disciplina15.SelectSQL.Add('Select * from Disciplina15 where N_MATRICULA=:N_MATRICULA');
         ibdataset_Disciplina15.Params[0].AsInteger:=Matricula;
         ibdataset_Disciplina15.Open;
         if ibdataset_disciplina15.RecordCount>0 then
            result:=ibdataset_disciplina15DISCIPLINA.Value;
      end;
   end;
end;

function yWallPapper(Image1:TImage):boolean;
var sair:boolean;
    ds,x,y,l,nm,c,c1:integer;
    d,datprm:tdatetime;
    Bitmap: TBitmap;
    scwidth:integer;
    scheight:integer;
    sem:array[1..7] of string;

begin
// IMAGEM DE FUNDO
     image1.Align:=alClient;
     sem[1]:='D';
     sem[2]:='S';
     sem[3]:='T';
     sem[4]:='Q';
     sem[5]:='Q';
     sem[6]:='S';
     sem[7]:='S';
     scwidth:=screen.width;
     scheight:=screen.height;
     Bitmap := TBitmap.Create;
     BitMap.Monochrome:=false;
     BitMap.canvas.Brush.color:=$ffffff;
     Bitmap.Width := scwidth; // 1024;
     Bitmap.Height := scHeight; //780;
     Image1.Picture.Graphic := Bitmap;
     Image1.Picture.assign(Bitmap);
     Bitmap.free;

//     Image1.align:=alClient;
// NOME DA FIRSTSOFT
     Image1.canvas.font.name:='Arial';
     Image1.Canvas.font.size:=8;
     Image1.Canvas.font.style:=[];
     Image1.Canvas.font.color:=clblack;
//     Image1.Canvas.font.color:=clSilver;
{     Image1.Canvas.Brush.color:=$00E1FFFF;
     Image1.Canvas.TextOut(round(scWidth*0.40),round(scHeight*0.40)+0 ,'           pdigital Informática Ltda.     ');
     Image1.Canvas.TextOut(round(scWidth*0.40),round(scHeight*0.40)+15,'           -------------------------     ');
     Image1.Canvas.TextOut(round(scWidth*0.40),round(scHeight*0.40)+30,'Todos os Direito Autorais Reservados');
     Image1.Canvas.TextOut(round(scWidth*0.40),round(scHeight*0.40)+45,'            Copyright(c) 1989-'+inttostr(yyear(date)));
 }

{     Image1.Canvas.font.size:=6;
     Image1.Canvas.TextOut(round(scWidth*0.05),round(scHeight*0.70)-15,'- Os dados fornecidos pelo sistema e suas respectivas consistências são de inteira responsabilidade do usuário.');
     Image1.Canvas.TextOut(round(scWidth*0.05),round(scHeight*0.70),   '  Caso não concorde com esses termos não o utilize.');
     Image1.Canvas.TextOut(round(scWidth*0.05),round(scHeight*0.70)+15,'Atualização da Versão: '+DateToStr(date())+'   *** Validade técnica 30 DIAS');
 }

// NOME DA EMPRESA NO FUNDO
     Image1.canvas.font.name:='Times New Roman';
     Image1.Canvas.font.size:=10;
     Image1.Canvas.font.style:=[fsBold];
     Image1.Canvas.font.color:=clSilver;
     Image1.Canvas.Brush.color:=$FFFFFF;


// CALENDÁRIO NORMAL
     Image1.canvas.font.name:='Arial';//'Times New Roman';
     Image1.Canvas.font.size:=11;
     Image1.Canvas.font.style:=[fsBold];
     Image1.Canvas.font.color:=clSilver;
     Image1.Canvas.Brush.color:=$FFFFFF;

  {   Impressao do Texto de TITULO }
     l:=5;
     c:=10;
  //   Impressao do Texto de TITULO
//     Image1.Canvas.TextOut(c+2,l,sys_empresa);

  //   Impressao do Texto de TITULO
     l:=80;
     Image1.Canvas.TextOut(c+2,l,ansiuppercase(FormatDateTime('mmmm  dd, yyyy',date)));
     l:=l+30;
     for x:=1 to 7 do
     begin
        Image1.Canvas.TextOut(c+((x-1)*30)+2,l,sem[x]);
     end;

     l:=l+30;
     c:=10;
     y:=yday(ytabd(ymonth(date),yyear(date)));
     for x:=1 to y do
     begin
         d:=strtodate(inttostr(x)+'/'+formatdateTime('mm/yyyy',date));
         ds:=dayofweek(d);
         if (ds=1) and (x<>1) then
         begin
             l:=l+25;
             c:=10;
         end;
         if d=date then
            Image1.Canvas.Brush.color:=clGray
         else
            Image1.Canvas.Brush.color:=$FFFFFF;
         Image1.Canvas.TextOut(c+((ds-1)*30),l,yright(' '+inttostr(x),2));
     end;


  datprm:=date;
  l:=80+5;
  for nm:=1 to 2 do
  begin
//     c1:=80;
     c1:=0+80;
// CALENDÁRIO PROXIMO MES REDUZIDO
     Image1.Canvas.font.name:='Arial'; //'Times New Roman';
     Image1.Canvas.font.size:=7;
     Image1.Canvas.font.style:=[];
     Image1.Canvas.font.color:=clSilver;
     Image1.Canvas.Brush.color:=$FFFFFF;

  {   Impressao do Texto de TITULO }
//   l:=l+30; // 40
//   l:=60;
     datprm:=ytabd(ymonth(datprm),yyear(datprm))+1;
     Image1.Canvas.TextOut((c1+(160*1))+2,l,ansiuppercase(FormatDateTime('mmmm, yyyy',datprm)));
     l:=l+12; //15;
//     c:=0;
     c:=(c1+(160*1));
     for x:=1 to 7 do
     begin
        Image1.Canvas.TextOut(c+((x-1)*15)+2,l,sem[x]);
     end;

     l:=l+12;// 15;//20;
     c:=(c1+(160*1));
     y:=yday(ytabd(ymonth(datprm),yyear(datprm)));
     for x:=1 to y do
     begin
         d:=strtodate(inttostr(x)+'/'+formatdateTime('mm/yyyy',datprm));
         ds:=dayofweek(d);
         if (ds=1) and (x<>1) then
         begin
             l:=l+10;//15;
             c:=(c1+(160*1));
         end;
         Image1.Canvas.TextOut(c+((ds-1)*15),l,yright(' '+inttostr(x),2));
     end;
     l:=l+20; // 40
  end;
end;

function yyear(data:tDateTime):integer;
begin
   yyear:=strtoint(FormatDateTime('yyyy',data));
end;

function yday(data:tDateTime):integer;
begin
   yday:=strtoint(FormatDateTime('dd',data));
end;

// **********************************************************************
//-- Recebe data e devolve o mês (numerico)
function ymonth(data:tDateTime):integer;
begin
   ymonth:=strtoint(FormatDateTime('mm',data));
end;

// **********************************************************************
//-- Recebe MM e AA (numericos) e devolve o ultimo dia do mes (data)
function yTabd ( mes:integer; ano:integer ):tdatetime;
var dd,mm,yy:string;
    dx:array[1..12] of string;
begin
   dx[1] := '31';
   dx[2] := yiif((ano mod 4)=0,'29','28');
   dx[3] := '31';
   dx[4] := '30';
   dx[5] := '31';
   dx[6] := '30';
   dx[7] := '31';
   dx[8] := '31';
   dx[9] := '30';
   dx[7] := '31';
   dx[8] := '31';
   dx[9] := '30';
   dx[10]:= '31';
   dx[11]:= '30';
   dx[12]:= '31';
   dd := dx [ mes ];
   mm := alfzeros(inttostr(mes),2);
   yy := inttostr(ano);
   ytabd:=strtodate(dd+'/'+mm+'/'+yy);
end;

// **********************************************************************
// -- Devolve "n" caracteres de "vr" a partir da direita
function yright(vr:string;n:integer): string;
var
 tm:string;
 x:integer;
 y:integer;
begin
  y:=(length(vr)-n)+1;
  if y<1 then
     y:=1;
  for x:=y to length(vr) do
         tm:=tm+vr[x];
  yright:=tm;
end;

// **********************************************************************
// -- Retorna valor1 se condicao verdadeira ou valor2 se condicao falsa
function yiif(condicao:boolean; valor1:variant; valor2:variant):variant;
begin
   if (condicao) then yiif:=valor1
   else yiif:=valor2;
end;

// **********************************************************************
// -- Completa "vr" com zeros a esquerda até tamanho de vr=n
function alfzeros(vr:string;n:integer): string;
var x:integer;
begin
  if n=0 then n:=length(vr);
  vr:=allbnull(vr);
  for x:=1 to (n-length(vr)) do
    vr:='0'+vr;
  alfzeros:=vr;
end;

// **********************************************************************
// -- Extrai todos espaços em branco de "vr"
function allbnull(vr:string ) : string;
var
   tm:string;
   x:integer;
begin
  for x:=1 to length(vr) do
    if not (vr[x]=' ') then tm:=tm+vr[x];
  allbnull:=tm;
end;

///
///
//
function acesso(Usuario: string; OpcaoMenu: String):boolean;
begin
  result:=false;
  if Usuario<>'-0' then begin
   with datamodule1 do begin
      ibdataset_usuarios.close;
      ibdataset_usuarios.SelectSQL.Clear;
      ibdataset_usuarios.SelectSql.Add('Select * from usuarios where usuario=:usuario');
      ibdataset_usuarios.Params[0].AsString:=usuario;
      ibdataset_usuarios.Open;
      if ibdataset_usuarios.RecordCount=0 then
         messagedlg('Usuário não encontrado',mtinformation,[mbok],0)
      else
      begin
       if ibdataset_usuariosADMINISTRADOR.Value='S' then result:=true else begin
         if Uppercase(OpcaoMenu)='MAN01' then begin
            if ibdataset_usuariosMAN01.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN02' then begin
            if ibdataset_usuariosMAN02.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN03' then begin
            if ibdataset_usuariosMAN03.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN04' then begin
            if ibdataset_usuariosMAN04.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN05' then begin
            if ibdataset_usuariosMAN05.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN06' then begin
            if ibdataset_usuariosMAN06.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN07' then begin
            if ibdataset_usuariosMAN07.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN08' then begin
            if ibdataset_usuariosMAN08.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN09' then begin
            if ibdataset_usuariosMAN09.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN10' then begin
            if ibdataset_usuariosMAN10.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN11' then begin
            if ibdataset_usuariosMAN11.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN12' then begin
            if ibdataset_usuariosMAN12.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN13' then begin
            if ibdataset_usuariosMAN13.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN14' then begin
            if ibdataset_usuariosMAN14.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN15' then begin
            if ibdataset_usuariosMAN15.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN16' then begin
            if ibdataset_usuariosMAN16.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN17' then begin
            if ibdataset_usuariosMAN17.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN18' then begin
            if ibdataset_usuariosMAN18.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN19' then begin
            if ibdataset_usuariosMAN19.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MAN19' then begin
            if ibdataset_usuariosMAN19.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV01' then begin
            if ibdataset_usuariosMOV01.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV02' then begin
            if ibdataset_usuariosMOV02.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV03' then begin
            if ibdataset_usuariosMOV03.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV04' then begin
            if ibdataset_usuariosMOV04.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV05' then begin
            if ibdataset_usuariosMOV05.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV06' then begin
            if ibdataset_usuariosMOV06.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV07' then begin
            if ibdataset_usuariosMOV07.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV08' then begin
            if ibdataset_usuariosMOV08.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV09' then begin
            if ibdataset_usuariosMOV09.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV10' then begin
            if ibdataset_usuariosMOV10.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV11' then begin
            if ibdataset_usuariosMOV11.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV12' then begin
            if ibdataset_usuariosMOV12.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV13' then begin
            if ibdataset_usuariosMOV13.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV14' then begin
            if ibdataset_usuariosMOV14.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV15' then begin
            if ibdataset_usuariosMOV15.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV16' then begin
            if ibdataset_usuariosMOV16.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV17' then begin
            if ibdataset_usuariosMOV17.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV18' then begin
            if ibdataset_usuariosMOV18.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV19' then begin
            if ibdataset_usuariosMOV19.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='MOV20' then begin
            if ibdataset_usuariosMOV20.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL01' then begin
            if ibdataset_usuariosREL01.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL02' then begin
            if ibdataset_usuariosREL02.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL03' then begin
            if ibdataset_usuariosREL03.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL04' then begin
            if ibdataset_usuariosREL04.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL05' then begin
            if ibdataset_usuariosREL05.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL06' then begin
            if ibdataset_usuariosREL06.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL07' then begin
            if ibdataset_usuariosREL07.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL08' then begin
            if ibdataset_usuariosREL08.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL09' then begin
            if ibdataset_usuariosREL09.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL10' then begin
            if ibdataset_usuariosREL10.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL11' then begin
            if ibdataset_usuariosREL11.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL12' then begin
            if ibdataset_usuariosREL12.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL13' then begin
            if ibdataset_usuariosREL13.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL14' then begin
            if ibdataset_usuariosREL14.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL15' then begin
            if ibdataset_usuariosREL15.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL16' then begin
            if ibdataset_usuariosREL16.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL17' then begin
            if ibdataset_usuariosREL17.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL18' then begin
            if ibdataset_usuariosREL18.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL19' then begin
            if ibdataset_usuariosREL19.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL20' then begin
            if ibdataset_usuariosREL20.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL21' then begin
            if ibdataset_usuariosREL21.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL22' then begin
            if ibdataset_usuariosREL22.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL23' then begin
            if ibdataset_usuariosREL23.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL24' then begin
            if ibdataset_usuariosREL24.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL25' then begin
            if ibdataset_usuariosREL25.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL26' then begin
            if ibdataset_usuariosREL26.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL27' then begin
            if ibdataset_usuariosREL27.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL28' then begin
            if ibdataset_usuariosREL28.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL29' then begin
            if ibdataset_usuariosREL29.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL30' then begin
            if ibdataset_usuariosREL30.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL31' then begin
            if ibdataset_usuariosREL31.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL32' then begin
            if ibdataset_usuariosREL32.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL33' then begin
            if ibdataset_usuariosREL33.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL34' then begin
            if ibdataset_usuariosREL34.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL35' then begin
            if ibdataset_usuariosREL35.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL36' then begin
            if ibdataset_usuariosREL36.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL37' then begin
            if ibdataset_usuariosREL37.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL38' then begin
            if ibdataset_usuariosREL38.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL39' then begin
            if ibdataset_usuariosREL39.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='REL40' then begin
            if ibdataset_usuariosREL40.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER01' then begin
            if ibdataset_usuariosFER01.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER02' then begin
            if ibdataset_usuariosFER02.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER03' then begin
            if ibdataset_usuariosFER03.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER04' then begin
            if ibdataset_usuariosFER04.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER05' then begin
            if ibdataset_usuariosFER05.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER06' then begin
            if ibdataset_usuariosFER06.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER07' then begin
            if ibdataset_usuariosFER07.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER08' then begin
            if ibdataset_usuariosFER08.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER09' then begin
            if ibdataset_usuariosFER09.Value='S' then result:=true;
         end;
         if Uppercase(OpcaoMenu)='FER10' then begin
            if ibdataset_usuariosFER10.Value='S' then result:=true;
         end;
       end;
      end;
   end;
  end
  else
     result:=true;
end;

///
///



procedure TPrincipal.ManipulaExcecoes(Sender: TObject; E: Exception);
begin
  MessageDlg(E.Message + #13#13 +
  'Suporte técnico:'#13 +
  'suporte@procedimentodigital.com.br'+#13+
  'www.procedimentodigital.com.br',
  mtError, [mbOK], 0);
end;

procedure TPrincipal.Sair1Click(Sender: TObject);
begin
   Close;
end;

procedure TPrincipal.Departamentos1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFDepartamentos then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFDepartamentos, FDepartamentos);
      end;
end;

procedure TPrincipal.Funcionrios1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFFuncionarios then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFFuncionarios, FFuncionarios);
      end;
    end;

procedure TPrincipal.Funcionrios2Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFFornecedores then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFFornecedores, FFornecedores);
    end;
end;

procedure TPrincipal.Timer1Timer(Sender: TObject);
begin
   if data_sys<>date() then begin
      yWallPapper(Image1);
      data_sys:=date();
      timer1.Interval:=60000;
   end;
   Statusbar1.Panels[3].Text:=TimeToStr(Time());
   Statusbar1.Panels[2].Text:=DateToStr(Date());
end;

procedure TPrincipal.FormActivate(Sender: TObject);
var
Janela: HWND;
arq: TextFile;
controle: string;
vcont: boolean;
vm, va: integer;
vdt: tdatetime;

begin
   Top:=-8;
   if vconfigurado_ok then begin
      vita:='141517';
      principal.WindowState:=wsMaximized;
      vadministrador:='S';
      vmar:='me141517';
      if ok_Sistem=false then begin
         sysutils.ShortDateFormat:='dd/mm/yyyy';
         sysutils.LongDateFormat:='dd/mm/yyyy';
         sysutils.DateSeparator:='/';
         sysutils.ShortTimeFormat:='hh:mm:ss';
         sysutils.TimeSeparator:=':';
         Janela := FindWindow('Shell_TrayWnd', nil);
         if Janela > 0 then
            ShowWindow(Janela, SW_SHOW);
            if vcloud=false then begin
            if fileexists(s0+'liberacao.dat')=false then begin
               if vestacao then begin
                  messagedlG('Servidor não configurado'+#13+
                             'entre em contato com a '+#13+
                             'Procedimento Digital e solicite'+#13+
                             'a configuração do servidor',mtinformation,[mbok],0);
                  close;
               end
               else
               begin
                  avaliacao:=true;
                  fcadastro.ShowModal;
               end;
            end
            else
            begin
               assignfile(arq,s0+'liberacao.dat');
               Reset(arq);
               Readln(arq,vlinha);
               closefile(arq);
               vlinha:=lercriptografia(vlinha);
               vcnpjlib:=right(vlinha,14);
               if (copy(vlinha,1,9)='AQUISICAO') or
                  (copy(vlinha,1,9)='AVALIACAO') then begin
                  fsenha_liberacao.EditSenha.Text:='';
                  fsenha_liberacao.showmodal;
                  if alltrim(fsenha_liberacao.EditSenha.text)<>'' then begin
                     if fsenha_liberacao.EditSenha.Text=vita+vmar then begin
                        grava_chave('DEFINITIVO',vcnpjlib,'dgescola');
                        messagedlg('Sistema Habilitado em Modo Definitivo',mtinformation,[mbok],0);
                        close;
                     end
                     else
                     begin
                        if fsenha_liberacao.EditSenha.text<>checa_chave(copy(datetostr(date()),4,7),vcnpjlib) then
                           avaliacao:=true
                        else
                        begin
                           vmes:=copy(datetostr(date()),4,7);
                           vlinha:=checa_chave(vmes,vcnpjlib);
                           grava_chave(vmes,vcnpjlib,'dgescola');
                        end;
                     end;
                  end
                  else
                     avaliacao:=true;
               end
               else
               begin
                  if vestacao=false then begin
                     gserie:=copy(right(vlinha,22),1,8);
                     if verifica_serie(gserie)=true then begin
                        messagedlg('Número de Série inválido'+#13+
                                   'entre em contato com a '+#13+
                                   'Procedimento Digital',mtinformation,[mbok],0);
                        fsenha_liberacao.EditSenha.Text:='';
                        fsenha_liberacao.showmodal;
                        if alltrim(fsenha_liberacao.EditSenha.text)<>'' then begin
                           if fsenha_liberacao.EditSenha.Text=vita+vmar then begin
                               grava_chave('DEFINITIVO',vcnpjlib,'dgescola');
                              messagedlg('Sistema Habilitado em Modo Definitivo',mtinformation,[mbok],0);
                              close;
                           end
                           else
                           begin
                              if fsenha_liberacao.EditSenha.text<>checa_chave(copy(datetostr(date()),4,7),vcnpjlib) then
                                 avaliacao:=true
                              else
                              begin
                                 vmes:=copy(datetostr(date()),4,7);
                                 vlinha:=checa_chave(vmes,vcnpjlib);
                                 grava_chave(vmes,vcnpjlib,'dgescola');
                              end;
                           end;
                        end
                        else
                           avaliacao:=true;
                        close;
                     end;
                  end
                  else
                  begin
                     if (copy(vlinha,1,8)='COMPLETO') then begin
                        avaliacao:=false;
                     end
                     else
                     begin
                        tlinha:=vlinha;
                     vmes:=copy(tlinha,1,2)+'/'+copy(tlinha,3,4);
                     vlinha:=copy(tlinha,7,15);
                     vcnpjlib:=right(tlinha,14);
                     vm:=strtointdef(copy(vmes,1,2),0)+1;
                     va:=strtointdef(copy(vmes,4,4),0);
                     if vm>12 then begin
                        vm:=1;
                        va:=va+1;
                     end;
                     vdt:=strtodate('10/'+strzero(vm,2,0)+'/'+strzero(va,4,0));
                     if (vmes<>copy(datetostr(date()),4,7)) and
                        ((date())>vdt) then begin
                           vmes:=strzero(vm,2,0)+'/'+strzero(va,4,0);
                           messagedlg('É necessário informar a senha para continuar'+#13+
                                      'a utilização completa do sistema',mtconfirmation,
                                      [mbok],0);
                           fsenha_liberacao.EditSenha.Text:='';
                           fsenha_liberacao.showmodal;
                           vlinha:=fsenha_liberacao.EditSenha.text;
                           if alltrim(vlinha)<>'' then begin
                              if vlinha<>checa_chave(vmes,vcnpjlib) then begin
                                 messagedlg('O sistema entrará em mode de avaliação'+#13+
                                            'Entre em contato com a Procedimento Digital'+#13+
                                            'e veja como obter a senha correta',
                                            mtinformation,[mbok],0);
                                 avaliacao:=true;
                              end
                              else
                                 grava_chave(vmes,vcnpjlib,'dgescola');
                           end;
                        end
                        else
                        begin
                           if vlinha<>checa_chave(vmes,vcnpjlib) then begin
                              fsenha_liberacao.EditSenha.Text:='';
                              fsenha_liberacao.showmodal;
                              vlinha:=fsenha_liberacao.EditSenha.text;
                              if alltrim(vlinha)<>'' then begin
                                 if vlinha<>checa_chave(vmes,vcnpjlib) then begin
                                    messagedlg('O sistema entrará em mode de avaliação'+#13+
                                               'Entre em contato com a Procedimento Digital'+#13+
                                               'e veja como obter a senha correta',
                                               mtinformation,[mbok],0);
                                    avaliacao:=true;
                                 end
                                 else
                                 begin
                                    vmes:=copy(datetostr(date()),4,7);
                                    vlinha:=checa_chave(vmes,vcnpjlib);
                                    grava_chave(vmes,vcnpjlib,'dgescola');
                                  end;
                              end;
                           end;
                        end;
                     end;
                  end;
               end;
            end;
         end;
         vsenha_ok:=true;
         vpede_senha:=false;
         vlogin:='-0';
         with datamodule1 do begin
            ibdataset_usuarios.Close;
            ibdataset_usuarios.SelectSQL.Clear;
            ibdataset_usuarios.SelectSQL.Add('Select * from usuarios');
            ibdataset_usuarios.Open;
            if ibdataset_usuarios.RecordCount>0 then
               vpede_senha:=true;
            ibdataset_empresa.Close;
            ibdataset_empresa.SelectSQL.Clear;
            ibdataset_empresa.SelectSQL.Add('Select * from empresa');
            ibdataset_empresa.Open;
            if ibdataset_empresa.RecordCount=0 then begin
               ibdataset_empresa.Append;
               ibdataset_empresaRAZAO.Value:='*** NOVA EMPRESA ***';
               ibdataset_empresa.Post;
               ibdataset_empresa.ApplyUpdates;
               ibtransaction1.CommitRetaining;
            end;
            ibdataset_empresa.Close;
            ibdataset_empresa.SelectSQL.Clear;
            ibdataset_empresa.SelectSQL.Add('Select * from empresa');
            ibdataset_empresa.Open;
            Empresa:=ibdataset_empresaRAZAO.Value;
            CgcEmp:=ibdataset_empresaCNPJ.Value;
            cidemp:=ibdataset_empresaMUNICIPIO.Value;
            endemp:=ibdataset_empresaENDERECO.Value;
            cepemp:=strzero(ibdataset_empresaCEP.Value,8,0);
            baiemp:=ibdataset_empresaBAIRRO.Value;
            vsimples_nac:=ibdataset_empresaSIMPLES.Value;
            complemp:=ibdataset_empresaCOMPLEMENTO.VALUE;
            numemp:=inttostr(ibdataset_empresaNUMERO.Value);
            vbobina:=ibdataset_parametrosBOBINA.Value;
            foneemp:=strzero(ibdataset_empresaFONE.Value,11,0);
            vinscrmuni:=ibdataset_empresaINSCRMUNI.Value;
            inscr_muni:=ibdataset_empresaINSCRMUNI.Value;
            vfaxemp:=ibdataset_empresaFAX.Text;
            ve_mail:=ibdataset_empresaE_MAIL.Value;
            estadoemp:=ibdataset_empresaESTADO.Value;
            //
            Principal.Caption:='Digital Escola - '+empresa+' / CNPJ: '+cgcemp;
            principal.SetFocus;
{            s1:=s+'dados\dgescola.gdb';
            IBDatabase1.CloseDataSets;
            IBDatabase1.Connected := False;
            IBDatabase1.DatabaseName := s1;
            //
            IbDatabase1.Connected:=True;
            if IBDatabase1.Connected then begin
               IBTransaction1.Active := True;
               for i:= 0 to IBDatabase1.DataSetCount - 1 do
                  IBDatabase1.DataSets[i].Active := True;
            end; }
            ibdataset_parametros.Close;
            ibdataset_parametros.SelectSQL.Clear;
            ibdataset_parametros.SelectSQL.Add('Select * from parametros');
            ibdataset_parametros.Open;
            vcampos_contrato.Lines.Clear;
            vcampos_contrato.Lines.Add(ibdataset_parametrosCAMPOS_CONTRATO.Value);
            vbanco_web:=ibdataset_parametrosBANCO_WEB.Value;
            vsms_login:=ibdataset_parametrosSMS_LOGIN.Value;
            vsms_password:=ibdataset_parametrosSMS_PASSWORD.Value;
            vsms_remetente:=IBDataSet_ParametrosSMS_DE.Value;
            vhost:=IBDataSet_ParametrosSMTP.Value;
            vsenha_email:=IBDataSet_ParametrosEMAIL_SENHA.Value;
            vusuario_email:=IBDataSet_ParametrosEMAIL_USUARIO.Value;
            vbaixa_parcial:=IBDataSet_ParametrosBAIXA_PARCIAL_RECEBER.Value;
            vjurosaomes:=ibdataset_parametrosJUROS.Value;
            vporta_smtp:=strtointdef(IBDataSet_ParametrosPORTA_EMAIL.Text,0);
            vmens_cadastro:=ibdataset_parametrosMENS_CADASTRO.Value;
            vmens_aniversariantes:=ibdataset_parametrosMENS_ANIVERSARIANTES.Value;
            if alltrim(vbaixa_parcial)='' then vbaixa_parcial:='S';
            ibdataset_ano_letivo.Close;
            ibdataset_ano_letivo.SelectSQL.Clear;
            ibdataset_ano_letivo.SelectSQL.Add('Select * from ano_letivo');
            ibdataset_ano_letivo.Open;
            if ibdataset_ano_letivo.RecordCount=0 then begin
               vano_letivo:=strtointdef(copy(datetostr(date()),7,4),0);
               ibdataset_ano_letivo.Append;
               ibdataset_ano_letivoANO.Value:=strtointdef(copy(datetostr(date()),7,4),0);
               ibdataset_ano_letivoFECHADO.Value:='N';
               ibdataset_ano_letivo.Post;
               ibdataset_ano_letivo.ApplyUpdates;
               ibtransaction1.CommitRetaining;
            end
            else
               vano_letivo:=ano_letivo();
            drlabel2.Caption:=inttostr(vano_letivo);
            drlabel2.Repaint;
         end;
      end;
      if vfecha then principal.close else begin
         if ok_sistem=false then begin
            if (vpede_senha) then begin
               fsenha.ShowModal;
            end
            else
               vsenha_ok:=true;
            if vsenha_ok=false then principal.close;
         end;
         ok_sistem:=true;
      end;
   end
   else
   begin
      messagedlg('Sistema necessita de configuração da pasta de execução',
              mtinformation,[mbok],0);
      fconfigurapasta.showmodal;
      if fileexists(fconfigurapasta.EditPasta.Text+'\emp001\dgescola.gdb')=true then begin
         assignfile(arq,'c:\pdigital\dgescola.inf');
         rewrite(arq);
         writeln(arq,fconfigurapastau.vpasta+'\pdigital\dgescola');
         writeln(arq,fconfigurapasta.editpasta.text+'\');
         closefile(arq);
         vconfigurado_ok:=true;
      end;
      close;
   end;
end;

procedure TPrincipal.DRLabel5Click(Sender: TObject);
begin
   browser('http://www.procedimentodigital.com.br');
end;

procedure TPrincipal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   Deletafile(s0+'temp\*.*');
end;

procedure TPrincipal.FormCreate(Sender: TObject);
var
arq: TextFile;
I: Integer;
SR: TSearchRec;
w: string;

begin
   vcampos_contrato:=TMemo.Create(application);
   vcampos_contrato.Parent := Principal;
   vcampos_contrato.Visible:=false;
   vcloud:=false;
   vacesso_ok:=true;
   vestacao:=false;
   if directoryexists('c:\pdigital\firebird')=false then begin
      messagedlg('Será executado a intalação do gerenciado de banco de dados'+#13+
                  'Caso o sistema peça confirmação confirme a execução',
                  mtinformation,[mbok],0);
      WinExec('c:\pdigital\dgescola\installfirebird.bat',SW_SHOWNORMAL);
      assignfile(arq,'c:\pdigital\dgescola\setpath.bat');
      rewrite(arq);
      writeln(arq,'SET FIREBIRD=c:\pdigital');
      closefile(arq);
      WinExec('c:\pdigital\dgescola\setpath.bat',SW_SHOWNORMAL);
   end
   else
      vconfigurado_ok:=false;
   vparametro:='';
   if paramcount>0 then
      vparametro:=paramstr(1);
   xpmenu1.Active:=true;
   if directoryexists('c:\pdigital')=false then
      MkDir('c:\pdigital');
   //
   if fileexists('c:\pdigital\dgescola.inf')=true then
      vconfigurado_ok:=true;
   if vconfigurado_ok then begin
      vita:='141517';
      if fileexists('c:\pdigital\dgescola\inicio.txt') then begin
         deletefile('c:\pdigital\dgescola\inicio.txt');
         deletefile('c:\pdigital\dgescola\firebird.exe');
      end;
      data_sys:=strtodate('01/01/1989');
      vfecha:=false;
      vversao:='1.0';
      label1.Caption:='Versão: '+vversao;
      ok_Sistem:=false;
      vempr_cloud:=0;
      Application.OnException := ManipulaExcecoes;
      I := FindFirst('\pdigital\dgescola\temp\*.*', faAnyFile, SR);
      vmar:='me141517';
      if directoryexists('\pdigital\dgescola\temp')=false then
         MkDir('\pdigital\dgescola\temp');
      while I = 0 do begin
        if (SR.Attr and faDirectory) <> faDirectory then
           DeleteFile('\pdigital\dgescola\temp\'+ SR.Name);
           I := FindNext(SR);
      end;
      if fileexists('C:\pdigital\dgescola.inf') then begin
         AssignFile(arq,'C:\pdigital\dgescola.inf');
         Reset(arq); //abre o arquivo para leitura;
         i:=0;
         While not eof(arq) do begin
            Readln(arq,w); //le do arquivo e desce uma linha. O conteúdo lido é transferido para a variável linha
            inc(i);
            if posicao('ESTACAO',upper(w))>0 then vestacao:=true;
            if i=1 then begin
               S:=W;
               If Not ( S[Length( S )] In ['\', '/'] ) Then
                  S := S + '\'; // Adciona a barra do final caso nao exista...
               if posicao('CLOUD',upper(w))>0 then
                  vcloud:=true;
            end
            else
            begin
               if i=2 then
                  S0:=W
            end;
         End;
         if i=1 then begin
            s0:=s;
         end;
         s2:=s;
         Closefile(arq);
      end
      else
      begin
         s:='\pdigital\dgescola\';
         s0:='\pdigital\dgescola\';
         s2:=s0;
      end;
      if upper(vparametro)='LOCAL' then begin
         s:='c:\pdigital\dgescola\';
         s0:=s;
      end;
      if vcloud then begin
         s:='179.190.53.138:/opt/firebird/data/dgescola/';
         if vempr_cloud=-1 then
            s1:=s+'emp000000/dgescola.GDB'
         else
            s1:=s+'emp'+strzero(vempr_cloud,6,0)+'/DGESCOLA.GDB';
         If Not ( S0[Length( S0 )] In ['\', '/'] ) Then
            S0 := S0 + '\'; // Adciona a barra do final caso nao exista...
      end
      else
      begin
         If Not ( S[Length( S )] In ['\', '/'] ) Then
            S := S + '\'; // Adciona a barra do final caso nao exista...
         If Not ( S0[Length( S0 )] In ['\', '/'] ) Then
            S0 := S0 + '\'; // Adciona a barra do final caso nao exista...
         s1:=s+'emp001\DGESCOLA.GDB';
      end;
      s7:=s+'cidades.gdb';
      s3:=s0+'temp\';
      //

      Application.CreateForm(TDataModule1, DataModule1);
      with DataModule1 do begin
         IBDatabase1.CloseDataSets;
         IBDatabase1.Connected := False;
         IBDatabase1.DatabaseName := s1;
         cidades.closedatasets;
         cidades.connected:=false;
         cidades.databasename :=s7;
         Try
            IbDatabase1.Connected:=True;
            if IBDatabase1.Connected then begin
               IBTransaction1.Active := True;
               for i:= 0 to IBDatabase1.DataSetCount - 1 do
                   IBDatabase1.DataSets[i].Active := True;
            end;
            cidades.Connected:=true;
            if cidades.Connected then begin
               IBTransaction4.Active := True;
               for i:= 0 to cidades.DataSetCount - 1 do
                   cidades.DataSets[i].Active := True;
            end;
            ibdataset_parametros.Close;
            ibdataset_parametros.SelectSQL.Clear;
            ibdataset_parametros.SelectSQL.Add('Select * from parametros');
            ibdataset_parametros.Open;
            vcaracter_decimal:=ibdataset_parametrosCARACTER_DECIMAL.Value;
            if alltrim(vcaracter_decimal)='' then vcaracter_decimal:=',';
         except
            ShowMessage('Erro ao conectar com o banco, verifique se o interbase está rodando.' + #10 +
                        'Caso persista esse erro entre em contato com a PROCEDIMENTO DIGITAL / (75)3624-7727');
            Principal.Close;
         end;
      end;
   end;
end;

procedure TPrincipal.Cursos1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFCursos then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFCursos, FCursos);
    end;
end;

procedure TPrincipal.Unidades1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFUnidades then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFUnidades, FUnidades);
    end;
end;

procedure TPrincipal.urmas1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFTurmas then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFTurmas, FTurmas);
    end;
end;

procedure TPrincipal.Disciplinas1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFDisciplinas then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFDisciplinas, FDisciplinas);
    end;
end;

procedure TPrincipal.Professores1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFProfessores then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFProfessores, FProfessores);
    end;
end;

procedure TPrincipal.LivrosMdulos1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFLivrosModulos then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFLivrosModulos, FLivrosModulos);
    end;
end;

procedure TPrincipal.Alunos1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFAlunos then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFAlunos, FAlunos);
    end;
end;

procedure TPrincipal.Produtos1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFProdutos then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFProdutos, FProdutos);
    end;
end;

procedure TPrincipal.ContasFluxo1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFContasFluxo then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFContasFluxo, FContasFluxo);
    end;
end;

procedure TPrincipal.Departamentos2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TiMPDepartamentos then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TiMPDepartamentos, iMPDepartamentos);
end;

procedure TPrincipal.ServidorSMTPEmail1Click(Sender: TObject);
begin
   fparametros.showmodal;
end;

procedure TPrincipal.SobreoSistema1Click(Sender: TObject);
begin
   ShellAbout(Handle,'Digital Escola - Controle Administrativo para Instituíção Educacional',
             'Procedimento Digital - www.procedimentodigital.com.br',
             Application.Icon.Handle);
end;

procedure TPrincipal.Entradas1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFEntradas then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFEntradas, FEntradas);
    end;
end;

procedure TPrincipal.Enviodeemail1Click(Sender: TObject);
begin
   formemail.showmodal;
end;

procedure TPrincipal.EnviodeSMS1Click(Sender: TObject);
begin
   FSaldoSMS.ShowModal;
end;

procedure TPrincipal.Matrculas1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFMatriculas then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFMatriculas, FMatriculas);
    end;
end;

procedure TPrincipal.ContasaPagaer1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFPagar then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFPagar, FPagar);
    end;
end;

procedure TPrincipal.ContasaReceber1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFReceber then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFReceber, FReceber);
    end;
end;

procedure TPrincipal.FluxodeCaixa1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFFluxo then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFFluxo, FFluxo);
    end;
end;

procedure TPrincipal.Eventos1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFEventos then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFEventos, FEventos);
    end;
end;

procedure TPrincipal.Funcionrios3Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpFuncionarios then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpFuncionarios, ImpFuncionarios);
end;

procedure TPrincipal.Fornecedores1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpFornecedores then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpFornecedores, ImpFornecedores);
end;

procedure TPrincipal.Crusos1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpCursos then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpCursos, ImpCursos);
end;

procedure TPrincipal.Unidades2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpFuncionarios then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpFuncionarios, ImpFuncionarios);
end;

procedure TPrincipal.urmas2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpTurmas then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpTurmas, ImpTurmas);
end;

procedure TPrincipal.Disciplinas2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpDisciplinas then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpDisciplinas, ImpDisciplinas);
end;

procedure TPrincipal.Professores2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpProfessores then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpProfessores, ImpProfessores);
end;

procedure TPrincipal.Alunos2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpAlunos then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpAlunos, ImpAlunos);
end;

procedure TPrincipal.LivrosMdulos2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpLivros then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpLivros, ImpLivros);
end;

procedure TPrincipal.Produtos2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpProdutos then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpProdutos, ImpProdutos);
end;

procedure TPrincipal.ContasFluxo2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpContasFluxo then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpContasFluxo, ImpContasFluxo);
end;

procedure TPrincipal.Mensalidades1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFMensalidades then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFMensalidades, FMensalidades);
    end;
end;

procedure TPrincipal.BaixadeMensalidades1Click(Sender: TObject);
begin
   fbaixamensalidades.showmodal;
end;

procedure TPrincipal.Usurios1Click(Sender: TObject);
begin
 if acesso(vlogin,'adm') then
   FCadastroUsuario.showmodal;
end;

procedure TPrincipal.Entradas2Click(Sender: TObject);
begin
 if acesso(vlogin,'mov02') then begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpEntradas then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TImpEntradas, ImpEntradas);
   end;
 end;
end;

procedure TPrincipal.DadosdaEscola1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFDadosEmpresa then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFDadosEmpresa, FDadosEmpresa);
    end;
end;

procedure TPrincipal.ContasaPagar1Click(Sender: TObject);
begin
 if acesso(vlogin,'mov02') then begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpContasPagar then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TImpContasPagar, ImpContasPagar);
   end;
 end;
end;

procedure TPrincipal.ContasaReceber2Click(Sender: TObject);
begin
 if acesso(vlogin,'mov02') then begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpContasReceber then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TImpContasReceber, ImpContasReceber);
   end;
 end;
end;


procedure TPrincipal.Eventos2Click(Sender: TObject);
begin
 if acesso(vlogin,'mov02') then begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpEventos then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TImpEventos, ImpEventos);
   end;
 end;
end;

procedure TPrincipal.Matrculas2Click(Sender: TObject);
begin
  if acesso(vlogin,'mov02') then begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpMatriculas then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TImpMatriculas, ImpMatriculas);
   end;
 end;
end;

procedure TPrincipal.ParametrosFinanceiroVendas1Click(Sender: TObject);
begin
   FParametrosFinanceira.showmodal;
end;

Procedure TPrincipal.AnoLetivo1Click(Sender: TObject);
begin
   fano_letivo.showmodal;
end;

procedure TPrincipal.ContasBancrias1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TFContasBancarias then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TFContasBancarias, FContasBancarias);
   end;
end;

procedure TPrincipal.LanamentoBancrio1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TFExtrato then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TFExtrato, FExtrato);
   end;
end;

procedure TPrincipal.HistricosPadres1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TFHistoricosPadroes then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TFHistoricosPadroes, FHistoricosPadroes);
   end;
end;

procedure TPrincipal.HistricosExtratoBancrio1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpHistoricos then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TImpHistoricos, ImpHistoricos);
   end;
end;

procedure TPrincipal.ContasBancrias2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpContasBancaria then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TImpContasBancaria, ImpContasBancaria);
   end;
end;

procedure TPrincipal.LanarNotas1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFLancar_notas then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFLancar_notas, FLancar_notas);
    end;
end;

procedure TPrincipal.ProcessamentoDirio1Click(Sender: TObject);
begin
   FProcessamento_Diario.ShowModal;
end;

procedure TPrincipal.Escalas2Click(Sender: TObject);
begin
 if acesso(vlogin,'mov02') then begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpFichadeAlunos then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TImpFichadeAlunos, ImpFichadeAlunos);
   end;
 end;
end;

procedure TPrincipal.MensalidadesemAberto1Click(Sender: TObject);
begin
 if acesso(vlogin,'mov02') then begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpMensalidades then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TImpMensalidades, ImpMensalidades);
   end;
 end;
end;

procedure TPrincipal.Event1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFEventos then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFEventos, FEventos);
    end;
end;

procedure TPrincipal.VendasdeProdutos1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TFVendasProdutos then begin
          Found := i;
       end;
   end;
   if Found >= 0 then begin
      Screen.Forms[Found].Show;
   end
   else
   begin
      Application.CreateForm(TFVendasProdutos, FVendasProdutos);
   end;
end;

procedure TPrincipal.Cont1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpContasBancaria then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpContasBancaria, ImpContasBancaria);
end;

procedure TPrincipal.HistricosPadro1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpHistoricos then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpHistoricos, ImpHistoricos);
end;

procedure TPrincipal.ExtratoBancrio1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpExtrato then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpExtrato, ImpExtrato);
end;

procedure TPrincipal.Mat1Click(Sender: TObject);
begin
   ImpDeclaracaoMatricula.ShowModal;
end;

procedure TPrincipal.Aprovao1Click(Sender: TObject);
begin
   ImpDecl_Aprovacao.ShowModal;
end;

procedure TPrincipal.FrequenciaRegular1Click(Sender: TObject);
begin
   ImpDecl_Frequencia.ShowModal;
end;

procedure TPrincipal.Renda1Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TImpDecl_Renda then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TImpDecl_Renda, ImpDecl_Renda);
end;

procedure TPrincipal.Quitao1Click(Sender: TObject);
begin
   ImpDecl_Quitacao.ShowModal;
end;

procedure TPrincipal.Coordenadores1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TFCoordenadores then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TFCoordenadores, FCoordenadores);
    end;
end;

procedure TPrincipal.Coordenadores2Click(Sender: TObject);
begin
   Found := -1;
   for i := 0 to Screen.FormCount - 1 do begin
       if Screen.Forms[i] is TFImpCoordenadores then begin
          Found := i;
       end;
   end;
   if Found >= 0 then
      Screen.Forms[Found].Show
   else
      Application.CreateForm(TFImpCoordenadores, FImpCoordenadores);
end;

procedure TPrincipal.AlunosAtivos1Click(Sender: TObject);
begin
   FConsultaAlunos_Matriculas.showmodal;
end;

procedure TPrincipal.VendasnoPerodo1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TImpVendas then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TImpVendas, ImpVendas);
    end;
end;

procedure TPrincipal.Responsaveis1Click(Sender: TObject);
begin
    Found := -1;
      for i := 0 to Screen.FormCount - 1 do begin
          if Screen.Forms[i] is TImpResponsaveis then begin
             Found := i;
          end;
      end;
      if Found >= 0 then begin
         Screen.Forms[Found].Show;
      end
      else
      begin
         Application.CreateForm(TImpResponsaveis, ImpResponsaveis);
    end;
end;

procedure TPrincipal.GeraArqpSMS1Click(Sender: TObject);
begin
   farqsms.showmodal;
end;

end.
