// Cadastro_Modelo.cs
// ✅ 100% completo (1 arquivo) baseado no seu código
// ✅ AGORA:
// - Writes no MySQL (grupometalContext) ficam COMENTADOS
// - TODA LEITURA (grid + detalhes + projeto + imagem + tipos) vem do PostGre (gmetalContext)
// - Inserts/Updates/Deletes acontecem APENAS no PostGre
// ✅ EXTRA:
// - Ao salvar NOVO PROJETO ou NOVA REVISÃO -> envia e-mail para Joao.rossitti@grupometal.com.br
// ✅ Regra Produção/Liga nobre (pedido aberto):
// - Novo Projeto: D (sem nobre) / B (com nobre)
// - Nova Revisão: E (sem nobre) / C (com nobre)
// ✅ Regra NOVA (RNC aberta + Revisar FA + liga nobre):
// - Somente para NOVA REVISÃO: se existir RNC aberta no setor "Projeto" com ação "Revisar FA" e a liga da RNC for nobre,
//   então ao SALVAR a revisão o Status deve ser 'C'.
// ✅ NOVO (Imagem Magma obrigatória p/ status B ou C):
// - Se status for B ou C -> AO SALVAR pede imagem Magma (file dialog) e faz upload no Azure images-magma
// ✅ NOVO (Panel2 - Imagem Magma):
// - pnlCardMagma fica visível se a última revisão tem link Imagem_Magma; clique abre popup com Zoom (maximizado)

using Azure;
using Azure.Storage.Blobs;
using Controle_Pedidos;
using Controle_Pedidos.Controle_Producao.Entities; // (MySQL) mantido só para compatibilidade do projeto, writes comentados
using Controle_Pedidos.Entities_GM; // (PostGre) agora é a fonte oficial (leitura + escrita)
using Controle_Pedidos.Sistema_RNCs;
using Microsoft.EntityFrameworkCore;
using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Controle_Pedidos_8.Tela_Cadastro
{
    public partial class Cadastro_Modelo : Form
    {
        private int? _empresaIdSelecionada = null;

        // =========================
        // SQL SERVER GM (DB Server) - tabela rncs
        // =========================
        // ✅ Ajuste para a sua string real (Server/Database/User/Password ou Integrated Security)
        //    Você estava usando Conexao() como se fosse método. O correto é "new Conexao()"
        private readonly string _sqlServerGMConnectionString;

        // =========================
        // AZURE (images-fa)
        // =========================
        private const string _baseUrl = "https://armazenamentoazure.blob.core.windows.net/";
        private const string _containerName = "images-fa";
        private const string _sasToken =
            "sp=racwdl&st=2026-02-10T10:11:30Z&se=2030-01-01T18:26:30Z&spr=https&sv=2024-11-04&sr=c&sig=Z2K2P0XFY5k23tXWJx7W5v9wECugxn2S3Fzne9mAqvg%3D";

        // =========================
        // AZURE MAGMA (images-magma)
        // =========================
        private const string _magmaBaseUrl = "https://armazenamentoazure.blob.core.windows.net/";
        private const string _magmaContainer = "images-magma";
        private const string _magmaSas =
            "sp=racwd&st=2026-02-27T18:27:08Z&se=2030-02-28T02:42:08Z&spr=https&sv=2024-11-04&sr=c&sig=W2YF2W7QWFjFJzPE42PiQ2r5fdrMvnAEzDxm3L8rK7c%3D";

        private CancellationTokenSource? _ctsImagem;

        // ✅ Busca responsiva (debounce)
        private readonly System.Windows.Forms.Timer _timerBuscaModelo;
        private CancellationTokenSource? _ctsBuscaModelo;

        // ✅ Detalhes (cancelamento + proteção durante bind)
        private CancellationTokenSource? _ctsCarregarDetalhes;
        private bool _isBindingGrid = false;

        // =========================
        // CARD: Projetos Status 'R'
        // =========================
        private int _qtdProjetosStatusR = 0;

        // =========================
        // CARD: Imagem Magma (pnlCardMagma)
        // =========================
        private string? _imagemMagmaUrlAtual = null;

        // =========================
        // MODO DA TELA (VIEW/ADD/EDIT)
        // =========================
        private enum FormMode { View, Add, Edit }
        private FormMode _mode = FormMode.View;

        // guarda o ModeloId quando estiver editando
        private int? _modeloIdEmEdicao = null;

        // =========================
        // PROJETO / REVISÃO (estado)
        // =========================
        private bool _edicaoProjetoAtiva = false; // true = criando novo projeto OU nova revisão
        private bool _edicaoEhRevisao = false;

        private int? _modeloIdProjeto = null;
        private int? _projetoIdAtual = null;
        private int _revProjetoAtual = 0;

        // ✅ Status calculado no clique (Novo Projeto / Nova Revisão) de acordo com produção/liga (+ RNC na revisão)
        private char _statusProjetoParaSalvar = 'D';

        // =========================
        // CONTROLE: Projeto reprovado (Status R) que foi aberto pelo CARD
        // =========================
        private int? _projetoIdReprovadoAberto = null;

        // ===== ORIGEM IMAGEM =====
        // Novo Projeto (arquivo)
        private string? _imagemLocalPath = null;
        private string? _imagemExt = null;

        // Nova Revisão (clipboard)
        private byte[]? _imagemBytes = null;
        private string? _imagemNome = null;

        // info do "projeto atual" (última revisão existente do modelo selecionado)
        private bool _temProjetoAtual = false;
        private int _revUltimoProjeto = 0;

        // =========================
        // EMAIL (notificação)
        // =========================
        private const string _emailNotificacaoDestino = "Joao.rossitti@grupometal.com.br";

        public Cadastro_Modelo()
        {
            InitializeComponent();

            // ✅ aqui é onde deve setar o readonly
            _sqlServerGMConnectionString = new Conexao().banco_rnc;

            // ✅ IMPORTANTÍSSIMO: ENTER não dispara buscar cliente (evita bug no RichText)
            this.AcceptButton = null;

            // botões do projeto
            btnNovoProjeto.Click += btnNovoProjeto_Click;
            btnNovaRevisao.Click += btnNovaRevisao_Click;
            btnSalvarProjeto.Click += btnSalvarProjeto_Click;
            btnImprimir.Click += btnImprimir_Click;

            // Card Projetos R (clique no painel/labels)
            pnlCardProjetosR.Click += pnlCardProjetosR_Click;
            lblProjetosRTitle.Click += pnlCardProjetosR_Click;
            lblProjetosRCount.Click += pnlCardProjetosR_Click;

            // ✅ Card Magma (panel2)
            try
            {
                pnlCardMagma.Click += pnlCardMagma_Click;
                label33.Click += pnlCardMagma_Click; // nome informado por você
                pnlCardMagma.Visible = false;
                pnlCardMagma.Enabled = false;
            }
            catch { /* se designer ainda não tiver o controle, não quebra */ }

            // revisão bloqueada por padrão
            txtRevisao.ReadOnly = true;

            // =========================
            // TIMER (DEBOUNCE) DA BUSCA
            // =========================
            _timerBuscaModelo = new System.Windows.Forms.Timer();
            _timerBuscaModelo.Interval = 300;
            _timerBuscaModelo.Tick += async (s, e) =>
            {
                _timerBuscaModelo.Stop();
                await BuscarModelosAsync(txtFiltroModelo.Text);
            };

            txtFiltroModelo.TextChanged += txtFiltroModelo_TextChanged;

            ConfigurarGridModelos();
            dgvModelos.SelectionChanged += dgvModelos_SelectionChanged;

            this.Load += async (s, e) =>
            {
                await CarregarTiposModeloAsync();
                await AtualizarCardProjetosStatusRAsync();
                AtualizarCardMagmaUI(); // inicial
            };

            // =========================
            // BOTÕES (CRUD)
            // =========================
            btnAdicionar.Click += btnAdicionar_Click;
            btnEditar.Click += btnEditar_Click;
            btnRemover.Click += btnRemover_Click;
            btnSalvar.Click += btnSalvar_Click;
            btnCancelar.Click += btnCancelar_Click;

            // começa em view
            SetCamposEditaveis(false);
            AtualizarUIEstado();
        }

        // =========================
        // BUSCAR EMPRESA (POPUP)
        // =========================
        private async void btnBuscarCliente_Click(object sender, EventArgs e)
        {
            using (var f = new FrmSelecionarCliente(txtCliente.Text))
            {
                if (f.ShowDialog(this) == DialogResult.OK)
                {
                    txtCliente.Text = f.EmpresaNomeSelecionado;
                    _empresaIdSelecionada = f.EmpresaIdSelecionado;

                    txtFiltroModelo.Text = "";

                    _mode = FormMode.View;
                    _modeloIdEmEdicao = null;

                    // se estava criando projeto/revisão, cancela
                    CancelarEdicaoProjetoUI();

                    SetCamposEditaveis(false);
                    dgvModelos.Enabled = true;

                    await BuscarModelosAsync("");
                }
            }
        }

        // =========================
        // DTO DO GRID (2 COLUNAS + ID)
        // =========================
        private class ModeloGridRow
        {
            public int ModeloId { get; set; }
            public string Empresa { get; set; } = "";
            public string Numeronocliente { get; set; } = "";
        }

        // =========================
        // FILTRO RESPONSIVO
        // =========================
        private void txtFiltroModelo_TextChanged(object sender, EventArgs e)
        {
            if (_empresaIdSelecionada == null) return;
            if (_mode != FormMode.View) return;
            if (_edicaoProjetoAtiva) return;

            _timerBuscaModelo.Stop();
            _timerBuscaModelo.Start();
        }

        // =========================
        // BUSCAR MODELOS (PostGre)
        // =========================
        private async Task BuscarModelosAsync(string texto)
        {
            if (_empresaIdSelecionada == null)
            {
                dgvModelos.DataSource = null;
                LimparDetalhes();
                AtualizarUIEstado();
                return;
            }

            texto = (texto ?? "").Trim();

            _ctsBuscaModelo?.Cancel();
            _ctsBuscaModelo = new CancellationTokenSource();
            var token = _ctsBuscaModelo.Token;

            try
            {
                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                var query =
                    from m in ctxPg.Modelo
                    join e in ctxPg.Empresa on m.ClienteempresaId equals e.EmpresaId
                    where m.ClienteempresaId == _empresaIdSelecionada.Value
                    select new ModeloGridRow
                    {
                        ModeloId = m.ModeloId,
                        Empresa = e.Sigla ?? "",
                        Numeronocliente = m.Numeronocliente ?? ""
                    };

                if (!string.IsNullOrWhiteSpace(texto))
                    query = query.Where(x => x.Numeronocliente.Contains(texto));

                var lista = await query
                    .AsNoTracking()
                    .OrderBy(x => x.Numeronocliente)
                    .Take(500)
                    .ToListAsync(token);

                if (token.IsCancellationRequested) return;

                _isBindingGrid = true;

                dgvModelos.AutoGenerateColumns = true;
                dgvModelos.DataSource = null;
                dgvModelos.DataSource = lista;

                AjustarGridModelosPosBind();

                if (dgvModelos.Rows.Count > 0)
                {
                    dgvModelos.ClearSelection();
                    dgvModelos.Rows[0].Selected = true;

                    var firstVisibleCol = dgvModelos.Columns
                        .Cast<DataGridViewColumn>()
                        .FirstOrDefault(c => c.Visible);

                    if (firstVisibleCol != null)
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            if (dgvModelos.Rows.Count > 0)
                                dgvModelos.CurrentCell = dgvModelos.Rows[0].Cells[firstVisibleCol.Index];
                        }));
                    }
                }

                _isBindingGrid = false;

                AtualizarUIEstado();

                if (dgvModelos.Rows.Count > 0)
                {
                    var row = dgvModelos.Rows[0].DataBoundItem as ModeloGridRow;
                    if (row != null)
                        await CarregarDetalhesAsync(row.ModeloId);
                }
                else
                {
                    LimparDetalhes();
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao buscar modelos (PostGre): " + ex.Message);
            }
        }

        // =========================
        // SELEÇÃO DO GRID -> DETALHES
        // =========================
        private async void dgvModelos_SelectionChanged(object sender, EventArgs e)
        {
            if (_mode != FormMode.View) return;
            if (_isBindingGrid) return;
            if (_edicaoProjetoAtiva) return;

            if (dgvModelos.CurrentRow?.DataBoundItem is not ModeloGridRow row)
                return;

            AtualizarUIEstado();
            await CarregarDetalhesAsync(row.ModeloId);
        }

        // =========================
        // CARREGAR DETALHES DO MODELO (PostGre)
        // =========================
        private async Task CarregarDetalhesAsync(int modeloId)
        {
            _ctsCarregarDetalhes?.Cancel();
            _ctsCarregarDetalhes = new CancellationTokenSource();
            var token = _ctsCarregarDetalhes.Token;

            try
            {
                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                _temProjetoAtual = false;
                _revUltimoProjeto = 0;

                _imagemMagmaUrlAtual = null;
                AtualizarCardMagmaUI();

                var m = await ctxPg.Modelo
                    .AsNoTracking()
                    .Where(x => x.ModeloId == modeloId)
                    .Select(x => new
                    {
                        x.ModeloId,
                        x.Numeronocliente,
                        x.TipodemodeloId,
                        x.Descricao,
                        x.Situacao,
                        x.Observacao,

                        x.Valordomodelo,
                        x.Nrodepartes,
                        x.Numerodefiguras,
                        x.Numerodecaixasdemachoporfigura,
                        x.Estado,

                        x.Pesomedio,   // PB estimado
                        x.Pesobruto,   // PB real
                        x.Pesoprevisto,// PL estimado
                        x.Pesoliquido  // PL real
                    })
                    .FirstOrDefaultAsync(token);

                if (token.IsCancellationRequested) return;

                if (m == null)
                {
                    LimparDetalhes();
                    return;
                }

                txtNumeroModelo.Text = m.Numeronocliente ?? "";
                cmbTipoModelo.SelectedValue = m.TipodemodeloId;
                txtDescricao.Text = m.Descricao ?? "";
                cmbSituacao.Text = m.Situacao.ToString();

                txtValorModelo.Text = m.Valordomodelo?.ToString() ?? "";
                txtNroPartes.Text = m.Nrodepartes?.ToString() ?? "";
                txtQtdFigurasPlaca.Text = m.Numerodefiguras?.ToString() ?? "";
                txtCxMachoFigura.Text = m.Numerodecaixasdemachoporfigura?.ToString() ?? "";
                cmbEstado.Text = m.Estado?.ToString() ?? "";

                txtObservacao.Text = m.Observacao ?? "";

                txtPesoBrutoEst.Text = m.Pesomedio?.ToString() ?? "";
                txtPesoBrutoReal.Text = m.Pesobruto?.ToString() ?? "";
                txtPesoLiquidoEst.Text = m.Pesoprevisto?.ToString() ?? "";
                txtPesoLiquidoReal.Text = m.Pesoliquido?.ToString() ?? "";

                string nomeCliente = "";
                if (_empresaIdSelecionada != null)
                {
                    nomeCliente = await ctxPg.Empresa
                        .AsNoTracking()
                        .Where(e => e.EmpresaId == _empresaIdSelecionada.Value)
                        .Select(e => e.Sigla ?? e.Nome ?? "")
                        .FirstOrDefaultAsync(token) ?? "";
                }
                if (token.IsCancellationRequested) return;

                var projetos2 = await ctxPg.Projeto
                    .AsNoTracking()
                    .Where(p => p.ModeloId == modeloId)
                    .OrderByDescending(p => p.NroRevisao)
                    .Select(p => new
                    {
                        p.NroRevisao,
                        p.Rendimento,
                        p.Elaborador,
                        p.DataCriacao,
                        p.ColaboradorAprovadorSigla
                    })
                    .Take(2)
                    .ToListAsync(token);

                if (token.IsCancellationRequested) return;

                var p1 = projetos2.Count > 0 ? projetos2[0] : null;
                var p2 = projetos2.Count > 1 ? projetos2[1] : null;

                var ultProj = await ctxPg.Projeto
                    .AsNoTracking()
                    .Where(p => p.ModeloId == modeloId)
                    .OrderByDescending(p => p.NroRevisao)
                    .Select(p => new
                    {
                        p.ProjetoId,
                        p.NroRevisao,
                        p.DadosTecnicos
                    })
                    .FirstOrDefaultAsync(token);

                if (token.IsCancellationRequested) return;

                _temProjetoAtual = ultProj != null;
                _revUltimoProjeto = ultProj?.NroRevisao ?? 0;

                txtClienteHeader.Text = nomeCliente;
                txtDescricaoHeader.Text = m.Descricao ?? "";
                txtModeloHeader.Text = m.Numeronocliente ?? "";

                AplicarHeaderPesosDoModelo(
                    pbEst: m.Pesomedio,
                    pbReal: m.Pesobruto,
                    plEst: m.Pesoprevisto,
                    plReal: m.Pesoliquido
                );

                txtRevHeader.Text = p1?.NroRevisao.ToString() ?? "";
                txtRevHeader2.Text = p2?.NroRevisao.ToString() ?? "";

                txtElabHeader.Text = p1?.Elaborador ?? "";
                txtElabHeader2.Text = p2?.Elaborador ?? "";

                txtDataHeader.Text = p1?.DataCriacao.ToString("dd/MM/yyyy") ?? "";
                txtDataHeader2.Text = p2?.DataCriacao.ToString("dd/MM/yyyy") ?? "";

                txtAprovHeader.Text = p1?.ColaboradorAprovadorSigla ?? "";
                txtAprovHeader2.Text = p2?.ColaboradorAprovadorSigla ?? "";

                if (!_edicaoProjetoAtiva)
                {
                    SetRichTextSafe(txtRevisao, ultProj?.DadosTecnicos ?? "");
                    txtRevisao.ReadOnly = true;
                }

                if (ultProj != null)
                {
                    var lastEntity = await ctxPg.Projeto
                        .AsNoTracking()
                        .FirstOrDefaultAsync(p => p.ProjetoId == ultProj.ProjetoId, token);

                    _imagemMagmaUrlAtual = TryGetProjetoImagemMagma(lastEntity);
                    AtualizarCardMagmaUI();
                }
                else
                {
                    _imagemMagmaUrlAtual = null;
                    AtualizarCardMagmaUI();
                }

                await CarregarImagemDoModeloViaProjetoImagemAsync(modeloId, token);
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar detalhes (PostGre): " + ex.Message);
            }
            finally
            {
                AtualizarUIEstado();
            }
        }

        // =========================
        // CARD MAGMA UI
        // =========================
        private void AtualizarCardMagmaUI()
        {
            try
            {
                bool tem = !string.IsNullOrWhiteSpace(_imagemMagmaUrlAtual);

                pnlCardMagma.Visible = tem;
                pnlCardMagma.Enabled = tem;
                pnlCardMagma.Cursor = tem ? Cursors.Hand : Cursors.Default;

                label33.Cursor = tem ? Cursors.Hand : Cursors.Default;
            }
            catch { }
        }

        private async void pnlCardMagma_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_edicaoProjetoAtiva) return;

                if (string.IsNullOrWhiteSpace(_imagemMagmaUrlAtual))
                    return;

                using var frm = new FrmImagemMagma(
                    url: _imagemMagmaUrlAtual,
                    magmaSas: _magmaSas,
                    parent: this
                );

                frm.ShowDialog(this);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao abrir imagem Magma:\n\n" + ex.GetBaseException().Message);
            }
        }

        // =========================
        // IMAGEM: busca bytes em ProjetoImagem (PostGre)
        // =========================
        private async Task CarregarImagemDoModeloViaProjetoImagemAsync(int modeloId, CancellationToken token)
        {
            _ctsImagem?.Cancel();
            _ctsImagem = new CancellationTokenSource();
            var imgToken = _ctsImagem.Token;

            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(token, imgToken);
            var ct = linkedCts.Token;

            try
            {
                SetImagemLoading();

                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                var imgRow = await ctxPg.ProjetoImagem
                    .AsNoTracking()
                    .Where(x => x.ModeloId == modeloId)
                    .OrderByDescending(x => x.ImagemPadrao)
                    .ThenByDescending(x => x.NroRevisao)
                    .Select(x => new
                    {
                        x.ImagemNome,
                        x.Imagem
                    })
                    .FirstOrDefaultAsync(ct);

                if (ct.IsCancellationRequested) return;

                if (imgRow == null || imgRow.Imagem == null || imgRow.Imagem.Length == 0)
                {
                    LimparImagem();
                    return;
                }

                using var ms = new MemoryStream(imgRow.Imagem);
                using var temp = Image.FromStream(ms);
                var bmp = new Bitmap(temp);

                if (picModelo.InvokeRequired)
                    picModelo.BeginInvoke(new Action(() => TrocarImagem(bmp)));
                else
                    TrocarImagem(bmp);
            }
            catch (OperationCanceledException) { }
            catch
            {
                LimparImagem();
            }
        }

        private void TrocarImagem(Image nova)
        {
            if (picModelo.Image != null)
            {
                var old = picModelo.Image;
                picModelo.Image = null;
                old.Dispose();
            }

            picModelo.Image = nova;
        }

        private void LimparDetalhes()
        {
            txtNumeroModelo.Text = "";
            cmbTipoModelo.Text = "";
            txtDescricao.Text = "";
            cmbSituacao.Text = "";

            txtValorModelo.Text = "";
            txtNroPartes.Text = "";
            txtQtdFigurasPlaca.Text = "";
            txtCxMachoFigura.Text = "";
            cmbEstado.Text = "";

            txtObservacao.Text = "";

            txtPesoBrutoEst.Text = "";
            txtPesoBrutoReal.Text = "";
            txtPesoLiquidoEst.Text = "";
            txtPesoLiquidoReal.Text = "";

            txtClienteHeader.Text = "";
            txtPBHeader.Text = "";
            txtPLHeader.Text = "";
            txtDescricaoHeader.Text = "";
            txtModeloHeader.Text = "";
            txtRendimentoHeader.Text = "";
            txtRevHeader.Text = "";
            txtRevHeader2.Text = "";
            txtElabHeader.Text = "";
            txtElabHeader2.Text = "";
            txtDataHeader.Text = "";
            txtDataHeader2.Text = "";
            txtAprovHeader.Text = "";
            txtAprovHeader2.Text = "";

            LimparImagem();
            SetRichTextSafe(txtRevisao, "");
            txtRevisao.ReadOnly = true;

            _temProjetoAtual = false;
            _revUltimoProjeto = 0;

            _imagemMagmaUrlAtual = null;
            AtualizarCardMagmaUI();
        }

        private void LimparImagem()
        {
            if (picModelo.InvokeRequired)
            {
                picModelo.BeginInvoke(new Action(LimparImagem));
                return;
            }

            if (picModelo.Image != null)
            {
                var old = picModelo.Image;
                picModelo.Image = null;
                old.Dispose();
            }
        }

        private void SetImagemLoading()
        {
            LimparImagem();
        }

        // =========================
        // CONFIG GRID
        // =========================
        private void ConfigurarGridModelos()
        {
            dgvModelos.RowHeadersVisible = false;

            dgvModelos.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvModelos.MultiSelect = false;
            dgvModelos.ReadOnly = true;
            dgvModelos.AllowUserToAddRows = false;
            dgvModelos.AllowUserToDeleteRows = false;
            dgvModelos.AllowUserToResizeRows = false;

            dgvModelos.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dgvModelos.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
        }

        private void AjustarGridModelosPosBind()
        {
            ConfigurarGridModelos();
            dgvModelos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            if (dgvModelos.Columns.Contains("Empresa"))
            {
                dgvModelos.Columns["Empresa"].HeaderText = "Empresa";
                dgvModelos.Columns["Empresa"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgvModelos.Columns["Empresa"].FillWeight = 30;
            }

            if (dgvModelos.Columns.Contains("Numeronocliente"))
            {
                dgvModelos.Columns["Numeronocliente"].HeaderText = "Número do Modelo";
                dgvModelos.Columns["Numeronocliente"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgvModelos.Columns["Numeronocliente"].FillWeight = 70;
            }

            if (dgvModelos.Columns.Contains("ModeloId"))
                dgvModelos.Columns["ModeloId"].Visible = false;
        }

        // =========================
        // ADD / EDIT / SAVE / CANCEL (MODELO)
        // =========================
        private void btnAdicionar_Click(object sender, EventArgs e) => EntrarModoAdicionar();

        private void EntrarModoAdicionar()
        {
            if (_empresaIdSelecionada == null)
            {
                MessageBox.Show("Selecione um cliente antes de adicionar um modelo.");
                return;
            }

            _mode = FormMode.Add;
            _modeloIdEmEdicao = null;

            CancelarEdicaoProjetoUI();

            LimparDetalhes();
            SetCamposEditaveis(true);

            dgvModelos.Enabled = false;
            dgvModelos.ClearSelection();

            AtualizarUIEstado();
            txtNumeroModelo.Focus();
        }

        private void btnEditar_Click(object sender, EventArgs e) => EntrarModoEditar();

        private void EntrarModoEditar()
        {
            if (_empresaIdSelecionada == null)
            {
                MessageBox.Show("Selecione um cliente antes de editar.");
                return;
            }

            if (dgvModelos.CurrentRow?.DataBoundItem is not ModeloGridRow row)
            {
                MessageBox.Show("Selecione um modelo no grid para editar.");
                return;
            }

            _mode = FormMode.Edit;
            _modeloIdEmEdicao = row.ModeloId;

            CancelarEdicaoProjetoUI();

            SetCamposEditaveis(true);
            dgvModelos.Enabled = false;

            AtualizarUIEstado();
            txtNumeroModelo.Focus();
        }

        private async void btnCancelar_Click(object sender, EventArgs e)
        {
            await CancelarEdicaoAsync();
        }

        private async Task CancelarEdicaoAsync()
        {
            _mode = FormMode.View;
            _modeloIdEmEdicao = null;

            SetCamposEditaveis(false);
            dgvModelos.Enabled = true;

            AtualizarUIEstado();

            if (dgvModelos.CurrentRow?.DataBoundItem is ModeloGridRow row)
                await CarregarDetalhesAsync(row.ModeloId);
            else
                LimparDetalhes();
        }

        private async void btnSalvar_Click(object sender, EventArgs e)
        {
            await SalvarModeloAsync();
        }

        // =========================
        // SALVAR MODELO (PostGre)
        // =========================
        private async Task SalvarModeloAsync()
        {
            if (_empresaIdSelecionada == null)
            {
                MessageBox.Show("Selecione um cliente.");
                return;
            }

            if (_mode != FormMode.Add && _mode != FormMode.Edit)
                return;

            var numero = (txtNumeroModelo.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(numero))
            {
                MessageBox.Show("Informe o Número do Modelo.");
                txtNumeroModelo.Focus();
                return;
            }

            try
            {
                // =========================================================
                // 🚫 MySQL write (COMENTADO)
                // =========================================================
                // using var ctx = new grupometalContext();
                // Modelo_EF entity;

                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();
                global::Controle_Pedidos.Entities_GM.Modelo_EF entityPg;

                if (_mode == FormMode.Add)
                {
                    entityPg = new global::Controle_Pedidos.Entities_GM.Modelo_EF();
                    entityPg.ClienteempresaId = _empresaIdSelecionada.Value;
                    ctxPg.Modelo.Add(entityPg);
                }
                else
                {
                    if (_modeloIdEmEdicao == null)
                    {
                        MessageBox.Show("Não foi possível identificar o modelo em edição.");
                        return;
                    }

                    entityPg = await ctxPg.Modelo
                        .FirstOrDefaultAsync(x => x.ModeloId == _modeloIdEmEdicao.Value);

                    if (entityPg == null)
                    {
                        MessageBox.Show("Modelo não encontrado (PostGre).");
                        await CancelarEdicaoAsync();
                        return;
                    }
                }

                entityPg.Numeronocliente = (txtNumeroModelo.Text ?? "").Trim();
                entityPg.Descricao = (txtDescricao.Text ?? "").Trim();
                entityPg.Observacao = (txtObservacao.Text ?? "").Trim();

                entityPg.Situacao = ParseCharOrDefault(cmbSituacao.Text, 'A');
                entityPg.Estado = ParseNullableChar(cmbEstado.Text);

                if (cmbTipoModelo.SelectedValue is int tipoId)
                    entityPg.TipodemodeloId = tipoId;
                else
                    entityPg.TipodemodeloId = 0;

                entityPg.Nrodepartes = ParseIntOrZero(txtNroPartes.Text);
                entityPg.Numerodefiguras = (txtQtdFigurasPlaca.Text ?? "").Trim();
                entityPg.Numerodecaixasdemachoporfigura = (txtCxMachoFigura.Text ?? "").Trim();

                entityPg.Valordomodelo = ParseFloatOrNull(txtValorModelo.Text);

                entityPg.Pesomedio = ParseFloatOrNull(txtPesoBrutoEst.Text);
                entityPg.Pesobruto = ParseFloatOrNull(txtPesoBrutoReal.Text);
                entityPg.Pesoprevisto = ParseFloatOrNull(txtPesoLiquidoEst.Text);
                entityPg.Pesoliquido = ParseFloatOrNull(txtPesoLiquidoReal.Text);

                await ctxPg.SaveChangesAsync();

                int savedId = entityPg.ModeloId;

                _mode = FormMode.View;
                _modeloIdEmEdicao = null;

                SetCamposEditaveis(false);
                dgvModelos.Enabled = true;
                AtualizarUIEstado();

                await BuscarModelosAsync(txtFiltroModelo.Text);
                SelecionarModeloNoGrid(savedId);
                await CarregarDetalhesAsync(savedId);

                MessageBox.Show("Modelo salvo com sucesso (PostGre)!");
            }
            catch (DbUpdateException dbEx)
            {
                var root = dbEx.GetBaseException();
                MessageBox.Show(
                    "Erro ao salvar modelo (PostGre):\n\n" +
                    root.Message +
                    "\n\nDetalhes:\n" +
                    (dbEx.InnerException?.Message ?? "(sem inner)"),
                    "Erro PostGre",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao salvar modelo (PostGre): " + ex.GetBaseException().Message);
            }
        }

        private async void btnRemover_Click(object sender, EventArgs e)
        {
            await RemoverModeloSelecionadoAsync();
        }

        private async Task RemoverModeloSelecionadoAsync()
        {
            if (_mode != FormMode.View)
            {
                MessageBox.Show("Finalize ou cancele a edição antes de remover.");
                return;
            }

            if (_empresaIdSelecionada == null)
            {
                MessageBox.Show("Selecione um cliente.");
                return;
            }

            if (dgvModelos.CurrentRow?.DataBoundItem is not ModeloGridRow row)
            {
                MessageBox.Show("Selecione um modelo no grid para remover.");
                return;
            }

            var modeloId = row.ModeloId;
            var numeroModelo = row.Numeronocliente;

            var resp = MessageBox.Show(
                $"Tem certeza que deseja REMOVER o modelo:\n\n{numeroModelo}\n(ModeloId: {modeloId})\n\nEssa ação não pode ser desfeita.",
                "Confirmar remoção",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (resp != DialogResult.Yes)
                return;

            try
            {
                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                bool temProjeto = await ctxPg.Projeto
                    .AsNoTracking()
                    .AnyAsync(p => p.ModeloId == modeloId);

                bool temImagem = await ctxPg.ProjetoImagem
                    .AsNoTracking()
                    .AnyAsync(i => i.ModeloId == modeloId);

                if (temProjeto || temImagem)
                {
                    MessageBox.Show(
                        "Não é possível remover este modelo porque existem registros relacionados.\n\n" +
                        $"Projetos: {(temProjeto ? "SIM" : "NÃO")}\n" +
                        $"Imagens: {(temImagem ? "SIM" : "NÃO")}\n\n" +
                        "Remova/ajuste os registros relacionados primeiro.",
                        "Remoção bloqueada",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                var entity = await ctxPg.Modelo
                    .FirstOrDefaultAsync(m => m.ModeloId == modeloId);

                if (entity == null)
                {
                    MessageBox.Show("Modelo não encontrado (PostGre).");
                    await BuscarModelosAsync(txtFiltroModelo.Text);
                    return;
                }

                ctxPg.Modelo.Remove(entity);
                await ctxPg.SaveChangesAsync();

                await BuscarModelosAsync(txtFiltroModelo.Text);

                if (dgvModelos.Rows.Count == 0) LimparDetalhes();

                MessageBox.Show("Modelo removido com sucesso (PostGre)!");
            }
            catch (DbUpdateException dbEx)
            {
                MessageBox.Show(
                    "Não foi possível remover o modelo (provável vínculo com outras tabelas).\n\n" +
                    dbEx.GetBaseException().Message,
                    "Erro ao remover",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao remover (PostGre): " + ex.Message);
            }
        }

        // =========================
        // HEADER: ReadOnly/Enabled
        // =========================
        private void SetCamposEditaveis(bool editavelModelo)
        {
            txtNumeroModelo.ReadOnly = !editavelModelo;
            txtDescricao.ReadOnly = !editavelModelo;
            txtValorModelo.ReadOnly = !editavelModelo;
            txtNroPartes.ReadOnly = !editavelModelo;
            txtQtdFigurasPlaca.ReadOnly = !editavelModelo;
            txtCxMachoFigura.ReadOnly = !editavelModelo;
            txtObservacao.ReadOnly = !editavelModelo;

            txtPesoBrutoEst.ReadOnly = !editavelModelo;
            txtPesoBrutoReal.ReadOnly = !editavelModelo;
            txtPesoLiquidoEst.ReadOnly = !editavelModelo;
            txtPesoLiquidoReal.ReadOnly = !editavelModelo;

            cmbTipoModelo.Enabled = editavelModelo;
            cmbSituacao.Enabled = editavelModelo;
            cmbEstado.Enabled = editavelModelo;

            bool editProjeto = _edicaoProjetoAtiva;

            txtClienteHeader.ReadOnly = true;
            txtDescricaoHeader.ReadOnly = true;
            txtModeloHeader.ReadOnly = true;

            txtRevHeader.ReadOnly = true;
            txtDataHeader.ReadOnly = true;
            txtAprovHeader.ReadOnly = true;

            txtPBHeader.ReadOnly = !editProjeto;
            txtPLHeader.ReadOnly = !editProjeto;
            txtRendimentoHeader.ReadOnly = !editProjeto;

            txtRevHeader2.ReadOnly = true;
            txtElabHeader.ReadOnly = true;
            txtElabHeader2.ReadOnly = true;
            txtDataHeader2.ReadOnly = true;
            txtAprovHeader2.ReadOnly = true;
        }

        private void AtualizarUIEstado()
        {
            bool isView = _mode == FormMode.View;
            bool isEditOrAdd = _mode == FormMode.Add || _mode == FormMode.Edit;

            btnAdicionar.Enabled = isView && !_edicaoProjetoAtiva;
            btnEditar.Enabled = isView && dgvModelos.CurrentRow != null && !_edicaoProjetoAtiva;
            btnRemover.Enabled = isView && dgvModelos.CurrentRow != null && !_edicaoProjetoAtiva;

            btnSalvar.Enabled = isEditOrAdd;
            btnCancelar.Enabled = isEditOrAdd;
            btnSalvar.Visible = isEditOrAdd;
            btnCancelar.Visible = isEditOrAdd;

            bool temCliente = _empresaIdSelecionada != null;
            bool temModeloSelecionado = dgvModelos.CurrentRow != null;

            btnNovoProjeto.Visible = temCliente;
            btnNovaRevisao.Visible = temCliente;

            btnNovoProjeto.Enabled = temCliente && temModeloSelecionado && isView && !_edicaoProjetoAtiva && !_temProjetoAtual;
            btnNovaRevisao.Enabled = temCliente && temModeloSelecionado && isView && !_edicaoProjetoAtiva && _temProjetoAtual;

            btnSalvarProjeto.Enabled = _edicaoProjetoAtiva;

            dgvModelos.Enabled = !_edicaoProjetoAtiva && _mode == FormMode.View;
        }

        private void SelecionarModeloNoGrid(int modeloId)
        {
            if (dgvModelos.Rows.Count == 0) return;

            foreach (DataGridViewRow r in dgvModelos.Rows)
            {
                if (r.DataBoundItem is ModeloGridRow row && row.ModeloId == modeloId)
                {
                    dgvModelos.ClearSelection();
                    r.Selected = true;

                    var firstVisibleCol = dgvModelos.Columns
                        .Cast<DataGridViewColumn>()
                        .FirstOrDefault(c => c.Visible);

                    if (firstVisibleCol != null)
                        dgvModelos.CurrentCell = r.Cells[firstVisibleCol.Index];

                    return;
                }
            }
        }

        // =========================================================
        // ✅ REGRA Produção/Liga nobre (pedido aberto) => status
        // =========================================================
        private async Task<char> CalcularStatusPorPedidoELigaAsync(int modeloId, bool ehRevisao)
        {
            char statusSemNobre = ehRevisao ? 'E' : 'D';
            char statusComNobre = ehRevisao ? 'C' : 'B';

            try
            {
                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                bool temLigaNobre = await (
                    from pi in ctxPg.ProducaoItem.AsNoTracking()
                    join l in ctxPg.Liga.AsNoTracking() on pi.LigaId equals l.LigaId
                    where pi.ModeloId == modeloId
                          && (pi.EtapaId == 50 || pi.EtapaId == 75 || pi.EtapaId == 100 || pi.EtapaId == 150 || pi.EtapaId == 200)
                          && l.Necessita_Aprovacao_FA == true
                    select 1
                ).AnyAsync();

                return temLigaNobre ? statusComNobre : statusSemNobre;
            }
            catch
            {
                return statusSemNobre;
            }
        }

        // =========================================================
        // ✅ NOVO: REGRA RNC aberta (Setor Projeto + Ação "Revisar FA")
        // =========================================================

        private static string NormalizeAlphaNumUpper(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim();
            var chars = s.Where(ch => char.IsLetterOrDigit(ch)).ToArray();
            return new string(chars).ToUpperInvariant();
        }

        private static bool IsAcaoRevisarFA(string? descricaoAcao)
        {
            var n = NormalizeAlphaNumUpper(descricaoAcao);
            return n.Contains("REVISARFA");
        }

        private static bool IsSetorProjeto(string? setor)
        {
            var n = NormalizeAlphaNumUpper(setor);
            return n.Contains("PROJETO");
        }

        private List<string> ListarNumerosRncAbertasSetorProjeto_RevisarFA()
        {
            try
            {
                var dados_acao = new Operacao_RNCs();
                var acoes_aberto = dados_acao.Acoes_Em_Aberto();

                if (acoes_aberto == null || acoes_aberto.Count == 0)
                    return new List<string>();

                var rncs = acoes_aberto
                    .Where(a =>
                        IsSetorProjeto(a.descricao_setor) &&
                        IsAcaoRevisarFA(a.descricao_acao) &&
                        !string.IsNullOrWhiteSpace(a.numero_rnc)
                    )
                    .Select(a => a.numero_rnc.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                return rncs;
            }
            catch
            {
                return new List<string>();
            }
        }

        private async Task<List<(string rnc, int modeloId, int ligaId)>> BuscarDadosRncsNoSqlServerAsync(List<string> numerosRnc)
        {
            var result = new List<(string rnc, int modeloId, int ligaId)>();

            if (numerosRnc == null || numerosRnc.Count == 0)
                return result;

            var lista = numerosRnc.Take(500).ToList();

            using var conn = new SqlConnection(_sqlServerGMConnectionString);
            await conn.OpenAsync();

            using var cmd = conn.CreateCommand();

            var paramNames = new List<string>();
            for (int i = 0; i < lista.Count; i++)
            {
                string p = "@r" + i;
                paramNames.Add(p);
                cmd.Parameters.AddWithValue(p, lista[i]);
            }

            // ✅ MUITO IMPORTANTE:
            // Você estava selecionando "numero_rnc, id_modelo, id_liga"
            // mas lendo "RNC" / "modelo_id" / "liga_id".
            // Aqui eu alinho tudo com ALIAS.
            cmd.CommandText = $@"
SELECT
    numero_rnc AS RNC,
    id_modelo  AS modelo_id,
    id_liga    AS liga_id
FROM rncs
WHERE numero_rnc IN ({string.Join(",", paramNames)})
";

            using var rdr = await cmd.ExecuteReaderAsync();
            while (await rdr.ReadAsync())
            {
                string rnc = (rdr["RNC"]?.ToString() ?? "").Trim();

                int modeloId = 0;
                int ligaId = 0;

                _ = int.TryParse(rdr["modelo_id"]?.ToString(), out modeloId);
                _ = int.TryParse(rdr["liga_id"]?.ToString(), out ligaId);

                if (!string.IsNullOrWhiteSpace(rnc) && modeloId > 0 && ligaId > 0)
                    result.Add((rnc, modeloId, ligaId));
            }

            return result;
        }

        private async Task<bool> ExisteRncAbertaLigaNobreParaModeloAsync(int modeloId)
        {
            try
            {
                var numeros = ListarNumerosRncAbertasSetorProjeto_RevisarFA();
                if (numeros.Count == 0) return false;

                var dadosRncs = await BuscarDadosRncsNoSqlServerAsync(numeros);

                var ligaIds = dadosRncs
                    .Where(x => x.modeloId == modeloId)
                    .Select(x => x.ligaId)
                    .Distinct()
                    .ToList();

                if (ligaIds.Count == 0) return false;

                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                bool temLigaNobre = await ctxPg.Liga
                    .AsNoTracking()
                    .AnyAsync(l => ligaIds.Contains(l.LigaId) && l.Necessita_Aprovacao_FA == true);

                return temLigaNobre;
            }
            catch
            {
                return false;
            }
        }

        // ✅ Revisão: primeiro regra de pedido; se não for 'C', checa RNC e pode virar 'C'
        private async Task<char> CalcularStatusRevisaoPorPedidoOuRncAsync(int modeloId)
        {
            var statusPedido = await CalcularStatusPorPedidoELigaAsync(modeloId, ehRevisao: true);
            if (statusPedido == 'C')
                return 'C';

            bool temRncNobre = await ExisteRncAbertaLigaNobreParaModeloAsync(modeloId);
            return temRncNobre ? 'C' : 'E';
        }

        // =========================
        // NOVO PROJETO / NOVA REVISÃO
        // =========================
        private async void btnNovoProjeto_Click(object sender, EventArgs e)
        {
            if (_empresaIdSelecionada == null) return;
            if (_mode != FormMode.View) return;

            if (dgvModelos.CurrentRow?.DataBoundItem is not ModeloGridRow row)
            {
                MessageBox.Show("Selecione um modelo.");
                return;
            }

            if (_temProjetoAtual)
            {
                MessageBox.Show("Este modelo já possui projeto. Use 'Nova Revisão'.");
                return;
            }

            _statusProjetoParaSalvar = await CalcularStatusPorPedidoELigaAsync(row.ModeloId, ehRevisao: false);
            var header = await GetHeaderPesosDoModeloAsync(row.ModeloId);

            _edicaoEhRevisao = false;

            await IniciarEdicaoProjetoAsync(row, rev: 1, rtfBase: "", pb: header.pbText, pl: header.plText, rend: header.rendText);
        }

        private async void btnNovaRevisao_Click(object sender, EventArgs e)
        {
            if (_empresaIdSelecionada == null) return;
            if (_mode != FormMode.View) return;

            if (dgvModelos.CurrentRow?.DataBoundItem is not ModeloGridRow row)
            {
                MessageBox.Show("Selecione um modelo.");
                return;
            }

            var info = await GetUltimoProjetoInfoAsync(row.ModeloId);
            if (!info.ok)
            {
                MessageBox.Show("Este modelo ainda não tem projeto. Use 'Novo Projeto' primeiro.");
                return;
            }

            int novaRev = info.lastRev + 1;

            var (okClip, clipImg) = TryGetClipboardImage();
            if (!okClip || clipImg == null)
            {
                MessageBox.Show(
                    "Nenhuma imagem encontrada no Clipboard.\n\n" +
                    "Copie uma imagem (Ctrl+C) e clique em 'Nova Revisão' novamente.",
                    "Imagem não encontrada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            // ✅ AQUI ENTRA A NOVA REGRA:
            // Revisão agora calcula status por (Pedido OU RNC)
            _statusProjetoParaSalvar = await CalcularStatusRevisaoPorPedidoOuRncAsync(row.ModeloId);

            _imagemBytes = ImageToPngBytes(clipImg);
            _imagemExt = ".png";
            _imagemNome = $"clipboard_rev_{novaRev}.png";

            _imagemLocalPath = null;

            TrocarImagem(new Bitmap(clipImg));

            var header = await GetHeaderPesosDoModeloAsync(row.ModeloId);

            _edicaoEhRevisao = true;

            await IniciarEdicaoProjetoAsync_SemFileDialog(
                row,
                rev: novaRev,
                rtfBase: info.lastRtf,
                pb: header.pbText,
                pl: header.plText,
                rend: header.rendText
            );
        }

        private async Task IniciarEdicaoProjetoAsync(
            ModeloGridRow row,
            int rev,
            string rtfBase,
            string pb,
            string pl,
            string rend)
        {
            _modeloIdProjeto = row.ModeloId;
            _revProjetoAtual = rev;
            _projetoIdAtual = null;

            _imagemBytes = null;
            _imagemNome = null;

            using var ofd = new OpenFileDialog();
            ofd.Title = _edicaoEhRevisao ? "Selecione a imagem da NOVA REVISÃO" : "Selecione a imagem do NOVO PROJETO";
            ofd.Filter = "Imagens|*.jpg;*.jpeg;*.png;*.bmp;*.webp";

            if (ofd.ShowDialog(this) != DialogResult.OK)
                return;

            _imagemLocalPath = ofd.FileName;
            _imagemExt = Path.GetExtension(_imagemLocalPath);

            try
            {
                using var fs = new FileStream(_imagemLocalPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var temp = Image.FromStream(fs);
                TrocarImagem(new Bitmap(temp));
            }
            catch
            {
                MessageBox.Show("Não foi possível carregar a imagem selecionada.");
                _imagemLocalPath = null;
                return;
            }

            _edicaoProjetoAtiva = true;

            SetCamposEditaveis(false);

            txtModeloHeader.Text = row.Numeronocliente ?? "";
            txtDescricaoHeader.Text = (txtDescricao.Text ?? "").Trim();

            txtRevHeader.Text = _revProjetoAtual.ToString();
            txtDataHeader.Text = DateTime.Today.ToString("dd/MM/yyyy");
            txtAprovHeader.Text = "SMR";

            txtPBHeader.Text = pb ?? "";
            txtPLHeader.Text = pl ?? "";
            txtRendimentoHeader.Text = rend ?? "";

            SetRichTextSafe(txtRevisao, rtfBase ?? "");
            txtRevisao.ReadOnly = false;
            txtRevisao.Focus();

            AtualizarUIEstado();
            await Task.CompletedTask;
        }

        private async Task IniciarEdicaoProjetoAsync_SemFileDialog(
            ModeloGridRow row,
            int rev,
            string rtfBase,
            string pb,
            string pl,
            string rend)
        {
            _modeloIdProjeto = row.ModeloId;
            _revProjetoAtual = rev;
            _projetoIdAtual = null;

            if (_imagemBytes == null || _imagemBytes.Length == 0)
            {
                MessageBox.Show("Nenhuma imagem carregada do Clipboard.");
                return;
            }

            _edicaoProjetoAtiva = true;

            SetCamposEditaveis(false);

            txtModeloHeader.Text = row.Numeronocliente ?? "";
            txtDescricaoHeader.Text = (txtDescricao.Text ?? "").Trim();

            txtRevHeader.Text = _revProjetoAtual.ToString();
            txtDataHeader.Text = DateTime.Today.ToString("dd/MM/yyyy");
            txtAprovHeader.Text = "SMR";

            txtPBHeader.Text = pb ?? "";
            txtPLHeader.Text = pl ?? "";
            txtRendimentoHeader.Text = rend ?? "";

            SetRichTextSafe(txtRevisao, rtfBase ?? "");
            txtRevisao.ReadOnly = false;
            txtRevisao.Focus();

            AtualizarUIEstado();
            await Task.CompletedTask;
        }

        private async void btnSalvarProjeto_Click(object sender, EventArgs e)
        {
            await SalvarProjetoOuRevisaoAsync();
        }

        // =========================
        // SALVAR PROJETO/REVISÃO (PostGre) + EMAIL + MAGMA (status B/C)
        // =========================
        private async Task SalvarProjetoOuRevisaoAsync()
        {
            if (!_edicaoProjetoAtiva) return;
            if (_modeloIdProjeto == null) return;

            if (_empresaIdSelecionada == null)
            {
                MessageBox.Show("Selecione um cliente.");
                return;
            }

            bool temArquivo = !string.IsNullOrWhiteSpace(_imagemLocalPath);
            bool temClipboard = _imagemBytes != null && _imagemBytes.Length > 0;

            if (!temArquivo && !temClipboard)
            {
                MessageBox.Show("Selecione uma imagem (arquivo) ou cole uma imagem no Clipboard.");
                return;
            }

            try
            {
                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();
                await using var tx = await ctxPg.Database.BeginTransactionAsync();

                var clienteInfo = await ctxPg.Empresa
                    .AsNoTracking()
                    .Where(e => e.EmpresaId == _empresaIdSelecionada.Value)
                    .Select(e => new
                    {
                        Razao = e.Razaosocial ?? e.Nome ?? "",
                        Sigla = e.Sigla ?? ""
                    })
                    .FirstOrDefaultAsync();

                var razao = (clienteInfo?.Razao ?? "").Trim();
                var sigla = (clienteInfo?.Sigla ?? "").Trim();

                if (string.IsNullOrWhiteSpace(razao)) razao = "SEM RAZAO SOCIAL";
                if (string.IsNullOrWhiteSpace(sigla)) sigla = "SEM SIGLA";

                byte[] bytesImagem = temClipboard
                    ? _imagemBytes!
                    : await File.ReadAllBytesAsync(_imagemLocalPath!);

                var projeto = new global::Controle_Pedidos.Entities_GM.Projeto
                {
                    ModeloId = _modeloIdProjeto.Value,
                    NroRevisao = _revProjetoAtual,

                    ColaboradorAprovadorId = "1",
                    Descricao = (txtDescricaoHeader.Text ?? "").Trim(),
                    Elaborador = "projeto",
                    Numeronocliente = (txtModeloHeader.Text ?? "").Trim(),
                    EmpresaId = _empresaIdSelecionada.Value,
                    Responsavel = "SMR",

                    PesoBruto = StripSuffixER(txtPBHeader.Text),
                    PesoLiquido = StripSuffixER(txtPLHeader.Text),
                    Rendimento = StripSuffixER(txtRendimentoHeader.Text),

                    // ✅ status já calculado no clique (Novo Projeto / Nova Revisão)
                    Status = _statusProjetoParaSalvar,

                    DataCriacao = EnsureUtc(DateTime.Today),
                    ColaboradorAprovadorSigla = "SMR",

                    Razaosocial = razao,
                    Sigla = sigla,

                    DadosTecnicos = txtRevisao?.Rtf ?? ""
                };

                ctxPg.Projeto.Add(projeto);
                await ctxPg.SaveChangesAsync();

                if (_edicaoEhRevisao && _projetoIdReprovadoAberto.HasValue)
                {
                    var projReprovado = await ctxPg.Projeto
                        .FirstOrDefaultAsync(p => p.ProjetoId == _projetoIdReprovadoAberto.Value);

                    if (projReprovado != null)
                    {
                        var prop = projReprovado.GetType().GetProperty("Status");
                        if (prop != null)
                        {
                            if (prop.PropertyType == typeof(char) || prop.PropertyType == typeof(char?))
                                prop.SetValue(projReprovado, 'H');
                            else if (prop.PropertyType == typeof(string))
                                prop.SetValue(projReprovado, "HR");
                        }

                        await ctxPg.SaveChangesAsync();
                    }

                    _projetoIdReprovadoAberto = null;
                }

                await AtualizarCardProjetosStatusRAsync();

                _projetoIdAtual = projeto.ProjetoId;

                bool precisaMagma = (_statusProjetoParaSalvar == 'B' || _statusProjetoParaSalvar == 'C');

                if (precisaMagma)
                {
                    var resp = MessageBox.Show(
                        "Liga nobre identificada (Status B/C).\n\nSelecione a IMAGEM DA SIMULAÇÃO (Magma) para continuar.\n\n" +
                        "Se cancelar, o projeto NÃO será salvo.",
                        "Imagem Magma obrigatória",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning);

                    if (resp != DialogResult.OK)
                    {
                        await tx.RollbackAsync();
                        MessageBox.Show("Operação cancelada. O projeto NÃO foi salvo.");
                        return;
                    }

                    using var ofdMagma = new OpenFileDialog();
                    ofdMagma.Title = "Selecione a imagem da Simulação (Magma)";
                    ofdMagma.Filter = "Imagens|*.jpg;*.jpeg;*.png;*.bmp;*.webp";

                    if (ofdMagma.ShowDialog(this) != DialogResult.OK)
                    {
                        await tx.RollbackAsync();
                        MessageBox.Show("Operação cancelada. O projeto NÃO foi salvo.");
                        return;
                    }

                    string blobBaseName = $"{_projetoIdAtual}_{_modeloIdProjeto}_{_revProjetoAtual}";
                    string urlMagma = await UploadImagemMagmaAsync(blobBaseName, ofdMagma.FileName);

                    if (!TrySetProjetoImagemMagma(projeto, urlMagma))
                        throw new Exception("Não encontrei a propriedade/coluna Imagem_Magma no entity Projeto.");

                    await ctxPg.SaveChangesAsync();

                    _imagemMagmaUrlAtual = urlMagma;
                    AtualizarCardMagmaUI();
                }

                var ext = string.IsNullOrWhiteSpace(_imagemExt) ? ".png" : _imagemExt!;
                var blobName = $"{_projetoIdAtual}_{_modeloIdProjeto}_{_revProjetoAtual}{ext}";
                _ = await UploadImagemProjetoAsync(blobName, bytesImagem);

                var antigas = await ctxPg.ProjetoImagem
                    .Where(x => x.ModeloId == _modeloIdProjeto.Value && x.ImagemPadrao == 1)
                    .ToListAsync();

                foreach (var img in antigas)
                    img.ImagemPadrao = 0;

                var nomeImg = !string.IsNullOrWhiteSpace(_imagemNome)
                    ? _imagemNome!
                    : (temArquivo ? Path.GetFileName(_imagemLocalPath!) : "clipboard.png");

                var pi = new global::Controle_Pedidos.Entities_GM.ProjetoImagem
                {
                    ModeloId = _modeloIdProjeto.Value,
                    NroRevisao = _revProjetoAtual,
                    ImagemNome = nomeImg,
                    ImagemPadrao = 1,
                    Imagem = bytesImagem
                };

                ctxPg.ProjetoImagem.Add(pi);
                await ctxPg.SaveChangesAsync();

                await EnviarEmailNotificacaoProjetoAsync(
                    ehRevisao: _edicaoEhRevisao,
                    modeloId: _modeloIdProjeto.Value,
                    projetoId: _projetoIdAtual.Value,
                    nroRevisao: _revProjetoAtual,
                    clienteSigla: sigla,
                    clienteRazao: razao,
                    numeroModelo: (txtModeloHeader.Text ?? "").Trim(),
                    descricaoModelo: (txtDescricaoHeader.Text ?? "").Trim()
                );

                await tx.CommitAsync();

                _edicaoProjetoAtiva = false;
                _edicaoEhRevisao = false;
                txtRevisao.ReadOnly = true;

                _imagemLocalPath = null;
                _imagemBytes = null;
                _imagemNome = null;

                MessageBox.Show("Projeto/Revisão salvo com sucesso (PostGre)!");
                await CarregarDetalhesAsync(_modeloIdProjeto.Value);
            }
            catch (DbUpdateException dbEx)
            {
                var real = dbEx.GetBaseException().Message;
                MessageBox.Show("Erro ao salvar projeto (PostGre):\n\n" + real + "\n\n" + dbEx.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao salvar:\n\n" + ex.ToString());
            }
            finally
            {
                AtualizarUIEstado();
            }
        }

        // Upload por BYTES (images-fa)
        private async Task<string> UploadImagemProjetoAsync(string blobName, byte[] bytes)
        {
            string containerUrlWithSas = $"{_baseUrl}{_containerName}?{_sasToken}";
            var containerClient = new BlobContainerClient(new Uri(containerUrlWithSas));
            var blobClient = containerClient.GetBlobClient(blobName);

            using var ms = new MemoryStream(bytes);
            await blobClient.UploadAsync(ms, overwrite: false);

            return $"{_baseUrl}{_containerName}/{blobName}";
        }

        // Upload imagem Magma
        private async Task<string> UploadImagemMagmaAsync(string blobNameSemExt, string filePath)
        {
            var containerUrlWithSas = $"{_magmaBaseUrl}{_magmaContainer}?{_magmaSas}";
            var containerClient = new BlobContainerClient(new Uri(containerUrlWithSas));

            var ext = Path.GetExtension(filePath);
            if (string.IsNullOrWhiteSpace(ext)) ext = ".png";

            var blobName = $"{blobNameSemExt}{ext}".Replace("\\", "/");
            var blobClient = containerClient.GetBlobClient(blobName);

            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(25));

            try
            {
                await blobClient.UploadAsync(filePath, overwrite: true, cancellationToken: cts.Token);
                return $"{_magmaBaseUrl}{_magmaContainer}/{blobName}";
            }
            catch (RequestFailedException rf)
            {
                throw new Exception($"Falha Azure Magma (HTTP {rf.Status}) - {rf.ErrorCode}\n{rf.Message}", rf);
            }
            catch (OperationCanceledException)
            {
                throw new Exception("Timeout ao enviar imagem para o Azure (images-magma). Verifique internet/permissões/SAS.");
            }
        }

        private bool TrySetProjetoImagemMagma(global::Controle_Pedidos.Entities_GM.Projeto projeto, string url)
        {
            try
            {
                var t = projeto.GetType();

                var prop =
                    t.GetProperty("Imagem_Magma") ??
                    t.GetProperty("ImagemMagma") ??
                    t.GetProperty("imagem_magma") ??
                    t.GetProperty("ImagemMagmaUrl");

                if (prop == null) return false;

                if (prop.PropertyType == typeof(string))
                {
                    prop.SetValue(projeto, url);
                    return true;
                }

                prop.SetValue(projeto, Convert.ChangeType(url, prop.PropertyType));
                return true;
            }
            catch
            {
                return false;
            }
        }

        private string? TryGetProjetoImagemMagma(global::Controle_Pedidos.Entities_GM.Projeto? projeto)
        {
            try
            {
                if (projeto == null) return null;

                var t = projeto.GetType();

                var prop =
                    t.GetProperty("Imagem_Magma") ??
                    t.GetProperty("ImagemMagma") ??
                    t.GetProperty("imagem_magma") ??
                    t.GetProperty("ImagemMagmaUrl");

                if (prop == null) return null;

                var val = prop.GetValue(projeto);
                var s = val?.ToString();
                if (string.IsNullOrWhiteSpace(s)) return null;

                return s.Trim();
            }
            catch
            {
                return null;
            }
        }

        private void CancelarEdicaoProjetoUI()
        {
            _edicaoProjetoAtiva = false;
            _edicaoEhRevisao = false;

            _modeloIdProjeto = null;
            _projetoIdAtual = null;
            _revProjetoAtual = 0;

            _imagemLocalPath = null;
            _imagemExt = null;

            _imagemBytes = null;
            _imagemNome = null;

            txtRevisao.ReadOnly = true;

            _statusProjetoParaSalvar = 'D';
        }

        private void SetRichTextSafe(RichTextBox rtb, string value)
        {
            value = value ?? "";

            try
            {
                if (value.TrimStart().StartsWith(@"{\rtf"))
                    rtb.Rtf = value;
                else
                    rtb.Text = value;
            }
            catch
            {
                rtb.Text = value;
            }
        }

        // ✅ EMAIL
        private async Task EnviarEmailNotificacaoProjetoAsync(
            bool ehRevisao,
            int modeloId,
            int projetoId,
            int nroRevisao,
            string clienteSigla,
            string clienteRazao,
            string numeroModelo,
            string descricaoModelo)
        {
            try
            {
                string assunto = ehRevisao
                    ? $"[GM] Nova REVISÃO Pendente Aprovacão - Modelo {numeroModelo} (Rev {nroRevisao})"
                    : $"[GM] Novo PROJETO Pendente Aprovacão - Modelo {numeroModelo} (Rev {nroRevisao})";

                string tipo = ehRevisao ? "UMA NOVA REVISÃO" : "UM NOVO PROJETO";

                string enc(string s) => System.Net.WebUtility.HtmlEncode(s ?? "");

                string msg = $@"
<div style='font-family:Segoe UI, Arial, sans-serif; font-size:14px; color:#222;'>
  <h2 style='margin:0 0 10px 0;'>{tipo} SALVO</h2>
  <p style='margin:0 0 10px 0;'>{tipo.ToLower()} foi salvo no sistema.</p>

  <table style='border-collapse:collapse;'>
    <tr><td style='padding:4px 10px 4px 0;'><b>ProjetoId:</b></td><td style='padding:4px 0;'>{projetoId}</td></tr>
    <tr><td style='padding:4px 10px 4px 0;'><b>ModeloId:</b></td><td style='padding:4px 0;'>{modeloId}</td></tr>
    <tr><td style='padding:4px 10px 4px 0;'><b>Nº Modelo:</b></td><td style='padding:4px 0;'>{enc(numeroModelo)}</td></tr>
    <tr><td style='padding:4px 10px 4px 0;'><b>Revisão:</b></td><td style='padding:4px 0;'>{nroRevisao}</td></tr>
    <tr><td style='padding:4px 10px 4px 0;'><b>Cliente (Sigla):</b></td><td style='padding:4px 0;'>{enc(clienteSigla)}</td></tr>
    <tr><td style='padding:4px 10px 4px 0;'><b>Cliente (Razão):</b></td><td style='padding:4px 0;'>{enc(clienteRazao)}</td></tr>
    <tr><td style='padding:4px 10px 4px 0;'><b>Descrição:</b></td><td style='padding:4px 0;'>{enc(descricaoModelo)}</td></tr>
    <tr><td style='padding:4px 10px 4px 0;'><b>Data:</b></td><td style='padding:4px 0;'>{DateTime.Now:dd/MM/yyyy HH:mm}</td></tr>
  </table>

  <p style='margin:14px 0 0 0; color:#666; font-size:12px;'>(Mensagem automática do Controle_Pedidos)</p>
</div>";

                var mailer = new Controller_Email();
                await mailer.SendMail_sem_anexo(_emailNotificacaoDestino, assunto, msg);
            }
            catch
            {
                // não derruba o salvar se o e-mail falhar
            }
        }

        private async Task<(bool ok, int lastRev, string lastRtf, string pb, string pl, string rend)> GetUltimoProjetoInfoAsync(int modeloId)
        {
            using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

            var ult = await ctxPg.Projeto
                .AsNoTracking()
                .Where(p => p.ModeloId == modeloId)
                .OrderByDescending(p => p.NroRevisao)
                .Select(p => new
                {
                    p.NroRevisao,
                    p.DadosTecnicos,
                    p.PesoBruto,
                    p.PesoLiquido,
                    p.Rendimento
                })
                .FirstOrDefaultAsync();

            if (ult == null)
                return (false, 0, "", "", "", "");

            return (true,
                ult.NroRevisao,
                ult.DadosTecnicos ?? "",
                ult.PesoBruto ?? "",
                ult.PesoLiquido ?? "",
                ult.Rendimento ?? ""
            );
        }

        // =========================
        // HEADER AUTO
        // =========================
        private (string pbText, string plText, string rendText) BuildHeaderPesos(
            float? pbEst, float? pbReal,
            float? plEst, float? plReal)
        {
            bool hasAnyReal = (pbReal.HasValue || plReal.HasValue);
            bool hasAnyEst = (pbEst.HasValue || plEst.HasValue);

            string suffix = hasAnyReal ? " R" : (hasAnyEst ? " E" : "");

            float? pb = pbReal ?? pbEst;
            float? pl = plReal ?? plEst;

            string pbText = pb.HasValue ? (FormatPeso(pb.Value) + suffix) : "";
            string plText = pl.HasValue ? (FormatPeso(pl.Value) + suffix) : "";

            string rendText = "";
            if (pb.HasValue && pl.HasValue && pb.Value > 0.000001f)
            {
                var rendimento = (pl.Value / pb.Value) * 100f;
                rendText = FormatPercent(rendimento) + "%" + suffix;
            }

            return (pbText, plText, rendText);
        }

        private string FormatPeso(float v) => v.ToString("0.###", CultureInfo.CurrentCulture);
        private string FormatPercent(float v) => v.ToString("0.##", CultureInfo.CurrentCulture);

        private void AplicarHeaderPesosDoModelo(float? pbEst, float? pbReal, float? plEst, float? plReal)
        {
            var built = BuildHeaderPesos(pbEst, pbReal, plEst, plReal);
            txtPBHeader.Text = built.pbText;
            txtPLHeader.Text = built.plText;
            txtRendimentoHeader.Text = built.rendText;
        }

        private async Task<(string pbText, string plText, string rendText)> GetHeaderPesosDoModeloAsync(int modeloId)
        {
            using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

            var w = await ctxPg.Modelo
                .AsNoTracking()
                .Where(x => x.ModeloId == modeloId)
                .Select(x => new
                {
                    PB_Est = x.Pesomedio,
                    PL_Est = x.Pesoprevisto,
                    PB_Real = x.Pesobruto,
                    PL_Real = x.Pesoliquido
                })
                .FirstOrDefaultAsync();

            if (w == null)
                return ("", "", "");

            return BuildHeaderPesos(w.PB_Est, w.PB_Real, w.PL_Est, w.PL_Real);
        }

        private string StripSuffixER(string? s)
        {
            s = (s ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return "";

            if (s.EndsWith(" E", StringComparison.OrdinalIgnoreCase) ||
                s.EndsWith(" R", StringComparison.OrdinalIgnoreCase))
            {
                s = s.Substring(0, s.Length - 2).Trim();
            }

            return s;
        }

        // =========================
        // TIPOS DE MODELO (COMBO)
        // =========================
        private class TipoModeloComboItem
        {
            public int TipodemodeloId { get; set; }
            public string Nome { get; set; } = "";
        }

        private async Task CarregarTiposModeloAsync()
        {
            using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

            var tiposRaw = await ctxPg.Tipodemodelo
                .AsNoTracking()
                .Select(t => new
                {
                    t.TipodemodeloId,
                    Nome = (t.Nome ?? "").Trim(),
                    Descricao = (t.Descricao ?? "").Trim(),
                    Sigla = (t.Sigla ?? "").Trim()
                })
                .ToListAsync();

            var tipos = tiposRaw
                .Select(x => new TipoModeloComboItem
                {
                    TipodemodeloId = x.TipodemodeloId,
                    Nome = !string.IsNullOrWhiteSpace(x.Nome) ? x.Nome
                         : !string.IsNullOrWhiteSpace(x.Descricao) ? x.Descricao
                         : !string.IsNullOrWhiteSpace(x.Sigla) ? x.Sigla
                         : ""
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.Nome))
                .OrderBy(x => x.Nome)
                .ToList();

            cmbTipoModelo.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbTipoModelo.DisplayMember = "Nome";
            cmbTipoModelo.ValueMember = "TipodemodeloId";
            cmbTipoModelo.DataSource = tipos;
            cmbTipoModelo.SelectedIndex = -1;
        }

        // =========================
        // HELPERS PARSE
        // =========================
        private int ParseIntOrZero(string? s)
        {
            s = (s ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return 0;

            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.CurrentCulture, out var v))
                return v;

            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out v))
                return v;

            return 0;
        }

        private float? ParseFloatOrNull(string? s)
        {
            s = (s ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return null;

            if (float.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out var v))
                return v;

            if (float.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out v))
                return v;

            return null;
        }

        private char ParseCharOrDefault(string? s, char defaultValue)
        {
            s = (s ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return defaultValue;
            return s[0];
        }

        private char? ParseNullableChar(string? s)
        {
            s = (s ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return null;
            return s[0];
        }

        private static DateTime EnsureUtc(DateTime dt)
        {
            if (dt.Kind == DateTimeKind.Unspecified)
                return DateTime.SpecifyKind(dt, DateTimeKind.Local).ToUniversalTime();

            if (dt.Kind == DateTimeKind.Local)
                return dt.ToUniversalTime();

            return dt;
        }

        // =========================
        // CLIPBOARD HELPERS
        // =========================
        private static (bool ok, Image? img) TryGetClipboardImage()
        {
            try
            {
                if (!Clipboard.ContainsImage())
                    return (false, null);

                var img = Clipboard.GetImage();
                return (img != null, img);
            }
            catch
            {
                return (false, null);
            }
        }

        private static byte[] ImageToPngBytes(Image img)
        {
            using var ms = new MemoryStream();
            img.Save(ms, ImageFormat.Png);
            return ms.ToArray();
        }

        // =========================
        // IMPRIMIR
        // =========================
        private void btnImprimir_Click(object? sender, EventArgs e)
        {
            try
            {
                if (txtClienteHeader == null) throw new Exception("txtClienteHeader é NULL (nome no Designer errado?)");
                if (txtModeloHeader == null) throw new Exception("txtModeloHeader é NULL");
                if (txtDescricaoHeader == null) throw new Exception("txtDescricaoHeader é NULL");
                if (txtPBHeader == null) throw new Exception("txtPBHeader é NULL");
                if (txtPLHeader == null) throw new Exception("txtPLHeader é NULL");
                if (txtRendimentoHeader == null) throw new Exception("txtRendimentoHeader é NULL");
                if (txtRevHeader == null) throw new Exception("txtRevHeader é NULL");
                if (txtDataHeader == null) throw new Exception("txtDataHeader é NULL");
                if (txtAprovHeader == null) throw new Exception("txtAprovHeader é NULL");

                if (txtRevisao == null) throw new Exception("txtRevisao é NULL (RichTextBox não existe com esse nome?)");
                if (picModelo == null) throw new Exception("picModelo é NULL (PictureBox não existe com esse nome?)");

                var imgPath = CriarImagemTempFileFromPictureBox(picModelo);

                var parametros = new Dictionary<string, string>
                {
                    ["Cliente"] = txtClienteHeader.Text ?? "",
                    ["Modelo"] = txtModeloHeader.Text ?? "",
                    ["Descricao"] = txtDescricaoHeader.Text ?? "",

                    ["PB"] = txtPBHeader.Text ?? "",
                    ["PL"] = txtPLHeader.Text ?? "",
                    ["Rendimento"] = txtRendimentoHeader.Text ?? "",
                    ["Rev2"] = txtRevHeader2.Text ?? "",
                    ["Elab2"] = txtElabHeader2.Text ?? "",
                    ["Data2"] = txtDataHeader2.Text ?? "",
                    ["Aprov2"] = txtAprovHeader2.Text ?? "",

                    ["Rev"] = txtRevHeader.Text ?? "",
                    ["Data"] = txtDataHeader.Text ?? "",
                    ["Aprov"] = txtAprovHeader.Text ?? "",

                    ["Elaborador"] = txtElabHeader.Text ?? "",

                    ["DadosTecnicos"] = txtRevisao.Text ?? "",

                    ["ImagemPath"] = string.IsNullOrWhiteSpace(imgPath) ? "" : new Uri(imgPath).AbsoluteUri
                };

                using var frm = new FrmImpressaoProjeto(parametros);
                frm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Erro ao abrir impressão:\n\n" +
                    ex.GetBaseException().Message +
                    "\n\nSTACK:\n" + ex.StackTrace
                );
            }
        }

        private static string CriarImagemTempFileFromPictureBox(PictureBox pb)
        {
            if (pb?.Image == null) return "";

            var dir = Path.Combine(Path.GetTempPath(), "GM_Relatorios");
            Directory.CreateDirectory(dir);

            var file = Path.Combine(dir, "projeto_" + Guid.NewGuid().ToString("N") + ".png");
            pb.Image.Save(file, System.Drawing.Imaging.ImageFormat.Png);
            return file;
        }

        // =========================
        // CARD Projetos R
        // =========================
        private async Task AtualizarCardProjetosStatusRAsync()
        {
            try
            {
                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                _qtdProjetosStatusR = await ctxPg.Projeto
                    .AsNoTracking()
                    .CountAsync(p => p.Status == 'R');

                lblProjetosRCount.Text = _qtdProjetosStatusR.ToString();

                bool tem = _qtdProjetosStatusR > 0;
                pnlCardProjetosR.Visible = tem;

                if (tem)
                {
                    pnlCardProjetosR.BackColor = Color.Firebrick;
                    lblProjetosRTitle.ForeColor = Color.White;
                    lblProjetosRCount.ForeColor = Color.White;

                    pnlCardProjetosR.Enabled = true;
                    pnlCardProjetosR.Cursor = Cursors.Hand;
                    lblProjetosRTitle.Cursor = Cursors.Hand;
                    lblProjetosRCount.Cursor = Cursors.Hand;
                }
                else
                {
                    pnlCardProjetosR.Enabled = false;
                    pnlCardProjetosR.Cursor = Cursors.Default;
                    lblProjetosRTitle.Cursor = Cursors.Default;
                    lblProjetosRCount.Cursor = Cursors.Default;
                }
            }
            catch
            {
                try
                {
                    lblProjetosRCount.Text = "-";
                    pnlCardProjetosR.Visible = false;
                }
                catch { }
            }
        }

        internal class ProjetoStatusRRow
        {
            public int ProjetoId { get; set; }
            public int EmpresaId { get; set; }
            public int ModeloId { get; set; }
            public int NroRevisao { get; set; }
            public DateTime? DataCriacao { get; set; }
            public string Sigla { get; set; } = "";
            public string RazaoSocial { get; set; } = "";
            public string NumeroNoCliente { get; set; } = "";
            public string Descricao { get; set; } = "";
        }

        private async Task<List<ProjetoStatusRRow>> ListarProjetosStatusRAsync()
        {
            using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

            var lista = await ctxPg.Projeto
                .AsNoTracking()
                .Where(p => p.Status == 'R')
                .OrderByDescending(p => p.DataCriacao)
                .Select(p => new ProjetoStatusRRow
                {
                    ProjetoId = p.ProjetoId,
                    EmpresaId = p.EmpresaId,
                    ModeloId = p.ModeloId,
                    NroRevisao = p.NroRevisao,
                    DataCriacao = p.DataCriacao,

                    Sigla = p.Sigla ?? "",
                    RazaoSocial = p.Razaosocial ?? "",
                    NumeroNoCliente = p.Numeronocliente ?? "",
                    Descricao = p.Descricao ?? ""
                })
                .ToListAsync();

            return lista;
        }

        // =========================
        // (Se você já tem esses handlers no Designer, deixa)
        // =========================
        private void txtCliente_TextChanged(object sender, EventArgs e) { }
        private void label1_Click(object sender, EventArgs e) { }
        private void button1_Click(object sender, EventArgs e) { }
        private void label1_Click_1(object sender, EventArgs e) { }
        private void label2_Click(object sender, EventArgs e) { }
        private void label3_Click(object sender, EventArgs e) { }
        private void label4_Click(object sender, EventArgs e) { }
        private void label4_Click_1(object sender, EventArgs e) { }
        private void label6_Click(object sender, EventArgs e) { }
        private void label8_Click(object sender, EventArgs e) { }
        private void label9_Click(object sender, EventArgs e) { }
        private void label7_Click(object sender, EventArgs e) { }
        private void label10_Click(object sender, EventArgs e) { }
        private void label12_Click(object sender, EventArgs e) { }
        private void label13_Click(object sender, EventArgs e) { }
        private void label13_Click_1(object sender, EventArgs e) { }
        private void button3_Click(object sender, EventArgs e) { }
        private void label23_Click(object sender, EventArgs e) { }
        private void label28_Click(object sender, EventArgs e) { }
        private void label30_Click(object sender, EventArgs e) { }
        private void pictureBox2_Click(object sender, EventArgs e) { }

        private void label33_Click(object sender, EventArgs e)
        {
            pnlCardMagma_Click(sender, e);
        }

        private async void pnlCardProjetosR_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_edicaoProjetoAtiva)
                    return;

                await AtualizarCardProjetosStatusRAsync();

                if (_qtdProjetosStatusR <= 0)
                    return;

                var lista = await ListarProjetosStatusRAsync();
                if (lista == null || lista.Count == 0)
                    return;

                ProjetoStatusRRow? selecionado = null;

                using (var frm = new FrmProjetosStatusR(lista))
                {
                    var dr = frm.ShowDialog(this);
                    if (dr != DialogResult.OK)
                        return;

                    selecionado = frm.ProjetoSelecionado;
                }

                if (selecionado == null)
                    return;

                _projetoIdReprovadoAberto = selecionado.ProjetoId;

                await Task.Yield();

                this.BeginInvoke(new Action(async () =>
                {
                    try
                    {
                        await CarregarClienteESelecionarModeloAsync(
                            empresaId: selecionado.EmpresaId,
                            modeloId: selecionado.ModeloId
                        );

                        await AtualizarCardProjetosStatusRAsync();
                    }
                    catch (Exception ex2)
                    {
                        MessageBox.Show("Erro ao carregar projeto selecionado: " + ex2.GetBaseException().Message);
                    }
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao abrir pendências (Status R): " + ex.GetBaseException().Message);
            }
        }

        private async Task CarregarClienteESelecionarModeloAsync(int empresaId, int modeloId)
        {
            using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

            var cli = await ctxPg.Empresa
                .AsNoTracking()
                .Where(e => e.EmpresaId == empresaId)
                .Select(e => new
                {
                    Nome = (e.Sigla ?? e.Nome ?? "").Trim()
                })
                .FirstOrDefaultAsync();

            _empresaIdSelecionada = empresaId;
            txtCliente.Text = cli?.Nome ?? "";

            txtFiltroModelo.Text = "";

            _mode = FormMode.View;
            _modeloIdEmEdicao = null;

            CancelarEdicaoProjetoUI();

            SetCamposEditaveis(false);
            dgvModelos.Enabled = true;

            await BuscarModelosAsync("");

            SelecionarModeloNoGrid(modeloId);

            await CarregarDetalhesAsync(modeloId);
        }

        internal class FrmProjetosStatusR : Form
        {
            private readonly DataGridView dgv = new DataGridView();
            private readonly Button btnOk = new Button();
            private readonly Button btnCancelar = new Button();

            public Cadastro_Modelo.ProjetoStatusRRow? ProjetoSelecionado { get; private set; }

            public FrmProjetosStatusR(List<Cadastro_Modelo.ProjetoStatusRRow> lista)
            {
                this.Text = "Projetos em Status R";
                this.StartPosition = FormStartPosition.CenterParent;
                this.Width = 1000;
                this.Height = 520;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ShowInTaskbar = false;

                dgv.Dock = DockStyle.Top;
                dgv.Height = 420;
                dgv.ReadOnly = true;
                dgv.AllowUserToAddRows = false;
                dgv.AllowUserToDeleteRows = false;
                dgv.AllowUserToResizeRows = false;
                dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dgv.MultiSelect = false;
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv.RowHeadersVisible = false;

                dgv.DataSource = lista;

                btnOk.Text = "Selecionar";
                btnOk.Width = 120;
                btnOk.Height = 35;
                btnOk.Left = this.ClientSize.Width - 260;
                btnOk.Top = 430;
                btnOk.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;

                btnCancelar.Text = "Cancelar";
                btnCancelar.Width = 120;
                btnCancelar.Height = 35;
                btnCancelar.Left = this.ClientSize.Width - 130;
                btnCancelar.Top = 430;
                btnCancelar.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;

                btnOk.DialogResult = DialogResult.OK;
                btnCancelar.DialogResult = DialogResult.Cancel;

                this.AcceptButton = btnOk;
                this.CancelButton = btnCancelar;

                btnOk.Click += (s, e) => ConfirmarSelecao();

                dgv.MouseDoubleClick += dgv_MouseDoubleClick;
                dgv.KeyDown += dgv_KeyDown;

                this.Controls.Add(dgv);
                this.Controls.Add(btnOk);
                this.Controls.Add(btnCancelar);

                if (dgv.Rows.Count > 0)
                {
                    dgv.ClearSelection();
                    dgv.Rows[0].Selected = true;

                    var cell = dgv.Rows[0].Cells.Cast<DataGridViewCell>().FirstOrDefault();
                    if (cell != null) dgv.CurrentCell = cell;
                }
            }

            private void dgv_MouseDoubleClick(object? sender, MouseEventArgs e)
            {
                var hit = dgv.HitTest(e.X, e.Y);
                if (hit.RowIndex < 0) return;

                dgv.ClearSelection();
                dgv.Rows[hit.RowIndex].Selected = true;

                if (hit.ColumnIndex >= 0)
                    dgv.CurrentCell = dgv.Rows[hit.RowIndex].Cells[hit.ColumnIndex];

                btnOk.PerformClick();
            }

            private void dgv_KeyDown(object? sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    btnOk.PerformClick();
                }
            }

            private void ConfirmarSelecao()
            {
                if (dgv.CurrentRow?.DataBoundItem is Cadastro_Modelo.ProjetoStatusRRow row)
                {
                    ProjetoSelecionado = row;

                    this.Hide();

                    this.BeginInvoke(new Action(() =>
                    {
                        this.DialogResult = DialogResult.OK;
                        this.Close();
                        this.Dispose();
                    }));
                }
                else
                {
                    MessageBox.Show("Selecione uma linha.");
                }
            }
        }

        // =========================================================
        // ✅ POPUP: EXIBE IMAGEM MAGMA (somente UI)
        // =========================================================
        internal class FrmImagemMagma : Form
        {
            private readonly PictureBox pic = new PictureBox();
            private readonly Label lbl = new Label();
            private readonly Button btnFechar = new Button();

            private readonly string _url;
            private readonly string _sas;

            public FrmImagemMagma(string url, string magmaSas, IWin32Window parent)
            {
                _url = (url ?? "").Trim();
                _sas = (magmaSas ?? "").Trim().TrimStart('?');

                Text = "Imagem Magma";
                StartPosition = FormStartPosition.CenterParent;
                FormBorderStyle = FormBorderStyle.Sizable;
                MinimizeBox = false;
                WindowState = FormWindowState.Maximized;

                lbl.Dock = DockStyle.Top;
                lbl.Height = 34;
                lbl.TextAlign = ContentAlignment.MiddleLeft;
                lbl.Padding = new Padding(10, 0, 10, 0);
                lbl.Text = "Carregando imagem...";

                btnFechar.Text = "Fechar";
                btnFechar.Width = 120;
                btnFechar.Height = 34;
                btnFechar.Top = 0;
                btnFechar.Left = this.ClientSize.Width - 130;
                btnFechar.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                btnFechar.Click += (s, e) => this.Close();

                pic.Dock = DockStyle.Fill;
                pic.SizeMode = PictureBoxSizeMode.Zoom;
                pic.BackColor = Color.Black;

                Controls.Add(pic);
                Controls.Add(btnFechar);
                Controls.Add(lbl);

                Shown += async (s, e) => await CarregarAsync();
                FormClosed += (s, e) =>
                {
                    if (pic.Image != null)
                    {
                        var old = pic.Image;
                        pic.Image = null;
                        old.Dispose();
                    }
                };
            }

            private async Task CarregarAsync()
            {
                try
                {
                    var finalUrl = MontarUrlComSasSeNecessario(_url);

                    using var http = new HttpClient();
                    http.Timeout = TimeSpan.FromSeconds(20);

                    var bytes = await http.GetByteArrayAsync(finalUrl);

                    using var ms = new MemoryStream(bytes);
                    using var img = Image.FromStream(ms);
                    var bmp = new Bitmap(img);

                    if (pic.Image != null)
                    {
                        var old = pic.Image;
                        pic.Image = null;
                        old.Dispose();
                    }

                    pic.Image = bmp;
                    lbl.Text = "OK";
                }
                catch (Exception ex)
                {
                    lbl.Text = "Falha ao carregar";
                    MessageBox.Show("Não foi possível carregar a imagem Magma.\n\n" + ex.GetBaseException().Message);
                }
            }

            private string MontarUrlComSasSeNecessario(string url)
            {
                if (url.Contains("?"))
                    return url;

                if (!string.IsNullOrWhiteSpace(_sas))
                    return url + "?" + _sas;

                return url;
            }
        }
    }
}
