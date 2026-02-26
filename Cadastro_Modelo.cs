// Cadastro_Modelo.cs
// ✅ 100% completo (1 arquivo) baseado no seu código
// ✅ AGORA:
//   - Writes no MySQL (grupometalContext) ficam COMENTADOS
//   - TODA LEITURA (grid + detalhes + projeto + imagem + tipos) vem do PostGre (gmetalContext)
//   - Inserts/Updates/Deletes acontecem APENAS no PostGre
// ✅ EXTRA:
//   - Ao salvar NOVO PROJETO ou NOVA REVISÃO -> envia e-mail para Joao.rossitti@grupometal.com.br
// ✅ NOVO (Regra Produção/Liga nobre):
//   - Ao clicar "Novo Projeto" ou "Nova Revisão" -> checa producao_item nas etapas 50/75/100/150/200
//   - Se não encontrar: libera e salva status D (projeto) / E (revisão)
//   - Se encontrar: se QUALQUER item tiver liga com Necessita_Aprovacao_FA = TRUE -> trata como nobre
//       -> status B (projeto) / C (revisão)
//       -> se TODOS não nobres -> status D (projeto) / E (revisão)

using Azure.Storage.Blobs;
using Controle_Pedidos;
using Controle_Pedidos.Controle_Producao.Entities; // (MySQL) mantido só para compatibilidade do projeto, writes comentados
using Controle_Pedidos.Entities_GM;                 // (PostGre) agora é a fonte oficial (leitura + escrita)
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Controle_Pedidos_8.Tela_Cadastro
{
    public partial class Cadastro_Modelo : Form
    {
        private int? _empresaIdSelecionada = null;

        // =========================
        // AZURE
        // =========================
        private const string _baseUrl = "https://armazenamentoazure.blob.core.windows.net/";
        private const string _containerName = "images-fa";
        private const string _sasToken =
            "sp=racwdl&st=2026-02-10T10:11:30Z&se=2030-01-01T18:26:30Z&spr=https&sv=2024-11-04&sr=c&sig=Z2K2P0XFY5k23tXWJx7W5v9wECugxn2S3Fzne9mAqvg%3D";

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

        // ✅ Status calculado no clique (Novo Projeto / Nova Revisão) de acordo com produção/liga
        private char _statusProjetoParaSalvar = 'D';

        // =========================
        // CONTROLE: Projeto reprovado (Status R) que foi aberto pelo CARD
        // =========================
        private int? _projetoIdReprovadoAberto = null;

        // ===== ORIGEM IMAGEM =====
        // Novo Projeto (arquivo) continua igual
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
            if (_edicaoProjetoAtiva) return; // não busca durante criação de projeto/revisão

            _timerBuscaModelo.Stop();
            _timerBuscaModelo.Start();
        }

        // =========================
        // BUSCAR MODELOS (AGORA: PostGre)
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
            if (_edicaoProjetoAtiva) return; // trava troca durante criação de projeto/revisão

            if (dgvModelos.CurrentRow?.DataBoundItem is not ModeloGridRow row)
                return;

            AtualizarUIEstado();
            await CarregarDetalhesAsync(row.ModeloId);
        }

        // =========================
        // CARREGAR DETALHES DO MODELO + HEADER + RTF + IMAGEM (AGORA: PostGre)
        // =========================
        private async Task CarregarDetalhesAsync(int modeloId)
        {
            _ctsCarregarDetalhes?.Cancel();
            _ctsCarregarDetalhes = new CancellationTokenSource();
            var token = _ctsCarregarDetalhes.Token;

            try
            {
                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                // ✅ evita "estado fantasma" de outro modelo
                _temProjetoAtual = false;
                _revUltimoProjeto = 0;

                // 1) Detalhes do modelo
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

                        // Estimados / Reais do modelo
                        x.Pesomedio,     // PB estimado
                        x.Pesobruto,     // PB real
                        x.Pesoprevisto,  // PL estimado
                        x.Pesoliquido    // PL real
                    })
                    .FirstOrDefaultAsync(token);

                if (token.IsCancellationRequested) return;

                if (m == null)
                {
                    LimparDetalhes();
                    return;
                }

                // Preenche controles
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

                // Pesos do modelo
                txtPesoBrutoEst.Text = m.Pesomedio?.ToString() ?? "";
                txtPesoBrutoReal.Text = m.Pesobruto?.ToString() ?? "";
                txtPesoLiquidoEst.Text = m.Pesoprevisto?.ToString() ?? "";
                txtPesoLiquidoReal.Text = m.Pesoliquido?.ToString() ?? "";

                // 2) Cliente (sigla/nome)
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

                // 3) Duas últimas revisões (header 1 e 2)
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

                // 4) Último projeto (para RTF atual e flag de projeto existente)
                var ultProj = await ctxPg.Projeto
                    .AsNoTracking()
                    .Where(p => p.ModeloId == modeloId)
                    .OrderByDescending(p => p.NroRevisao)
                    .Select(p => new
                    {
                        p.NroRevisao,
                        p.DadosTecnicos
                    })
                    .FirstOrDefaultAsync(token);

                if (token.IsCancellationRequested) return;

                _temProjetoAtual = ultProj != null;
                _revUltimoProjeto = ultProj?.NroRevisao ?? 0;

                // Header esquerda da imagem
                txtClienteHeader.Text = nomeCliente;
                txtDescricaoHeader.Text = m.Descricao ?? "";
                txtModeloHeader.Text = m.Numeronocliente ?? "";

                // ✅ HEADER AUTOMÁTICO: Real > Estimado, sufixo R/E + rendimento (PL/PB)
                AplicarHeaderPesosDoModelo(
                    pbEst: m.Pesomedio,
                    pbReal: m.Pesobruto,
                    plEst: m.Pesoprevisto,
                    plReal: m.Pesoliquido
                );

                // Headers de revisões (1 e 2)
                txtRevHeader.Text = p1?.NroRevisao.ToString() ?? "";
                txtRevHeader2.Text = p2?.NroRevisao.ToString() ?? "";

                txtElabHeader.Text = p1?.Elaborador ?? "";
                txtElabHeader2.Text = p2?.Elaborador ?? "";

                txtDataHeader.Text = p1?.DataCriacao.ToString("dd/MM/yyyy") ?? "";
                txtDataHeader2.Text = p2?.DataCriacao.ToString("dd/MM/yyyy") ?? "";

                txtAprovHeader.Text = p1?.ColaboradorAprovadorSigla ?? "";
                txtAprovHeader2.Text = p2?.ColaboradorAprovadorSigla ?? "";

                // ✅ RTF do "projeto atual" (última revisão), só se NÃO estiver em edição
                if (!_edicaoProjetoAtiva)
                {
                    SetRichTextSafe(txtRevisao, ultProj?.DadosTecnicos ?? "");
                    txtRevisao.ReadOnly = true;
                }

                // 5) Imagem: agora carrega do PostGre (bytea)
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
        // IMAGEM: busca bytes (bytea) em ProjetoImagem (PostGre) e mostra
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

                // prioridade: imagem padrão (1), depois maior revisão
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

        // =========================
        // LIMPAR DETALHES
        // =========================
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
        // SALVAR MODELO (AGORA: PostGre)
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

                // ========= UI -> Entity =========
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

                // recarrega lista no PG e seleciona
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

        // =========================
        // REMOVER MODELO (AGORA: PostGre) + MySQL comentado
        // =========================
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

                // ✅ remove no PostGre
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
            // TextBox: ReadOnly (cadastro do modelo)
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

            // HEADER DO PROJETO: editável SOMENTE durante _edicaoProjetoAtiva
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

            // CRUD modelo
            btnAdicionar.Enabled = isView && !_edicaoProjetoAtiva;
            btnEditar.Enabled = isView && dgvModelos.CurrentRow != null && !_edicaoProjetoAtiva;
            btnRemover.Enabled = isView && dgvModelos.CurrentRow != null && !_edicaoProjetoAtiva;

            btnSalvar.Enabled = isEditOrAdd;
            btnCancelar.Enabled = isEditOrAdd;
            btnSalvar.Visible = isEditOrAdd;
            btnCancelar.Visible = isEditOrAdd;

            // projeto/revisão
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
        // ✅ NOVO: REGRA Produção/Liga nobre => status do projeto/revisão
        // =========================================================

        // Etapas "pedido em aberto"
        private static bool IsEtapaAberta(int? etapaId)
        {
            if (!etapaId.HasValue) return false;

            // evita Contains/inferência/Npgsql
            return etapaId.Value == 50
                || etapaId.Value == 75
                || etapaId.Value == 100
                || etapaId.Value == 150
                || etapaId.Value == 200;
        }

        // Retorna o status para salvar, conforme regra:
        // - Sem pedido (ou sem nobre): Projeto 'D' / Revisão 'E'
        // - Com nobre (QUALQUER UM):  Projeto 'B' / Revisão 'C'
        private async Task<char> CalcularStatusPorPedidoELigaAsync(int modeloId, bool ehRevisao)
        {
            char statusSemNobre = ehRevisao ? 'E' : 'D';
            char statusComNobre = ehRevisao ? 'C' : 'B';

            try
            {
                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                // 1) Existe algum pedido aberto nessas etapas?
                // 2) Se existir, QUALQUER UM tem liga nobre? (Necessita_Aprovacao_FA = TRUE)
                //    -> se 15 não nobre + 1 nobre => trata como nobre.
                //
                // Fazemos um Any() com JOIN direto para ficar 100% server-side.
                bool temLigaNobre = await (
                    from pi in ctxPg.ProducaoItem.AsNoTracking()
                    join l in ctxPg.Liga.AsNoTracking() on pi.LigaId equals l.LigaId
                    where pi.ModeloId == modeloId
                          && (pi.EtapaId == 50 || pi.EtapaId == 75 || pi.EtapaId == 100 || pi.EtapaId == 150 || pi.EtapaId == 200)
                          && l.Necessita_Aprovacao_FA == true
                    select 1
                ).AnyAsync();

                if (temLigaNobre)
                    return statusComNobre;

                // Se não achou liga nobre, ainda pode existir pedido aberto com liga não nobre.
                // Nesse caso retorna status sem nobre (D/E).
                // Se nem pedido existir, também D/E.
                return statusSemNobre;
            }
            catch
            {
                // Se der qualquer erro, não trava o usuário: considera sem nobre (D/E)
                return statusSemNobre;
            }
        }

        // =========================
        // NOVO PROJETO / NOVA REVISÃO (RTF)
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

            // ✅ calcula status ANTES de iniciar (regra produção/liga)
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

            // ✅ calcula status ANTES de iniciar (regra produção/liga)
            _statusProjetoParaSalvar = await CalcularStatusPorPedidoELigaAsync(row.ModeloId, ehRevisao: true);

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
        // SALVAR PROJETO/REVISÃO (AGORA: PostGre) + EMAIL
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

                // bytes da imagem (para gravar no PostGre)
                byte[] bytesImagem = temClipboard
                    ? _imagemBytes!
                    : await File.ReadAllBytesAsync(_imagemLocalPath!);

                // ✅ cria nova linha em Projeto (PostGre)
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

                    // ✅ AQUI: status conforme regra (B/C/D/E)
                    Status = _statusProjetoParaSalvar,

                    DataCriacao = EnsureUtc(DateTime.Today),
                    ColaboradorAprovadorSigla = "SMR",

                    Razaosocial = razao,
                    Sigla = sigla,

                    DadosTecnicos = txtRevisao?.Rtf ?? ""
                };

                ctxPg.Projeto.Add(projeto);
                await ctxPg.SaveChangesAsync();

                // ✅ Se esta revisão foi criada a partir de um projeto reprovado (R),
                // então muda o status do projeto antigo para HR.
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
                            {
                                // Se Status for char, não existe "HR". Use 'H'.
                                prop.SetValue(projReprovado, 'H');
                            }
                            else if (prop.PropertyType == typeof(string))
                            {
                                prop.SetValue(projReprovado, "HR");
                            }
                        }

                        await ctxPg.SaveChangesAsync();
                    }

                    _projetoIdReprovadoAberto = null;
                }

                await AtualizarCardProjetosStatusRAsync();

                _projetoIdAtual = projeto.ProjetoId;

                // ✅ upload para Azure (mantido)
                var ext = string.IsNullOrWhiteSpace(_imagemExt) ? ".png" : _imagemExt!;
                var blobName = $"{_projetoIdAtual}_{_modeloIdProjeto}_{_revProjetoAtual}{ext}";
                _ = await UploadImagemProjetoAsync(blobName, bytesImagem);

                // zera imagem padrão anterior (PostGre)
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

                // ✅ envia email (depois de gravar Projeto + ProjetoImagem com sucesso)
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
                MessageBox.Show("Erro ao salvar projeto (PostGre):\n\n" + real);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao salvar projeto (PostGre): " + ex.Message);
            }
            finally
            {
                AtualizarUIEstado();
            }
        }

        // Upload por BYTES (usado para arquivo ou clipboard)
        private async Task<string> UploadImagemProjetoAsync(string blobName, byte[] bytes)
        {
            string containerUrlWithSas = $"{_baseUrl}{_containerName}?{_sasToken}";
            var containerClient = new BlobContainerClient(new Uri(containerUrlWithSas));
            var blobClient = containerClient.GetBlobClient(blobName);

            using var ms = new MemoryStream(bytes);
            await blobClient.UploadAsync(ms, overwrite: false);

            return $"{_baseUrl}{_containerName}/{blobName}";
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

            // reseta status default (não interfere em nada)
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

        // ✅ EMAIL: Monta assunto + HTML e envia (sem anexo)
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
  <p style='margin:0 0 10px 0;'>
    {tipo.ToLower()} foi salvo no sistema.
  </p>

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
    
  <p style='margin:14px 0 0 0; color:#666; font-size:12px;'>
    (Mensagem automática do Controle_Pedidos)
  </p>
</div>";
                var mailer = new Controller_Email();
                await mailer.SendMail_sem_anexo(_emailNotificacaoDestino, assunto, msg);
            }
            catch
            {
                // não derruba o salvar se o e-mail falhar
            }
        }

        // Último projeto agora lê do PostGre
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
        // ✅ HEADER AUTO (PB/PL/Rendimento)
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

        // Header pesos agora lê do PostGre
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

            while (s.Contains("  ")) s = s.Replace("  ", " ");

            if (s.EndsWith(" E", StringComparison.OrdinalIgnoreCase) ||
                s.EndsWith(" R", StringComparison.OrdinalIgnoreCase))
            {
                s = s.Substring(0, s.Length - 2).Trim();
            }

            return s;
        }

        // =========================
        // TIPOS DE MODELO (COMBO) - agora do PostGre
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
        // ✅ CLIPBOARD HELPERS
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

        private async Task AtualizarCardProjetosStatusRAsync()
        {
            try
            {
                using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

                _qtdProjetosStatusR = await ctxPg.Projeto
                    .AsNoTracking()
                    .CountAsync(p => p.Status == 'R');

                // UI
                lblProjetosRCount.Text = _qtdProjetosStatusR.ToString();

                bool tem = _qtdProjetosStatusR > 0;

                // ✅ Se 0: some
                pnlCardProjetosR.Visible = tem;

                if (tem)
                {
                    // ✅ Se > 0: aparece vermelho
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
        // OPCIONAIS (eventos vazios)
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

        private async void pnlCardProjetosR_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_edicaoProjetoAtiva)
                    return; // sem mensagem

                // Atualiza contagem antes de abrir
                await AtualizarCardProjetosStatusRAsync();

                if (_qtdProjetosStatusR <= 0)
                    return; // sem mensagem

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

                // ✅ guarda qual projeto "R" o usuário abriu
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

            // 1) Busca nome/sigla do cliente
            var cli = await ctxPg.Empresa
                .AsNoTracking()
                .Where(e => e.EmpresaId == empresaId)
                .Select(e => new
                {
                    Nome = (e.Sigla ?? e.Nome ?? "").Trim()
                })
                .FirstOrDefaultAsync();

            // 2) Ajusta estado do form (igual seu fluxo de "buscar cliente")
            _empresaIdSelecionada = empresaId;
            txtCliente.Text = cli?.Nome ?? "";

            txtFiltroModelo.Text = "";

            _mode = FormMode.View;
            _modeloIdEmEdicao = null;

            CancelarEdicaoProjetoUI();

            SetCamposEditaveis(false);
            dgvModelos.Enabled = true;

            // 3) Carrega grid modelos desse cliente
            await BuscarModelosAsync("");

            // 4) Seleciona o modelo do projeto
            SelecionarModeloNoGrid(modeloId);

            // 5) Carrega detalhes (modelo + headers + RTF + imagem)
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

                // ✅ DataGrid
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

                // Headers
                if (dgv.Columns.Contains("EmpresaId")) dgv.Columns["EmpresaId"].HeaderText = "EmpresaId";
                if (dgv.Columns.Contains("ModeloId")) dgv.Columns["ModeloId"].HeaderText = "ModeloId";
                if (dgv.Columns.Contains("ProjetoId")) dgv.Columns["ProjetoId"].HeaderText = "ProjetoId";
                if (dgv.Columns.Contains("Sigla")) dgv.Columns["Sigla"].HeaderText = "Sigla";
                if (dgv.Columns.Contains("RazaoSocial")) dgv.Columns["RazaoSocial"].HeaderText = "Cliente";
                if (dgv.Columns.Contains("NumeroNoCliente")) dgv.Columns["NumeroNoCliente"].HeaderText = "Nº Modelo";
                if (dgv.Columns.Contains("Descricao")) dgv.Columns["Descricao"].HeaderText = "Descrição";
                if (dgv.Columns.Contains("NroRevisao")) dgv.Columns["NroRevisao"].HeaderText = "Rev";
                if (dgv.Columns.Contains("DataCriacao")) dgv.Columns["DataCriacao"].HeaderText = "Data";

                // Botões
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

                // ✅ DialogResult nos botões (fluxo estável)
                btnOk.DialogResult = DialogResult.OK;
                btnCancelar.DialogResult = DialogResult.Cancel;

                this.AcceptButton = btnOk;
                this.CancelButton = btnCancelar;

                // ✅ Eventos
                btnOk.Click += (s, e) => ConfirmarSelecao();

                // 🔥 Duplo clique blindado
                dgv.MouseDoubleClick += dgv_MouseDoubleClick;

                // Enter também seleciona
                dgv.KeyDown += dgv_KeyDown;

                this.Controls.Add(dgv);
                this.Controls.Add(btnOk);
                this.Controls.Add(btnCancelar);

                // Seleciona primeira linha
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

        private void label30_Click(object sender, EventArgs e) { }
        private void pictureBox2_Click(object sender, EventArgs e) { }
    }
}
