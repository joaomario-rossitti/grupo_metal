using Azure.Storage.Blobs;
using Controle_Pedidos;
using Controle_Pedidos.Controle_Producao.Entities;
using Controle_Pedidos.Entities_GM;
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
        // AZURE (MESMO DO MIGRADOR)
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
        // BUSCAR MODELOS
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
                using var ctx = new grupometalContext();

                var query =
                    from m in ctx.Modelo_EF
                    join e in ctx.Empresa_GM on m.ClienteempresaId equals e.EmpresaId
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
                MessageBox.Show("Erro ao buscar modelos: " + ex.Message);
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
        // CARREGAR DETALHES DO MODELO + HEADER + RTF + IMAGEM
        // =========================
        private async Task CarregarDetalhesAsync(int modeloId)
        {
            _ctsCarregarDetalhes?.Cancel();
            _ctsCarregarDetalhes = new CancellationTokenSource();
            var token = _ctsCarregarDetalhes.Token;

            try
            {
                using var ctx = new grupometalContext();

                // ✅ evita "estado fantasma" de outro modelo
                _temProjetoAtual = false;
                _revUltimoProjeto = 0;

                // 1) Detalhes do modelo
                var m = await ctx.Modelo_EF
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

                // Aqui mantém seus TextBoxes do cadastro (modelo)
                txtPesoBrutoEst.Text = m.Pesomedio?.ToString() ?? "";
                txtPesoBrutoReal.Text = m.Pesobruto?.ToString() ?? "";
                txtPesoLiquidoEst.Text = m.Pesoprevisto?.ToString() ?? "";
                txtPesoLiquidoReal.Text = m.Pesoliquido?.ToString() ?? "";

                // 2) Cliente (sigla/nome)
                string nomeCliente = "";
                if (_empresaIdSelecionada != null)
                {
                    nomeCliente = await ctx.Empresa_GM
                        .AsNoTracking()
                        .Where(e => e.EmpresaId == _empresaIdSelecionada.Value)
                        .Select(e => e.Sigla ?? e.Nome ?? "")
                        .FirstOrDefaultAsync(token) ?? "";
                }
                if (token.IsCancellationRequested) return;

                // 3) Duas últimas revisões (header 1 e 2)
                var projetos2 = await ctx.Projeto
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
                var ultProj = await ctx.Projeto
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

                // 5) Imagem (ProjetoImagem)
                await CarregarImagemDoModeloViaProjetoImagemAsync(modeloId, token);
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar detalhes: " + ex.Message);
            }
            finally
            {
                AtualizarUIEstado();
            }
        }

        // =========================
        // IMAGEM: busca ImagemUrl em ProjetoImagem e baixa do Blob
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

                using var ctx = new grupometalContext();

                var imagemUrl = await ctx.ProjetoImagem
                    .AsNoTracking()
                    .Where(x => x.ModeloId == modeloId)
                    .OrderByDescending(x => x.ImagemPadrao)
                    .ThenByDescending(x => x.NroRevisao)
                    .Select(x => x.ImagemUrl)
                    .FirstOrDefaultAsync(ct);

                if (ct.IsCancellationRequested) return;

                if (string.IsNullOrWhiteSpace(imagemUrl))
                {
                    LimparImagem();
                    return;
                }

                var prefix = $"{_baseUrl}{_containerName}/";
                var blobName = imagemUrl.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    ? imagemUrl.Substring(prefix.Length)
                    : imagemUrl;

                await CarregarImagemBlobAsync(blobName, ct);
            }
            catch (OperationCanceledException) { }
            catch
            {
                LimparImagem();
            }
        }

        // =========================
        // BAIXAR IMAGEM DO BLOB (SAS) -> PictureBox
        // =========================
        private async Task CarregarImagemBlobAsync(string blobName, CancellationToken ct)
        {
            string containerUrlWithSas = $"{_baseUrl}{_containerName}?{_sasToken}";
            var containerClient = new BlobContainerClient(new Uri(containerUrlWithSas));
            var blobClient = containerClient.GetBlobClient(blobName);

            if (!await blobClient.ExistsAsync(ct))
            {
                LimparImagem();
                return;
            }

            using var ms = new MemoryStream();
            await blobClient.DownloadToAsync(ms, ct);
            ms.Position = 0;

            if (ct.IsCancellationRequested) return;

            using var temp = Image.FromStream(ms);
            var img = new Bitmap(temp);

            if (picModelo.InvokeRequired)
            {
                picModelo.BeginInvoke(new Action(() => TrocarImagem(img)));
            }
            else
            {
                TrocarImagem(img);
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
                using var ctx = new grupometalContext();

                Modelo_EF entity;

                if (_mode == FormMode.Add)
                {
                    entity = new Modelo_EF();
                    entity.ClienteempresaId = _empresaIdSelecionada.Value;
                    ctx.Modelo_EF.Add(entity);
                }
                else
                {
                    if (_modeloIdEmEdicao == null)
                    {
                        MessageBox.Show("Não foi possível identificar o modelo em edição.");
                        return;
                    }

                    entity = await ctx.Modelo_EF
                        .FirstOrDefaultAsync(x => x.ModeloId == _modeloIdEmEdicao.Value);

                    if (entity == null)
                    {
                        MessageBox.Show("Modelo não encontrado (pode ter sido removido).");
                        await CancelarEdicaoAsync();
                        return;
                    }
                }

                // ========= UI -> Entity =========
                entity.Numeronocliente = (txtNumeroModelo.Text ?? "").Trim();
                entity.Descricao = (txtDescricao.Text ?? "").Trim();
                entity.Observacao = (txtObservacao.Text ?? "").Trim();

                entity.Situacao = ParseCharOrDefault(cmbSituacao.Text, 'A');
                entity.Estado = ParseNullableChar(cmbEstado.Text);

                if (cmbTipoModelo.SelectedValue is int tipoId)
                    entity.TipodemodeloId = tipoId;
                else
                    entity.TipodemodeloId = 0;

                entity.Nrodepartes = ParseIntOrZero(txtNroPartes.Text);
                entity.Numerodefiguras = (txtQtdFigurasPlaca.Text ?? "").Trim();
                entity.Numerodecaixasdemachoporfigura = (txtCxMachoFigura.Text ?? "").Trim();

                entity.Valordomodelo = ParseFloatOrNull(txtValorModelo.Text);

                // Pesos do MODELO (estimados/reais)
                entity.Pesomedio = ParseFloatOrNull(txtPesoBrutoEst.Text);
                entity.Pesobruto = ParseFloatOrNull(txtPesoBrutoReal.Text);
                entity.Pesoprevisto = ParseFloatOrNull(txtPesoLiquidoEst.Text);
                entity.Pesoliquido = ParseFloatOrNull(txtPesoLiquidoReal.Text);
                // =================================

                // ✅ salva no MySQL
                await ctx.SaveChangesAsync();

                int savedId = entity.ModeloId;

                // ✅ replica no PostGre (upsert)
                try
                {
                    await ReplicarModeloParaPostgreAsync(entity);
                }
                catch (Exception exPg)
                {
                    MessageBox.Show("⚠ Modelo salvo no MySQL, mas falhou ao replicar no PostGre:\n\n" + exPg.GetBaseException().Message);
                }

                _mode = FormMode.View;
                _modeloIdEmEdicao = null;

                SetCamposEditaveis(false);
                dgvModelos.Enabled = true;
                AtualizarUIEstado();

                await BuscarModelosAsync(txtFiltroModelo.Text);
                SelecionarModeloNoGrid(savedId);

                await CarregarDetalhesAsync(savedId);

                MessageBox.Show("Modelo salvo com sucesso!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao salvar modelo: " + ex.Message);
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
                using var ctx = new grupometalContext();

                bool temProjeto = await ctx.Projeto
                    .AsNoTracking()
                    .AnyAsync(p => p.ModeloId == modeloId);

                bool temImagem = await ctx.ProjetoImagem
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

                var entity = await ctx.Modelo_EF
                    .FirstOrDefaultAsync(m => m.ModeloId == modeloId);

                if (entity == null)
                {
                    MessageBox.Show("Modelo não encontrado (pode ter sido removido por outro usuário).");
                    await BuscarModelosAsync(txtFiltroModelo.Text);
                    return;
                }

                // ✅ remove no MySQL
                ctx.Modelo_EF.Remove(entity);
                await ctx.SaveChangesAsync();

                // ✅ tenta remover também no PostGre
                try
                {
                    using var ctxPg = new gmetalContext();

                    bool temProjetoPg = await ctxPg.Projeto.AsNoTracking().AnyAsync(p => p.ModeloId == modeloId);
                    bool temImagemPg = await ctxPg.ProjetoImagem.AsNoTracking().AnyAsync(i => i.ModeloId == modeloId);

                    if (temProjetoPg || temImagemPg)
                    {
                        MessageBox.Show(
                            "⚠ Modelo removido no MySQL, mas NÃO foi removido no PostGre porque existem registros relacionados.\n\n" +
                            $"Projetos (PostGre): {(temProjetoPg ? "SIM" : "NÃO")}\n" +
                            $"Imagens (PostGre): {(temImagemPg ? "SIM" : "NÃO")}\n",
                            "Remoção parcial",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                    }
                    else
                    {
                        var pg = await ctxPg.Modelo.FirstOrDefaultAsync(m => m.ModeloId == modeloId);
                        if (pg != null)
                        {
                            ctxPg.Modelo.Remove(pg);
                            await ctxPg.SaveChangesAsync();
                        }
                    }
                }
                catch (Exception exPg)
                {
                    MessageBox.Show(
                        "⚠ Modelo removido no MySQL, mas falhou ao remover no PostGre:\n\n" + exPg.GetBaseException().Message,
                        "Remoção parcial",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }

                await BuscarModelosAsync(txtFiltroModelo.Text);

                if (dgvModelos.Rows.Count == 0) LimparDetalhes();

                MessageBox.Show("Modelo removido com sucesso!");
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
                MessageBox.Show("Erro ao remover: " + ex.Message);
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

            // =========================
            // HEADER DO PROJETO: editável SOMENTE durante _edicaoProjetoAtiva
            // =========================
            bool editProjeto = _edicaoProjetoAtiva;

            // sempre leitura (vem do sistema)
            txtClienteHeader.ReadOnly = true;
            txtDescricaoHeader.ReadOnly = true;
            txtModeloHeader.ReadOnly = true;

            // defaults do sistema
            txtRevHeader.ReadOnly = true;
            txtDataHeader.ReadOnly = true;
            txtAprovHeader.ReadOnly = true;

            // usuário edita durante criação
            txtPBHeader.ReadOnly = !editProjeto;
            txtPLHeader.ReadOnly = !editProjeto;
            txtRendimentoHeader.ReadOnly = !editProjeto;

            // segunda linha do header (histórico)
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

            // regras:
            // - Novo Projeto só se NÃO tem projeto atual
            // - Nova Revisão só se TEM projeto atual
            btnNovoProjeto.Enabled = temCliente && temModeloSelecionado && isView && !_edicaoProjetoAtiva && !_temProjetoAtual;
            btnNovaRevisao.Enabled = temCliente && temModeloSelecionado && isView && !_edicaoProjetoAtiva && _temProjetoAtual;

            btnSalvarProjeto.Enabled = _edicaoProjetoAtiva;

            // trava grid enquanto criando projeto/revisão
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

        // =========================
        // NOVO PROJETO / NOVA REVISÃO (RTF)
        // =========================

        // Novo Projeto: só se NÃO existir projeto no modelo (Rev = 1) e RTF em branco
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

            // ✅ Header automático pelo MODELO
            var header = await GetHeaderPesosDoModeloAsync(row.ModeloId);

            _edicaoEhRevisao = false;

            // ✅ Novo Projeto mantém FILEDIALOG e arquivo (como era)
            await IniciarEdicaoProjetoAsync(row, rev: 1, rtfBase: "", pb: header.pbText, pl: header.plText, rend: header.rendText);
        }

        // Nova Revisão: agora pega IMAGEM DO CLIPBOARD e RTF COPIADO da última revisão
        private async void btnNovaRevisao_Click(object sender, EventArgs e)
        {
            if (_empresaIdSelecionada == null) return;
            if (_mode != FormMode.View) return;

            if (dgvModelos.CurrentRow?.DataBoundItem is not ModeloGridRow row)
            {
                MessageBox.Show("Selecione um modelo.");
                return;
            }

            // pega dados da última revisão (inclui RTF)
            var info = await GetUltimoProjetoInfoAsync(row.ModeloId);
            if (!info.ok)
            {
                MessageBox.Show("Este modelo ainda não tem projeto. Use 'Novo Projeto' primeiro.");
                return;
            }

            int novaRev = info.lastRev + 1;

            // ✅ captura imagem do Clipboard
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

            // ✅ guarda bytes + nome (clipboard)
            _imagemBytes = ImageToPngBytes(clipImg);
            _imagemExt = ".png";
            _imagemNome = $"clipboard_rev_{novaRev}.png";

            // ✅ não usar arquivo nesta revisão
            _imagemLocalPath = null;

            // mostra imagem no preview
            TrocarImagem(new Bitmap(clipImg));

            // ✅ Header automático pelo MODELO (Real > Estimado)
            var header = await GetHeaderPesosDoModeloAsync(row.ModeloId);

            _edicaoEhRevisao = true;

            await IniciarEdicaoProjetoAsync_SemFileDialog(
                row,
                rev: novaRev,
                rtfBase: info.lastRtf,     // ✅ COPIA RTF ANTERIOR
                pb: header.pbText,
                pl: header.plText,
                rend: header.rendText
            );
        }

        // Método original (Novo Projeto): abre file picker
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

            // ✅ Novo Projeto é por arquivo: zera clipboard para não misturar
            _imagemBytes = null;
            _imagemNome = null;

            // selecionar imagem
            using var ofd = new OpenFileDialog();
            ofd.Title = _edicaoEhRevisao ? "Selecione a imagem da NOVA REVISÃO" : "Selecione a imagem do NOVO PROJETO";
            ofd.Filter = "Imagens|*.jpg;*.jpeg;*.png;*.bmp;*.webp";

            if (ofd.ShowDialog(this) != DialogResult.OK)
                return;

            _imagemLocalPath = ofd.FileName;
            _imagemExt = Path.GetExtension(_imagemLocalPath);

            // mostra imagem
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

            // entra em modo edição de projeto/revisão
            _edicaoProjetoAtiva = true;

            // aplica readOnly/enable conforme flag
            SetCamposEditaveis(false);

            // cabeçalho (defaults)
            txtModeloHeader.Text = row.Numeronocliente ?? "";
            txtDescricaoHeader.Text = (txtDescricao.Text ?? "").Trim();

            txtRevHeader.Text = _revProjetoAtual.ToString();
            txtDataHeader.Text = DateTime.Today.ToString("dd/MM/yyyy");
            txtAprovHeader.Text = "SMR";

            // defaults editáveis (agora vêm AUTOMÁTICOS do modelo)
            txtPBHeader.Text = pb ?? "";
            txtPLHeader.Text = pl ?? "";
            txtRendimentoHeader.Text = rend ?? "";

            // ✅ RTF: novo projeto -> em branco; revisão -> copia anterior
            SetRichTextSafe(txtRevisao, rtfBase ?? "");
            txtRevisao.ReadOnly = false;
            txtRevisao.Focus();

            AtualizarUIEstado();
            await Task.CompletedTask;
        }

        // Método novo (Nova Revisão): sem file dialog, usa imagem já carregada do Clipboard
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

            // entra em modo edição de projeto/revisão
            _edicaoProjetoAtiva = true;

            // aplica readOnly/enable conforme flag
            SetCamposEditaveis(false);

            // cabeçalho (defaults)
            txtModeloHeader.Text = row.Numeronocliente ?? "";
            txtDescricaoHeader.Text = (txtDescricao.Text ?? "").Trim();

            txtRevHeader.Text = _revProjetoAtual.ToString();
            txtDataHeader.Text = DateTime.Today.ToString("dd/MM/yyyy");
            txtAprovHeader.Text = "SMR";

            // defaults editáveis
            txtPBHeader.Text = pb ?? "";
            txtPLHeader.Text = pl ?? "";
            txtRendimentoHeader.Text = rend ?? "";

            // ✅ RTF: revisão -> copia anterior
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

        // ✅ Salva sempre como NOVA LINHA em Projeto e NOVA LINHA em ProjetoImagem
        // ✅ Replica no PostGre (gmetalContext) em projeto, projeto_imagem
        //     - PostGre.projeto_imagem.imagem = byte[]
        //     - MySQL.projeto_imagem.imagem = NULL
        //     - Não envia imagemurl para o PostGre
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
                using var ctx = new grupometalContext();

                // pega razão social e sigla (NOT NULL no projeto)
                var clienteInfo = await ctx.Empresa_GM
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

                // ✅ cria nova linha em Projeto (MySQL)
                var projeto = new Projeto
                {
                    ModeloId = _modeloIdProjeto.Value,
                    NroRevisao = _revProjetoAtual,

                    // ===== NOVOS CAMPOS =====
                    ColaboradorAprovadorId = "1",
                    Descricao = (txtDescricaoHeader.Text ?? "").Trim(),
                    Elaborador = "projeto",
                    Numeronocliente = (txtModeloHeader.Text ?? "").Trim(),
                    EmpresaId = _empresaIdSelecionada.Value,
                    Responsavel = "SMR",

                    // ===== JÁ EXISTIA (AGORA GARANTINDO LIMPO) =====
                    PesoBruto = StripSuffixER(txtPBHeader.Text),
                    PesoLiquido = StripSuffixER(txtPLHeader.Text),
                    Rendimento = StripSuffixER(txtRendimentoHeader.Text),

                    DataCriacao = DateTime.Today,
                    ColaboradorAprovadorSigla = "SMR",

                    Razaosocial = razao,
                    Sigla = sigla,

                    DadosTecnicos = txtRevisao?.Rtf ?? "" // ✅ RTF
                };

                ctx.Projeto.Add(projeto);
                await ctx.SaveChangesAsync();

                _projetoIdAtual = projeto.ProjetoId;

                // ✅ replica PROJETO para PostGre
                try
                {
                    await ReplicarProjetoParaPostgreAsync(projeto);
                }
                catch (Exception exPg)
                {
                    MessageBox.Show("⚠ Projeto salvo no MySQL, mas falhou ao replicar PROJETO no PostGre:\n\n" + exPg.GetBaseException().Message);
                }

                // ✅ upload imagem para azure (nome usa projetoId + modeloId + rev)
                var ext = string.IsNullOrWhiteSpace(_imagemExt) ? ".png" : _imagemExt!;
                var blobName = $"{_projetoIdAtual}_{_modeloIdProjeto}_{_revProjetoAtual}{ext}";

                string imageUrl;
                if (temClipboard)
                {
                    imageUrl = await UploadImagemProjetoAsync(blobName, _imagemBytes!);
                }
                else
                {
                    imageUrl = await UploadImagemProjetoAsync(blobName, _imagemLocalPath!);
                }

                // zera imagem padrão anterior
                var antigas = await ctx.ProjetoImagem
                    .Where(x => x.ModeloId == _modeloIdProjeto.Value && x.ImagemPadrao == 1)
                    .ToListAsync();

                foreach (var img in antigas)
                    img.ImagemPadrao = 0;

                var nomeImg = !string.IsNullOrWhiteSpace(_imagemNome)
                    ? _imagemNome!
                    : (temArquivo ? Path.GetFileName(_imagemLocalPath!) : "clipboard.png");

                // ✅ sempre cria nova linha em ProjetoImagem (MySQL)
                var pi = new ProjetoImagem
                {
                    ModeloId = _modeloIdProjeto.Value,
                    NroRevisao = _revProjetoAtual,
                    ImagemUrl = imageUrl,
                    ImagemNome = nomeImg,
                    ImagemPadrao = 1,

                    // ✅ requisito: no MySQL fica NULL
                    Imagem = null
                };

                ctx.ProjetoImagem.Add(pi);
                await ctx.SaveChangesAsync();

                // ✅ replica PROJETO_IMAGEM para PostGre (byte[]), sem imagemurl
                try
                {
                    byte[] bytesPg = temClipboard
                        ? _imagemBytes!
                        : await File.ReadAllBytesAsync(_imagemLocalPath!);

                    await ReplicarProjetoImagemParaPostgreAsync(
                        modeloId: _modeloIdProjeto.Value,
                        nroRevisao: _revProjetoAtual,
                        imagemNome: nomeImg,
                        imagemPadrao: 1,
                        imagemBytes: bytesPg
                    );
                }
                catch (Exception exPg)
                {
                    MessageBox.Show("⚠ Imagem salva no MySQL, mas falhou ao replicar PROJETO_IMAGEM no PostGre:\n\n" + exPg.GetBaseException().Message);
                }

                // sai do modo criação
                _edicaoProjetoAtiva = false;
                _edicaoEhRevisao = false;
                txtRevisao.ReadOnly = true;

                // limpa estado de imagem da sessão
                _imagemLocalPath = null;
                _imagemBytes = null;
                _imagemNome = null;

                MessageBox.Show("Projeto/Revisão salvo com sucesso!");

                // recarrega detalhes do modelo atual
                await CarregarDetalhesAsync(_modeloIdProjeto.Value);
            }
            catch (DbUpdateException dbEx)
            {
                var real = dbEx.GetBaseException().Message;
                MessageBox.Show("Erro ao salvar projeto:\n\n" + real);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao salvar projeto: " + ex.Message);
            }
            finally
            {
                AtualizarUIEstado();
            }
        }

        // Upload por ARQUIVO (Novo Projeto - como era)
        private async Task<string> UploadImagemProjetoAsync(string blobName, string localPath)
        {
            string containerUrlWithSas = $"{_baseUrl}{_containerName}?{_sasToken}";
            var containerClient = new BlobContainerClient(new Uri(containerUrlWithSas));
            var blobClient = containerClient.GetBlobClient(blobName);

            using var fs = new FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            // ✅ Deixa false se você quer evitar sobrescrita (projetoId não deve repetir).
            // Se quiser robustez contra duplo clique, mude para true.
            await blobClient.UploadAsync(fs, overwrite: false);

            return $"{_baseUrl}{_containerName}/{blobName}";
        }

        // Upload por BYTES (Nova Revisão - Clipboard)
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

        private async Task<(bool ok, int lastRev, string lastRtf, string pb, string pl, string rend)> GetUltimoProjetoInfoAsync(int modeloId)
        {
            using var ctx = new grupometalContext();

            var ult = await ctx.Projeto
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
        // ✅ HEADER AUTO (PB/PL/Rendimento) baseado no MODELO
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
            using var ctx = new grupometalContext();

            var w = await ctx.Modelo_EF
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

        // Remove sufixo " E" / " R" do fim (e deixa "%" intacto se existir)
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
        // TIPOS DE MODELO (COMBO)
        // =========================
        private class TipoModeloComboItem
        {
            public int TipodemodeloId { get; set; }
            public string Nome { get; set; } = "";
        }

        private async Task CarregarTiposModeloAsync()
        {
            using var ctx = new grupometalContext();

            var tiposRaw = await ctx.TipoModelo
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

        // ==========================================================
        // ✅ REPLICAÇÃO MySQL -> PostGre (gmetalContext)
        // ==========================================================

        private async Task ReplicarModeloParaPostgreAsync(Modelo_EF mysqlEntity)
        {
            using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

            var pg = await ctxPg.Modelo.FirstOrDefaultAsync(x => x.ModeloId == mysqlEntity.ModeloId);

            if (pg == null)
            {
                pg = new global::Controle_Pedidos.Entities_GM.Modelo_EF
                {
                    ModeloId = mysqlEntity.ModeloId
                };
                ctxPg.Modelo.Add(pg);
            }

            pg.ClienteempresaId = mysqlEntity.ClienteempresaId;
            pg.Numeronocliente = mysqlEntity.Numeronocliente;
            pg.TipodemodeloId = mysqlEntity.TipodemodeloId;
            pg.Descricao = mysqlEntity.Descricao;
            pg.Situacao = mysqlEntity.Situacao;
            pg.Observacao = mysqlEntity.Observacao;

            pg.Valordomodelo = mysqlEntity.Valordomodelo;
            pg.Nrodepartes = mysqlEntity.Nrodepartes;
            pg.Numerodefiguras = mysqlEntity.Numerodefiguras;
            pg.Numerodecaixasdemachoporfigura = mysqlEntity.Numerodecaixasdemachoporfigura;
            pg.Estado = mysqlEntity.Estado;

            pg.Pesomedio = mysqlEntity.Pesomedio;
            pg.Pesobruto = mysqlEntity.Pesobruto;
            pg.Pesoprevisto = mysqlEntity.Pesoprevisto;
            pg.Pesoliquido = mysqlEntity.Pesoliquido;

            await ctxPg.SaveChangesAsync();
        }

        private async Task ReplicarProjetoParaPostgreAsync(Projeto mysqlProjeto)
        {
            using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

            var pg = new global::Controle_Pedidos.Entities_GM.Projeto
            {
                ModeloId = mysqlProjeto.ModeloId,
                NroRevisao = mysqlProjeto.NroRevisao,

                ColaboradorAprovadorId = mysqlProjeto.ColaboradorAprovadorId,
                Descricao = mysqlProjeto.Descricao,
                Elaborador = mysqlProjeto.Elaborador,
                Numeronocliente = mysqlProjeto.Numeronocliente,
                EmpresaId = mysqlProjeto.EmpresaId,
                Responsavel = mysqlProjeto.Responsavel,

                PesoBruto = mysqlProjeto.PesoBruto,
                PesoLiquido = mysqlProjeto.PesoLiquido,
                Rendimento = mysqlProjeto.Rendimento,

                DataCriacao = EnsureUtc(mysqlProjeto.DataCriacao),
                ColaboradorAprovadorSigla = mysqlProjeto.ColaboradorAprovadorSigla,

                Razaosocial = mysqlProjeto.Razaosocial,
                Sigla = mysqlProjeto.Sigla,

                DadosTecnicos = mysqlProjeto.DadosTecnicos
            };

            ctxPg.Projeto.Add(pg);
            await ctxPg.SaveChangesAsync();
        }

        // ✅ agora usa bytes (serve tanto arquivo quanto clipboard)
        private async Task ReplicarProjetoImagemParaPostgreAsync(
            int modeloId,
            int nroRevisao,
            string imagemNome,
            int imagemPadrao,
            byte[] imagemBytes)
        {
            using var ctxPg = new global::Controle_Pedidos.Entities_GM.gmetalContext();

            var pg = new global::Controle_Pedidos.Entities_GM.ProjetoImagem
            {
                ModeloId = modeloId,
                NroRevisao = nroRevisao,
                ImagemNome = imagemNome,
                ImagemPadrao = imagemPadrao,
                Imagem = imagemBytes
            };

            ctxPg.ProjetoImagem.Add(pg);
            await ctxPg.SaveChangesAsync();
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

        private void btnImprimir_Click(object? sender, EventArgs e)
        {
            try
            {
                // 1) Diagnóstico rápido de nulls
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

                // 2) Imagem temporária (pode falhar se não tiver imagem)
                var imgPath = CriarImagemTempFileFromPictureBox(picModelo);

                // 3) Parâmetros (todos seguros)
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

                // 4) Abre relatório
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
    }
}
