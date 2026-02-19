using Controle_Pedidos.Controle_Producao.Entities;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Controle_Pedidos_8.Tela_Cadastro
{
    public partial class FrmSelecionarCliente : Form
    {
        private readonly string _filtroInicial;

        private List<EmpresaRow> _todos = new();
        private List<EmpresaRow> _visivel = new();

        // ✅ retornos
        public int EmpresaIdSelecionado { get; private set; }
        public string EmpresaNomeSelecionado { get; private set; } = "";
        public string CnpjSelecionado { get; private set; } = "";
        public string SiglaSelecionada { get; private set; } = "";

        public FrmSelecionarCliente(string filtroInicial)
        {
            InitializeComponent();
            _filtroInicial = filtroInicial ?? "";

            this.AcceptButton = btnSelecionar;
            this.CancelButton = btnCancelar;
        }

        private async void FrmSelecionarCliente_Load(object sender, EventArgs e)
        {
            try
            {
                ConfigurarGrid();

                await CarregarDadosAsync();

                // filtro inicial vindo do form principal
                txtFiltro.Text = _filtroInicial.Trim();
                AplicarFiltro(txtFiltro.Text);

                txtFiltro.Focus();
                txtFiltro.SelectionStart = txtFiltro.Text.Length;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar clientes: " + ex.Message);
            }
        }

        // =========================================================
        // 1) CONFIG GRID (pra não depender do Designer)
        // =========================================================
        private void ConfigurarGrid()
        {
            dgvEmpresas.AutoGenerateColumns = false;
            dgvEmpresas.Columns.Clear();

            dgvEmpresas.ReadOnly = true;
            dgvEmpresas.MultiSelect = false;
            dgvEmpresas.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvEmpresas.RowHeadersVisible = false;

            dgvEmpresas.AllowUserToAddRows = false;
            dgvEmpresas.AllowUserToDeleteRows = false;
            dgvEmpresas.AllowUserToResizeRows = false;

            dgvEmpresas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "EmpresaId",
                DataPropertyName = "EmpresaId",
                Name = "colEmpresaId",
                FillWeight = 15
            });

            dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Cnpj",
                DataPropertyName = "Cnpj",
                Name = "colCnpj",
                FillWeight = 25
            });

            dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Sigla",
                DataPropertyName = "Sigla",
                Name = "colSigla",
                FillWeight = 20
            });

            dgvEmpresas.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Empresa",
                DataPropertyName = "Empresa",
                Name = "colEmpresa",
                FillWeight = 40
            });
        }

        // =========================================================
        // 2) CARREGAR DO EF CORE (MySQL)
        // =========================================================
        private async Task CarregarDadosAsync()
        {
            // 🔴 AJUSTE AQUI:
            // Nome do seu DbContext
            using var ctx = new grupometalContext(); // <-- TROQUE se necessário

            // 🔴 AJUSTE AQUI:
            // ctx.Empresa (DbSet) + propriedades
            _todos = await ctx.Empresa_GM // <-- TROQUE para seu DbSet (ex: ctx.Cliente)
                .AsNoTracking()
                .Select(e => new EmpresaRow
                {
                    EmpresaId = e.EmpresaId,  // <-- TROQUE se seu campo chama diferente
                    Cnpj = e.Cnpj ?? "",
                    Sigla = e.Sigla ?? "",
                    Empresa = e.Nome ?? ""
                })
                .OrderBy(x => x.Empresa)
                .ToListAsync();

            _visivel = _todos.ToList();

            dgvEmpresas.DataSource = null;
            dgvEmpresas.DataSource = _visivel;
        }

        // =========================================================
        // 3) FILTRO
        // =========================================================
        private void txtFiltro_TextChanged(object sender, EventArgs e)
        {
            AplicarFiltro(txtFiltro.Text);
        }

        private void AplicarFiltro(string texto)
        {
            texto = (texto ?? "").Trim();

            if (string.IsNullOrWhiteSpace(texto))
            {
                _visivel = _todos.ToList();
            }
            else
            {
                string t = texto.ToLowerInvariant();

                _visivel = _todos
                    .Where(x =>
                           (x.Empresa ?? "").ToLowerInvariant().Contains(t)
                        || (x.Sigla ?? "").ToLowerInvariant().Contains(t)
                        || (x.Cnpj ?? "").ToLowerInvariant().Contains(t)
                        || x.EmpresaId.ToString().Contains(t))
                    .ToList();
            }

            dgvEmpresas.DataSource = null;
            dgvEmpresas.DataSource = _visivel;
        }

        private void btnLimpar_Click(object sender, EventArgs e)
        {
            txtFiltro.Text = "";
        }

        // =========================================================
        // 4) SELEÇÃO
        // =========================================================
        private void btnSelecionar_Click(object sender, EventArgs e)
        {
            ConfirmarSelecao();
        }

        private void dgvEmpresas_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            ConfirmarSelecao();
        }

        private void ConfirmarSelecao()
        {
            if (dgvEmpresas.CurrentRow?.DataBoundItem is not EmpresaRow row)
                return;

            EmpresaIdSelecionado = row.EmpresaId;
            EmpresaNomeSelecionado = row.Empresa;
            CnpjSelecionado = row.Cnpj;
            SiglaSelecionada = row.Sigla;

            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        // =========================================================
        // DTO pro Grid
        // =========================================================
        private class EmpresaRow
        {
            public int EmpresaId { get; set; }
            public string Cnpj { get; set; } = "";
            public string Sigla { get; set; } = "";
            public string Empresa { get; set; } = "";
        }
    }
}
