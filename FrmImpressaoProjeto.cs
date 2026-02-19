using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Controle_Pedidos_8.Tela_Cadastro
{
    public partial class FrmImpressaoProjeto : Form
    {
        public FrmImpressaoProjeto(Dictionary<string, string> parametros)
        {
            InitializeComponent();

            reportViewer1.Reset();
            reportViewer1.ProcessingMode = ProcessingMode.Local;
            reportViewer1.LocalReport.EnableExternalImages = true;
            reportViewer1.LocalReport.DataSources.Clear();

            // 🔍 Descobre o nome REAL embutido do RDLC
            var names = typeof(FrmImpressaoProjeto).Assembly.GetManifestResourceNames();
            var rdlcName = names.FirstOrDefault(n => n.EndsWith("RelatorioProjeto.rdlc", StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(rdlcName))
            {
                MessageBox.Show("Não achei RelatorioProjeto.rdlc embutido.\n\n" + string.Join("\n", names));
                throw new Exception("RDLC não embutido. Verifique Build Action = Embedded Resource.");
            }

            reportViewer1.LocalReport.ReportEmbeddedResource = rdlcName;

            // ✅ Só manda parâmetros que existem no RDLC
            var nomesExistentes = reportViewer1.LocalReport.GetParameters()
                .Select(p => p.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var lista = new List<ReportParameter>();
            foreach (var p in parametros)
            {
                if (nomesExistentes.Contains(p.Key))
                    lista.Add(new ReportParameter(p.Key, p.Value ?? ""));
            }

            reportViewer1.LocalReport.SetParameters(lista);

            this.Load += (s, e) => reportViewer1.RefreshReport();
        }
        private void FrmImpressaoProjeto_Load(object sender, EventArgs e)
        {
            reportViewer1.RefreshReport();
        }
    }
}
