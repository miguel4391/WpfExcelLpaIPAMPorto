using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Data.SqlClient;
using System.Drawing;

namespace WpfExcelLpaIPAMPorto
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }
        #region GlobalVariables
        public struct Competencias
        {
            public int adaptacao;
            public int trabalhar;
            public int decisoes;
            public int objactivos;
            public int ideias;
            public int aprendizagem;
            public int mentalidadeGlobal;
            public int gestaoEquipas;
        }

        public struct CursoFinalizado
        {
            public int codCurso;
            public string curso;
            public string data;
        }

        public class Notas
        {
            public string UC { get; set; }
            public string Nota { get; set; }
            public string ECTS { get; set; }
        }

        public struct Linguas
        {
            public string lingua;
            public string nivel;
        }

        public struct ResponsabilidadesAcademicas
        {
            public string responsabilidade;
            public string anoLetivo;
        }

        public struct Estagios
        {
            public string tipoEstagio;
            public string empresa;
            public string dataInicio;
            public string dataFim;
        }

        public struct Mobilidade
        {
            public string programa;
            public string tipo;
            public string anoLetivo;
        }

        public struct ResposabilidadeSocial
        {
            public string AcaoSocial;
            public string AnoLetivo;
        }

        public struct Premios
        {
            public string premio;
            public string anoLetivo;
        }

        public struct ActDesportivas
        {
            public string atividade;
            public string anoLetivo;
        }

        public struct OutrasExperiencias
        {
            public string experiencia;
            public string anoLetivo;
        }

        object replaceNome;
        object replaceNrAluno;
        object replaceCurso;
        object replaceData;
        object replaceCdCurso;
        object replaceMedia;

        ArrayList listaCursosFinalizados = new ArrayList();
        List<Notas> listaNotas = new List<Notas>();
        ArrayList listaLinguas = new ArrayList();
        ArrayList listaRespAcademicas = new ArrayList();
        ArrayList listaEstagios = new ArrayList();
        ArrayList listaMobilidade = new ArrayList();
        ArrayList listaRespSocial = new ArrayList();
        ArrayList listaPremios = new ArrayList();
        ArrayList listaAtividades = new ArrayList();
        ArrayList listaExperiencias = new ArrayList();
        #endregion

        #region GetCursos
        protected void GetCursos(string nrAluno)
        {
            Dictionary<string, string> listaCursos = new Dictionary<string, string>();
            string str = ConfigurationManager.ConnectionStrings["connectionStringSophia"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);

            string queryCursos = "use bdsophis; SELECT [RAlc_Ce_CdCurso],[Curs_NmCurso],[RAlc_DtFim],[RAlc_Estado] FROM[dbo].[TRAluCur] INNER JOIN TCURSOS ON[RAlc_Ce_CdCurso] = [Curs_Cp_CdCurso] where[RAlc_Cp_NAluno] = " + nrAluno;

            SqlCommand cmdCursos = new SqlCommand(queryCursos, conn);

            conn.Open();
            SqlDataReader dr = cmdCursos.ExecuteReader();
            CursoComboBox.Items.Clear();
            try
            {
                while (dr.Read())
                {
                    if (Convert.ToString(dr.GetValue(3)) == "F")
                    {
                        CursoFinalizado novoCurso = new CursoFinalizado();
                        novoCurso.codCurso = Convert.ToInt32(dr.GetValue(0));
                        novoCurso.curso = Convert.ToString(dr.GetValue(1));
                        string data = Convert.ToString(dr.GetValue(2));
                        int l = data.IndexOf(" ");
                        novoCurso.data = data.Substring(0, l);

                        CursoComboBox.Items.Add(novoCurso.codCurso + " : " + novoCurso.curso);
                        listaCursosFinalizados.Add(novoCurso);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            conn.Close();
        }
        #endregion

        #region GetDadosAlunos
        private void GetDadosAluno(string nrAluno)
        {
            Dictionary<string, string> listaCursos = new Dictionary<string, string>();
            pt.ipam.elpusonlineporto1.WebSapi request = new pt.ipam.elpusonlineporto1.WebSapi();
            string query = request.Execute("GetAluDadosPessoais", 1, "1", "2", "TpUtil=0;CdUtil=2029;PwdUtil=S1st3m0nl1ne#;CdAluno=" + nrAluno, "CdAluno;NmAluno");
            query = query.Replace("<sapiOutput><resultado><EstRes>0</EstRes><c1><![CDATA[", " ");
            query = query.Replace("]]></c1><c2><![CDATA[", "#");
            query = query.Replace("]]></c2></resultado><resultado><EstRes>0</EstRes><c1><![CDATA[", "#");
            query = query.Replace("]]></c2></resultado></sapiOutput>", "#");

            string[] campos = query.Split('#');

            replaceNome = (string)campos[1];
            replaceNrAluno = (string)campos[0];
        }
        #endregion

        #region GetNotas
        private void GetNotas()
        {
            string nrAluno = (string)replaceNrAluno;
            pt.ipam.elpusonlineporto1.WebSapi request = new pt.ipam.elpusonlineporto1.WebSapi();

            String query = request.Execute("GetAluNotas", 1, "1", "2", "TpUtil=0;CdUtil=2029;PwdUtil=S1st3m0nl1ne#;CdAluno=" + nrAluno + ";CdCurso=" + replaceCdCurso + ";NotaFinal=S", "NmDisc;Nota;ECTS;Classificacao");
            query = query.Replace("<sapiOutput><resultado><EstRes>0</EstRes><c1><![CDATA[", " ");
            query = query.Replace("]]></c1><c2><![CDATA[", "#");
            query = query.Replace("]]></c2><c3><![CDATA[", "#");
            query = query.Replace("]]></c3><c4><![CDATA[", "#");
            query = query.Replace("]]></c4></resultado><resultado><EstRes>0</EstRes><c1><![CDATA[", "#");
            query = query.Replace("]]></c4></resultado></sapiOutput>", "#");

            string[] campos = query.Split('#');

            listaNotas.Clear();
            for (int i = 0; i < campos.Length - 1; i += 4)
            {
                if (campos[i + 3] == "Aprovado" || campos[i + 3] == "Creditação" || campos[i + 3] == "Equivalência")
                {
                    Notas novaNota = new Notas();
                    novaNota.UC = campos[i];
                    try
                    {
                        int nota = Convert.ToInt32(campos[i + 1]);
                        if (nota >= 1)
                            novaNota.Nota = campos[i + 1];
                        else
                            novaNota.Nota = "Aprov.";
                    }
                    catch (Exception ex)
                    {
                        novaNota.Nota = "Aprov.";
                    }
                    novaNota.ECTS = campos[i + 2];
                    listaNotas.Add(novaNota);
                }
            }
        }
        #endregion

        #region GetMedia
        private void GetMedia(string nrAluno, string cdCurso)
        {
            string str = ConfigurationManager.ConnectionStrings["connectionStringSophia"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);

            string queryMedia = "use bdsophis; SELECT Ralc_NotaFim FROM TRAluCur WHERE RAlc_Cp_NAluno = " + nrAluno + "AND RAlc_Ce_CdCurso = " + cdCurso;

            SqlCommand cmdMedia = new SqlCommand(queryMedia, conn);
            conn.Open();
            SqlDataReader dr = cmdMedia.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    replaceMedia = Convert.ToString(dr.GetValue(0));
                }
            }
            catch (Exception)
            {

                throw;
            }
            conn.Close();
        }
        #endregion


        private void EscolherCursoButton_Click(object sender, RoutedEventArgs e)
        {
            string nrAluno = AlunoTextBox.Text;
            string codCurso = CursoComboBox.Text;
            string nomeCurso = CursoComboBox.Text;
            int l = codCurso.IndexOf(":");
            if (l > 0)
            {
                codCurso = codCurso.Substring(0, l);
                nomeCurso = nomeCurso.Substring(l + 1);
            }

            replaceCurso = nomeCurso;

            foreach (CursoFinalizado curso in listaCursosFinalizados)
            {
                if (curso.codCurso == Convert.ToInt32(codCurso))
                {
                    replaceData = curso.data;
                    replaceCdCurso = Convert.ToString(curso.codCurso);
                }
            }
            CreateDocument();
        }

        #region getDadosCompetencias
        private Competencias GetDadosCompetencias(string nrAluno)
        {
            Competencias competencias = new Competencias();
            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(str);

            string query = "Select [adaptacao],[trabalhar],[decisoes],[objactivos],[ideias],[aprendizagem],[mentalidadeGlobal],[gestaoEquipas] FROM [LPA].[dbo].[competenciasV2] where [Aluno] = " + nrAluno;
            SqlCommand cmd = new SqlCommand(query, conn);

            conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                try
                {
                    competencias.adaptacao = Convert.ToInt32(dr.GetValue(0));
                    competencias.trabalhar = Convert.ToInt32(dr.GetValue(1));
                    competencias.decisoes = Convert.ToInt32(dr.GetValue(2));
                    competencias.objactivos = Convert.ToInt32(dr.GetValue(3));
                    competencias.ideias = Convert.ToInt32(dr.GetValue(4));
                    competencias.aprendizagem = Convert.ToInt32(dr.GetValue(5));
                    competencias.mentalidadeGlobal = Convert.ToInt32(dr.GetValue(6));
                    competencias.gestaoEquipas = Convert.ToInt32(dr.GetValue(7));
                }
                catch (SqlException ex) { throw ex; }
            }
            conn.Close();

            return competencias;
        }
        #endregion

        #region GetLinguas
        private void GetLinguas(string nrAluno)
        {

            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);
            string query = "SELECT [Lingua],[Nivel] FROM [LPA].[dbo].[Linguas] WHERE Linguas.nrAluno = " + nrAluno;
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                try
                {
                    Linguas novaLingua = new Linguas();
                    novaLingua.lingua = Convert.ToString(dr.GetValue(0));
                    novaLingua.nivel = Convert.ToString(dr.GetValue(1));
                    listaLinguas.Add(novaLingua);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            conn.Close();
        }
        #endregion

        #region GetResponsabildiadesAcademicas
        private void GetRespAcademicas(string nrAluno)
        {
            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);
            string query = "SELECT [Responsabilidade],[AnoLetivo] FROM [LPA].[dbo].[ResponsAcademicas] where Aluno= " + nrAluno;
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                try
                {
                    ResponsabilidadesAcademicas novaResponsabilidade = new ResponsabilidadesAcademicas();
                    novaResponsabilidade.responsabilidade = Convert.ToString(dr.GetValue(0));
                    novaResponsabilidade.anoLetivo = Convert.ToString(dr.GetValue(1));
                    listaRespAcademicas.Add(novaResponsabilidade);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            conn.Close();
        }
        #endregion

        #region GetEstagios
        private void GetEstagios(string nrAluno)
        {
            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);
            string query = "SELECT TipoEstagio.TipoEstagio, [Empresa], [DataInicio], [DataFim] FROM [LPA].[dbo].[Estagio] INNER JOIN TipoEstagio on Estagio.TipoEstagio = TipoEstagio.IdTipoEstagio where Aluno = " + nrAluno;
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                try
                {
                    Estagios novaEstagio = new Estagios();
                    novaEstagio.tipoEstagio = Convert.ToString(dr.GetValue(0));
                    novaEstagio.empresa = Convert.ToString(dr.GetValue(1));
                    novaEstagio.dataInicio = Convert.ToDateTime(dr.GetValue(2)).Date.ToString();
                    novaEstagio.dataFim = Convert.ToDateTime(dr.GetValue(3)).Date.ToString();
                    listaEstagios.Add(novaEstagio);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            conn.Close();
        }
        #endregion

        #region GetMobilidade
        private void GetMobilidade(string nrAluno)
        {
            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);
            string query = "SELECT [Universidade],TipoMobInt.TipoMobInt,[AnoLetivo] FROM [LPA].[dbo].[MobInt] INNER JOIN TipoMobInt ON MobInt.TipoMobInt = TipoMobInt.IdTipoMobInt WHERE Aluno = " + nrAluno;
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                try
                {
                    Mobilidade novaMobilidade = new Mobilidade();
                    novaMobilidade.programa = Convert.ToString(dr.GetValue(0));
                    novaMobilidade.tipo = Convert.ToString(dr.GetValue(1));
                    novaMobilidade.anoLetivo = Convert.ToString(dr.GetValue(2));
                    listaMobilidade.Add(novaMobilidade);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            conn.Close();
        }
        #endregion

        #region GetResponsabildiadeSocial
        private void GetRespSocial(string nrAluno)
        {
            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);
            string query = "SELECT [AcaoSocial],[AnoLetivo] FROM [LPA].[dbo].[RespSocialVolutar] where Aluno = " + nrAluno;
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                try
                {
                    ResposabilidadeSocial novaResponsabilidade = new ResposabilidadeSocial();
                    novaResponsabilidade.AcaoSocial = Convert.ToString(dr.GetValue(0));
                    novaResponsabilidade.AnoLetivo = Convert.ToString(dr.GetValue(1));
                    listaRespSocial.Add(novaResponsabilidade);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            conn.Close();
        }
        #endregion

        #region getPremios
        private void GetPremios(string nrAluno)
        {
            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);
            string query = "SELECT [Premio],[AnoLetivo] FROM [LPA].[dbo].[PremiosReconhec] WHERE aluno = " + nrAluno;
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                try
                {
                    Premios novoPremio = new Premios();
                    novoPremio.premio = Convert.ToString(dr.GetValue(0));
                    novoPremio.anoLetivo = Convert.ToString(dr.GetValue(1));
                    listaPremios.Add(novoPremio);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            conn.Close();
        }
        #endregion

        #region GetAtividades
        private void GetAtividades(string nrAluno)
        {
            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);
            string query = "SELECT [Atividade],[AnoLetivo],[Aluno] FROM [LPA].[dbo].[AtividadesDesportivas] where Aluno = " + nrAluno;
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                try
                {
                    ActDesportivas novaAtividade = new ActDesportivas();
                    novaAtividade.atividade = Convert.ToString(dr.GetValue(0));
                    novaAtividade.anoLetivo = Convert.ToString(dr.GetValue(1));
                    listaAtividades.Add(novaAtividade);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            conn.Close();
        }
        #endregion

        #region GetOutrasexpiriencias
        private void GetOutrasExperiencias(string nrAluno)
        {
            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);
            string query = "SELECT [Experiencias],[AnoLetivo] FROM [LPA].[dbo].[OutrasExperiencias] where Aluno = " + nrAluno;
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                try
                {
                    OutrasExperiencias novaExpiriencia = new OutrasExperiencias();
                    novaExpiriencia.experiencia = Convert.ToString(dr.GetValue(0));
                    novaExpiriencia.anoLetivo = Convert.ToString(dr.GetValue(1));
                    listaExperiencias.Add(novaExpiriencia);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            conn.Close();
        }
        #endregion

        #region CreateDocument
        private void CreateDocument()
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            app.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
            object misValue = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(@"C:\LPA_IPAMPorto_V2\LPA_IPAM.xlsx",
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing);

            Worksheet ws = wb.Worksheets[1];

            int col = 1;
            int row = 16;

            ws.Range["B8"].Value = Convert.ToString(replaceNome).TrimStart();
            ws.Range["B9"].Value = Convert.ToString(replaceNrAluno).TrimStart();
            ws.Range["B10"].Value = Convert.ToString(replaceCurso).TrimStart();
            ws.Range["B11"].Value = Convert.ToString(replaceData).TrimStart();

            GetNotas();


            foreach (Notas nota in listaNotas)
            {
                ws.Cells[row, col].Font.Size = 10;
                ws.Cells[row, col].Value = nota.UC.TrimStart();
                ws.Cells[row, col + 5].Font.Size = 10;
                ws.Cells[row, col + 5].Value = nota.Nota;
                ws.Cells[row, col + 6].Font.Size = 10;
                ws.Cells[row, col + 6].Value = nota.ECTS;
                row++;
            }

            GetMedia((string)replaceNrAluno, (string)replaceCdCurso);
            ws.Range["B59"].Value = (string)replaceMedia + " valores";

            ws.Range["B73"].Value = Convert.ToString(replaceNome).TrimStart();
            ws.Range["B74"].Value = Convert.ToString(replaceNrAluno).TrimStart();
            ws.Range["B75"].Value = Convert.ToString(replaceCurso).TrimStart();
            ws.Range["B76"].Value = Convert.ToString(replaceData).TrimStart();

            Competencias competencias = GetDadosCompetencias((string)replaceNrAluno);



            Microsoft.Office.Interop.Excel.ChartObjects xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects(Type.Missing);
            Microsoft.Office.Interop.Excel.ChartObject myChart = (Microsoft.Office.Interop.Excel.ChartObject)xlCharts.Add(5, 1165, 470, 300);
            Microsoft.Office.Interop.Excel.Chart chartPage = myChart.Chart;

            Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection = (Microsoft.Office.Interop.Excel.SeriesCollection)chartPage.SeriesCollection();
            var ser = seriesCollection.NewSeries();

            chartPage.Legend.Delete();

            ser.Values = new double[] { competencias.adaptacao, competencias.trabalhar, competencias.decisoes, competencias.objactivos, competencias.ideias, competencias.aprendizagem, competencias.mentalidadeGlobal, competencias.gestaoEquipas };
            ser.XValues = new string[] { "Adaptação", "Trabalhar com os Outros", "Tomar Decisões", "Alcançar Objetivos", "Geração de Ideias", "Aprendizagem", "Mentalidade Global", "Gestão de Equipas" };


            //chartRange = ws.get_Range("A1", "d5");
            //chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlBarClustered;

            chartPage.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 212, 83, 10);
            chartPage.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 212, 83, 10);
            chartPage.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 212, 83, 10);
            chartPage.SeriesCollection(1).Points(4).Format.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 212, 83, 10);
            chartPage.SeriesCollection(1).Points(5).Format.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 212, 83, 10);
            chartPage.SeriesCollection(1).Points(6).Format.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 212, 83, 10);
            chartPage.SeriesCollection(1).Points(7).Format.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 212, 83, 10);
            chartPage.SeriesCollection(1).Points(8).Format.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 212, 83, 10);

            ws.Range["B127"].Value = Convert.ToString(replaceNome).TrimStart();
            ws.Range["B128"].Value = Convert.ToString(replaceNrAluno).TrimStart();
            ws.Range["B129"].Value = Convert.ToString(replaceCurso).TrimStart();
            ws.Range["B130"].Value = Convert.ToString(replaceData).TrimStart();

            GetLinguas((string)replaceNrAluno);
            GetRespAcademicas((string)replaceNrAluno);
            GetEstagios((string)replaceNrAluno);
            GetMobilidade((string)replaceNrAluno);
            GetPremios((string)replaceNrAluno);
            GetAtividades((string)replaceNrAluno);
            GetOutrasExperiencias((string)replaceNrAluno);
            GetRespSocial((string)replaceNrAluno);

            col = 1;
            row = 135;

            if (listaLinguas.Count > 0)
            {
                string header = "LÍNGUAS";

                ws.Cells[row, col].Font.Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);
                ws.Cells[row, col].Font.Bold = true;
                ws.Cells[row, col].Value = header;

                Microsoft.Office.Interop.Excel.Range cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                Microsoft.Office.Interop.Excel.Borders border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 4d;
                border[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);

                foreach (Linguas l in listaLinguas)
                {
                    row++;
                    ws.Cells[row, col].Value = " > " + l.lingua + " | Nível " + l.nivel;
                }

                cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                row += 2;
            }

            if (listaRespAcademicas.Count > 0)
            {
                string header = "RESPONSABILIDADES ACADÉMICAS";

                ws.Cells[row, col].Font.Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);
                ws.Cells[row, col].Font.Bold = true;
                ws.Cells[row, col].Value = header;

                Microsoft.Office.Interop.Excel.Range cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                Microsoft.Office.Interop.Excel.Borders border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 4d;
                border[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);

                foreach (ResponsabilidadesAcademicas l in listaRespAcademicas)
                {
                    row++;
                    ws.Cells[row, col].Value = " > " + l.responsabilidade + " | " + l.anoLetivo;
                }

                cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                row += 2;
            }

            if (listaEstagios.Count > 0)
            {
                string header = "ESTÁGIOS";

                ws.Cells[row, col].Font.Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);
                ws.Cells[row, col].Font.Bold = true;
                ws.Cells[row, col].Value = header;

                Microsoft.Office.Interop.Excel.Range cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                Microsoft.Office.Interop.Excel.Borders border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 4d;
                border[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);

                foreach (Estagios l in listaEstagios)
                {
                    if (l.dataInicio != "1900-01-01")
                    {
                        row++;
                        ws.Cells[row, col].Value = " > " + l.tipoEstagio + " na " + l.empresa + " | De " + l.dataInicio.Substring(0, 10) + " a " + l.dataFim.Substring(0, 10);
                    }
                    else
                    {
                        row++;
                        ws.Cells[row, col].Value = " > " + l.tipoEstagio + " na " + l.empresa;
                    }
                }

                cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                row += 2;
            }

            if (listaMobilidade.Count > 0)
            {
                string header = "MOBILIDADE INTERNACIONAL";

                ws.Cells[row, col].Font.Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);
                ws.Cells[row, col].Font.Bold = true;
                ws.Cells[row, col].Value = header;

                Microsoft.Office.Interop.Excel.Range cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                Microsoft.Office.Interop.Excel.Borders border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 4d;
                border[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);

                foreach (Mobilidade l in listaMobilidade)
                {
                    row++;
                    ws.Cells[row, col].Value = " > " + l.tipo + " na " + l.programa + " | " + l.anoLetivo;
                }

                cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                row += 2;
            }

            if (listaRespSocial.Count > 0)
            {
                string header = "RESPONSABILIDADE SOCIAL & VOLUNTARIADO";

                ws.Cells[row, col].Font.Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);
                ws.Cells[row, col].Font.Bold = true;
                ws.Cells[row, col].Value = header;

                Microsoft.Office.Interop.Excel.Range cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                Microsoft.Office.Interop.Excel.Borders border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 4d;
                border[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);

                foreach (ResposabilidadeSocial l in listaRespSocial)
                {
                    row++;
                    ws.Cells[row, col].Value = " > " + l.AcaoSocial + " | " + l.AnoLetivo;
                }

                cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                row += 2;
            }

            if (listaPremios.Count > 0)
            {
                string header = "PRÉMIOS & RECONHECIMENTOS";

                ws.Cells[row, col].Font.Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);
                ws.Cells[row, col].Font.Bold = true;
                ws.Cells[row, col].Value = header;

                Microsoft.Office.Interop.Excel.Range cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                Microsoft.Office.Interop.Excel.Borders border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 4d;
                border[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);

                foreach (Premios l in listaPremios)
                {
                    row++;
                    ws.Cells[row, col].Value = " > " + l.premio + " | " + l.anoLetivo;
                }

                cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                row += 2;
            }

            if (listaAtividades.Count > 0)
            {
                string header = "ATIVIDADES DESPORTIVAS";

                ws.Cells[row, col].Font.Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);
                ws.Cells[row, col].Font.Bold = true;
                ws.Cells[row, col].Value = header;

                Microsoft.Office.Interop.Excel.Range cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                Microsoft.Office.Interop.Excel.Borders border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 4d;
                border[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);

                foreach (ActDesportivas l in listaAtividades)
                {
                    row++;
                    ws.Cells[row, col].Value = " > " + l.atividade + " | " + l.anoLetivo;
                }

                cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                row += 2;
            }

            if (listaExperiencias.Count > 0)
            {
                string header = "OUTRAS EXPERIÊNCIAS";

                ws.Cells[row, col].Font.Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);
                ws.Cells[row, col].Font.Bold = true;
                ws.Cells[row, col].Value = header;

                Microsoft.Office.Interop.Excel.Range cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                Microsoft.Office.Interop.Excel.Borders border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 4d;
                border[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.Color.FromArgb(0, 212, 83, 10);

                foreach (OutrasExperiencias l in listaExperiencias)
                {
                    row++;
                    ws.Cells[row, col].Value = " > " + l.experiencia + " | " + l.anoLetivo;
                }

                cells = ws.Range[ws.Cells[row, col], ws.Cells[row, col + 6]];
                border = cells.Borders;
                border[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border[XlBordersIndex.xlEdgeBottom].Weight = 2d;

                row += 2;
            }

            string data = Convert.ToString(DateTime.Now.ToLongDateString());
            ws.Cells[164, 1].Value = "Porto, " + data;

            int nrCertificado = GravaCertificado(competencias);
            string anoLetivo = getAnoLetivo((string)replaceNrAluno, (string)replaceCdCurso);

            ws.Range["E11"].Value = nrCertificado + " | POR | " + anoLetivo;
            ws.Range["E76"].Value = nrCertificado + " | POR | " + anoLetivo;
            ws.Range["E130"].Value = nrCertificado + " | POR | " + anoLetivo;
        }
        #endregion

        #region Grava Certificado
        private int GravaCertificado(Competencias competencias)
        {
            int nrCertificado = 0;
            string str = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            SqlConnection conn = new SqlConnection(str);

            string queryCertificado = "SELECT id_certificados FROM CertificadosV2";
            SqlCommand cmdCertificado = new SqlCommand(queryCertificado, conn);
            conn.Open();
            SqlDataReader dr = cmdCertificado.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    int temp = Convert.ToInt32(dr.GetValue(0));
                    if (temp > nrCertificado) nrCertificado = temp;
                }
                nrCertificado++;
            }
            catch { nrCertificado = 431; }
            conn.Close();

            string queryCompetencias = "INSERT INTO [dbo].[CertificadosV2] ([nr_estudante],[data],[adaptacao],[trabalhar],[decisoes],[objactivos],[ideias],[aprendizagem],[mentalidadeGlobal],[gestaoEquipas]) VALUES (" + replaceNrAluno + ",'" + Convert.ToString(DateTime.Now) + "'," + competencias.adaptacao + "," + competencias.trabalhar + "," + competencias.decisoes + "," + competencias.objactivos + "," + competencias.ideias + "," + competencias.aprendizagem + "," + competencias.mentalidadeGlobal + "," + competencias.gestaoEquipas + ")";
            SqlCommand cmdCompetencias = new SqlCommand(queryCompetencias, conn);
            conn.Open();
            cmdCompetencias.ExecuteNonQuery();
            conn.Close();

            foreach (Notas nota in listaNotas)
            {
                string query = "insert into [dbo].[Notas] ([uc],[nota],[ects],[nrEstudante],[codcurso],[data],[certificado]) Values ('" + nota.UC + "'," + nota.Nota + "," + nota.ECTS + "," + replaceNrAluno + "," + replaceCdCurso + ",'" + Convert.ToString(DateTime.Now) + "'," + nrCertificado + ")";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            foreach (Linguas lingua in listaLinguas)
            {
                string query = "INSERT INTO [dbo].[Experiencias] ([tipo],[descritivo],[nrAluno],[data],[certificado]) VALUES ('" + lingua.lingua + "','" + lingua.nivel + "'," + replaceNrAluno + ",'" + Convert.ToString(DateTime.Now) + "'," + nrCertificado + ")";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            foreach (ResponsabilidadesAcademicas resp in listaRespAcademicas)
            {
                string query = "INSERT INTO [dbo].[Experiencias] ([tipo],[descritivo],[nrAluno],[data], [certificado]) VALUES ('" + resp.responsabilidade + "','" + resp.anoLetivo + "'," + replaceNrAluno + ",'" + Convert.ToString(DateTime.Now) + "'," + nrCertificado + ")";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            foreach (Estagios estagio in listaEstagios)
            {
                string query = "INSERT INTO [dbo].[Experiencias] ([tipo],[descritivo],[nrAluno],[data],[certificado]) VALUES ('" + estagio.tipoEstagio + "','" + estagio.empresa + " inicio " + estagio.dataInicio + " fim " + estagio.dataFim + "'," + replaceNrAluno + ",'" + Convert.ToString(DateTime.Now) + "'," + nrCertificado + ")";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            foreach (Mobilidade mobilidade in listaMobilidade)
            {
                string query = "INSERT INTO [dbo].[Experiencias] ([tipo],[descritivo],[nrAluno],[data],[certificado]) VALUES ('" + mobilidade.tipo + "','" + mobilidade.programa + " anoletivo " + mobilidade.anoLetivo + "'," + replaceNrAluno + ",'" + Convert.ToString(DateTime.Now) + "'," + nrCertificado + ")";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            foreach (ResposabilidadeSocial responsabilidade in listaRespSocial)
            {
                string query = "INSERT INTO [dbo].[Experiencias] ([tipo],[descritivo],[nrAluno],[data],[certificado]) VALUES ('" + responsabilidade.AcaoSocial + "','" + responsabilidade.AnoLetivo + "'," + replaceNrAluno + ",'" + Convert.ToString(DateTime.Now) + "'," + nrCertificado + ")";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            foreach (Premios premio in listaPremios)
            {
                string query = "INSERT INTO [dbo].[Experiencias] ([tipo],[descritivo],[nrAluno],[data],[certificado]) VALUES ('" + premio.premio + "','" + premio.anoLetivo + "'," + replaceNrAluno + ",'" + Convert.ToString(DateTime.Now) + "'," + nrCertificado + ")";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            foreach (ActDesportivas atividade in listaAtividades)
            {
                string query = "INSERT INTO [dbo].[Experiencias] ([tipo],[descritivo],[nrAluno],[data],[certificado]) VALUES ('" + atividade.atividade + "','" + atividade.anoLetivo + "'," + replaceNrAluno + ",'" + Convert.ToString(DateTime.Now) + "'," + nrCertificado + ")";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            foreach (OutrasExperiencias experiencias in listaExperiencias)
            {
                string query = "INSERT INTO [dbo].[Experiencias] ([tipo],[descritivo],[nrAluno],[data],[certificado]) VALUES ('" + experiencias.experiencia + "','" + experiencias.anoLetivo + "'," + replaceNrAluno + ",'" + Convert.ToString(DateTime.Now) + "'," + nrCertificado + ")";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }

            listaNotas.Clear();
            listaLinguas.Clear();
            listaRespAcademicas.Clear();
            listaEstagios.Clear();
            listaMobilidade.Clear();
            listaRespSocial.Clear();
            listaPremios.Clear();
            listaAtividades.Clear();
            listaExperiencias.Clear();

            return nrCertificado;
        }
        #endregion

        #region Get Ano Letivo
        private string getAnoLetivo(string nrAluno, string codCurso)
        {
            string anoLetivo = "";
            string anoLetivo2 = "";

            string str = ConfigurationManager.ConnectionStrings["connectionStringSophia"].ConnectionString;

            string query = "use bdSophis; SELECT Ralc_UltAnoLect FROM TRAluCur WHERE Ralc_Cp_Naluno = " + nrAluno + " AND Ralc_Ce_CdCurso = " + codCurso;
            SqlConnection conn = new SqlConnection(str);
            SqlCommand cmd = new SqlCommand(query, conn);

            conn.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                anoLetivo = Convert.ToString(reader.GetValue(0));
            }
            conn.Close();

            anoLetivo2 = Convert.ToString(Convert.ToInt32(anoLetivo) + 1);
            return anoLetivo + "/" + anoLetivo2;
        }
        #endregion
        private void EscolherButton_Click(object sender, RoutedEventArgs e)
        {
            string nrAluno = AlunoTextBox.Text;
            GetCursos(nrAluno);
            CursoComboBox.IsEnabled = true;
            GetDadosAluno(nrAluno);
        }
    }
}
