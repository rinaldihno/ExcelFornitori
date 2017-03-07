using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelFornitori
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<Fornitore> fornitori = ReadFornitori();
            bool b = ValidaFornitori(fornitori, "BETA");
            ScriviFornitoriFinal(fornitori, "BETA");
        }

        private void ScriviFornitoriV1(List<Fornitore> fornitori, string wave)
        {
            //Scrive i fornitori su un file Excel per il popolamento del modello bravosolution

            IWorkbook workbook = null;  //IWorkbook determina si es xls o xlsx              
            ISheet worksheet = null;

            workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Test");

            int rowNum = 0;

            foreach (var fornitore in fornitori.Where(q => q.Wave == wave))
            {
                IRow row = sheet.CreateRow(rowNum++);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(fornitore.CodiceSap);

                cell = row.CreateCell(1);
                cell.SetCellValue(fornitore.StatoAusiliario);

                cell = row.CreateCell(2);
                cell.SetCellValue(fornitore.Nazione);

                cell = row.CreateCell(3);
                cell.SetCellValue(fornitore.RagioneSociale);

                cell = row.CreateCell(4);
                cell.SetCellValue(fornitore.FormaGiuridica);

                cell = row.CreateCell(5);
                cell.SetCellValue(fornitore.PartitaIva);

                cell = row.CreateCell(6);
                cell.SetCellValue(fornitore.CodiceFiscale);

                //Telefono centralino
                cell = row.CreateCell(7);
                cell.SetCellValue(".");

                //Indirizzo email azienda
                cell = row.CreateCell(8);
                cell.SetCellValue(".");

                //Indirizzo
                cell = row.CreateCell(9);
                cell.SetCellValue(fornitore.Indirizzo);

                //Cap
                cell = row.CreateCell(10);
                cell.SetCellValue(fornitore.Cap);

                //Comune
                cell = row.CreateCell(11);
                cell.SetCellValue(fornitore.Comune);

                //Provincia
                cell = row.CreateCell(12);
                cell.SetCellValue(fornitore.Provincia);

                //Colonne a blank non obbligatorie
                cell = row.CreateCell(13);
                cell.SetCellValue(".");

                cell = row.CreateCell(14);
                cell.SetCellValue(".");

                cell = row.CreateCell(15);
                cell.SetCellValue(".");

                cell = row.CreateCell(16);
                cell.SetCellValue(".");

                cell = row.CreateCell(17);
                cell.SetCellValue(".");

                cell = row.CreateCell(18);
                cell.SetCellValue(".");

                //Nome
                cell = row.CreateCell(19);
                cell.SetCellValue(fornitore.Nome);

                //Cognome
                cell = row.CreateCell(20);
                cell.SetCellValue(fornitore.Cognome);

                //Email
                cell = row.CreateCell(21);
                cell.SetCellValue(fornitore.Email);

                //Telefono
                cell = row.CreateCell(22);
                cell.SetCellValue(fornitore.Telefono);

                //Cellulare
                cell = row.CreateCell(23);
                cell.SetCellValue(fornitore.Telefono);

                //Username
                cell = row.CreateCell(24);
                cell.SetCellValue(fornitore.PartitaIva);

                //LinguaPredefinita
                cell = row.CreateCell(25);
                cell.SetCellValue(fornitore.LinguaPredefinita);

                //FusoOrario
                cell = row.CreateCell(26);
                cell.SetCellValue(fornitore.FusoOrario);

            }


            using (FileStream file = new FileStream(@"c:\temp\MassiveSellerInsert.xlsx", FileMode.OpenOrCreate, FileAccess.Write))
            {
                workbook.Write(file);
                file.Close();
            }
        }


       
        private void ScriviFornitori2(List<Fornitore> fornitori, string wave)
        {
            IWorkbook workbook = null;  //IWorkbook determina si es xls o xlsx              
            ISheet worksheet = null;


            using (FileStream file = new FileStream(@"c:\temp\TemplateMassivo.xlsx", FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
                file.Close();
            }


            using (FileStream file = new FileStream(@"c:\temp\TemplateMassivo.xlsx", FileMode.Create, FileAccess.Write))
            {

                //  workbook = WorkbookFactory.Create(file);
                worksheet = workbook.GetSheetAt(0);

                int rownum = 4;

                foreach (var fornitore in fornitori.Where(q => q.Wave == wave))
                {
                    IRow row = worksheet.CreateRow(rownum++);

                    int numcol = 2;


                    ICell cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.CodiceSap);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.StatoAusiliario);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Nazione);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.RagioneSociale);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.FormaGiuridica);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.PartitaIva);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.CodiceFiscale);

                    //Telefono centralino
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(".");

                    //Indirizzo email azienda
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(".");

                    //Indirizzo
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Indirizzo);

                    //Cap
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Cap);

                    //Comune
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Comune);

                    //Provincia
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Provincia);

                    //Colonne a blank non obbligatorie
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(".");

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(".");

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(".");

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(".");

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(".");

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(".");

                    //Nome
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Nome);

                    //Cognome
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Cognome);

                    //Email
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Email);

                    //Telefono
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Telefono);

                    //Cellulare
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Telefono);

                    //Username
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.PartitaIva);

                    //LinguaPredefinita
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.LinguaPredefinita);

                    //FusoOrario
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.FusoOrario);
                }



                workbook.Write(file);
                file.Close();
            }
        }


        private void ScriviFornitoriFinal(List<Fornitore> fornitori, string wave)
        {
            IWorkbook workbook = null;  //IWorkbook determina si es xls o xlsx              
            ISheet worksheet = null;


            using (FileStream file = new FileStream(@"c:\temp\MassiveSellerInsert.xlsx", FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
                file.Close();
            }


            using (FileStream file = new FileStream(@"c:\temp\MassiveSellerInsert.xlsx", FileMode.Create, FileAccess.Write))
            {

                //  workbook = WorkbookFactory.Create(file);
                worksheet = workbook.GetSheetAt(0);

                int rownum = 3;

                foreach (var fornitore in fornitori.Where(q => q.Wave == wave))
                {
                    IRow row = worksheet.CreateRow(rownum++);

                    int numcol = 1;
                    ICell cell = row.CreateCell(numcol++);

                    //User activation
                    cell.SetCellValue("N");

                    //User alias
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.RagioneSociale);


                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.RagioneSociale);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.FormaGiuridica);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.CodiceFiscale);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.PartitaIva);

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Nazione);

                    //Indirizzo
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Indirizzo);

                    //Cap
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Cap);

                    //Comune
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Comune);

                    //Provincia
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Provincia);

                    //Telefono centralino
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(".");

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue("");

                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.CodiceSap);

                    //Cognome
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Cognome);

                    //Nome
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Nome);


                    //Email
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Email);

                    //Telefono
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Telefono);

                    //Cellulare
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.Telefono);

                    //Username
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.PartitaIva);

                    //LinguaPredefinita
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.LinguaPredefinita);

                    //FusoOrario
                    cell = row.CreateCell(numcol++);
                    cell.SetCellValue(fornitore.FusoOrario);

                  

                  
                }



                workbook.Write(file);
                file.Close();
            }
        }




        private bool ValidaFornitori(List<Fornitore> fornitori, string wave)
        {

            //Verifica che tutti siano in possesso di PartitaIva o CodiceFiscale in quanto è lo username
            long checkPartitaIva = fornitori.Where(q => q.Wave == wave && string.IsNullOrEmpty(q.CodiceFiscale) && string.IsNullOrEmpty(q.PartitaIva)).LongCount();

            //Verifica che tutti siano in possesso di un indirizzo email
            long checkEmail = fornitori.Where(q => q.Wave == wave && string.IsNullOrEmpty(q.Email)).LongCount();

            return (checkPartitaIva == 0) && (checkEmail == 0);
        }

        private List<Fornitore> ReadFornitori()
        {
            IWorkbook workbook = null;  //IWorkbook determina si es xls o xlsx              
            ISheet worksheet = null;
            string first_sheet_name = "";
            List<Fornitore> lstFornitori = new List<Fornitore>();

            using (FileStream FS = new FileStream(@"c:\temp\FornitoriFinal.xlsx", FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(FS);
                worksheet = workbook.GetSheetAt(1);
                first_sheet_name = worksheet.SheetName;



                for (int rowIndex = 1; rowIndex <= worksheet.LastRowNum; rowIndex++)
                {
                    IRow row = worksheet.GetRow(rowIndex);
                    ICell cell = null;

                    if (row != null)
                    {
                        Fornitore f = new Fornitore();

                        f.StatoAusiliario = "ATTIVO";
                        f.FormaGiuridica = "NOT_DEF";
                        f.LinguaPredefinita = "it_IT";
                        f.FusoOrario = "Europe/Brussels";


                        cell = row.GetCell(0, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                            f.MatriceRischi = cell.StringCellValue;

                        cell = row.GetCell(1, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                            f.Wave = cell.StringCellValue;

                        cell = row.GetCell(9, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                            f.CodiceSap = cell.NumericCellValue;

                        cell = row.GetCell(10, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                        {
                            f.RagioneSociale = cell.StringCellValue;

                            //Elaborare la forma giuridica nella codifica Bravosolution
                            f.FormaGiuridica = GetFormaGiuridica(cell.StringCellValue);

                        }

                        cell = row.GetCell(11, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                        {
                            if (cell.CellType == CellType.Numeric)
                                f.CodiceFiscale = cell.NumericCellValue.ToString();
                            else
                                f.CodiceFiscale = cell.StringCellValue;

                        }

                        cell = row.GetCell(12, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                        {
                            if (cell.CellType == CellType.Numeric)
                                f.PartitaIva = cell.NumericCellValue.ToString();
                            else
                                f.PartitaIva = cell.StringCellValue;
                        }


                        cell = row.GetCell(13, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                            f.Indirizzo = cell.StringCellValue;


                        cell = row.GetCell(14, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                        {
                            if (cell.CellType == CellType.Numeric)
                                f.Cap = cell.NumericCellValue.ToString();
                            else
                                f.Cap = cell.StringCellValue;
                        }

                        cell = row.GetCell(15, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                            f.Comune = cell.StringCellValue;


                        cell = row.GetCell(16, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                        {
                            //TODO sistemare la provincia con la codifica bravosolution
                            f.Provincia = decodeProvincia(cell.StringCellValue);

                            
                        }

                        

                        cell = row.GetCell(17, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                            f.Nazione = cell.StringCellValue;

                        cell = row.GetCell(18, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                            f.Nome = cell.StringCellValue.ToUpper();

                        cell = row.GetCell(19, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                            f.Cognome = cell.StringCellValue.ToUpper();

                        cell = row.GetCell(20, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                            f.Email = cell.StringCellValue.ToLower();

                        cell = row.GetCell(21, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                        {
                            if (cell.CellType == CellType.Numeric)
                                f.Telefono = cell.NumericCellValue.ToString();
                            else
                                f.Telefono = cell.StringCellValue;


                            f.Telefono = f.Telefono.Replace(" ", "").Replace("/", "");

                            if (!f.Telefono.StartsWith("+")) {
                                if(f.Telefono.Length==12 && f.Telefono.StartsWith("39"))
                                    f.Telefono = "+" + f.Telefono;
                                else if (f.Telefono.Length <= 10)
                                    f.Telefono = "+39"+f.Telefono;
                            }
                        }

                        cell = row.GetCell(22, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        if (cell != null)
                        {
                            if (cell.CellType == CellType.Numeric)
                                f.Cellulare = cell.NumericCellValue.ToString();
                            else
                                f.Cellulare = cell.StringCellValue;


                            f.Cellulare = f.Cellulare.Replace(" ", "").Replace("/", "");

                            if (!f.Cellulare.StartsWith("+"))
                            {
                                if (f.Cellulare.Length == 12 && f.Cellulare.StartsWith("39"))
                                    f.Cellulare = "+" + f.Cellulare;
                                else if (f.Cellulare.Length <= 10)
                                    f.Cellulare = "+39" + f.Cellulare;
                            }
                        }

                        lstFornitori.Add(f);

                    }





                }
            }

            int i = lstFornitori.Where(q => q.Wave == "BETA").Count();

            return lstFornitori;
        }

        private string decodeProvincia(string provincia)
        {
            string provinciaDecodificata = "ALTRO";

            switch (provincia.ToUpper())
            {
                case "MI":
                    provinciaDecodificata = "Milano";
                    break;
                case "FI":
                    provinciaDecodificata = "Firenze";
                    break;
                case "CN":
                    provinciaDecodificata = "Cuneo";
                    break;
                case "PI":
                    provinciaDecodificata = "Pisa";
                    break;
                case "BO":
                    provinciaDecodificata = "Bologna";
                    break;
                case "TO":
                    provinciaDecodificata = "Torino";
                    break;
                case "PV":
                    provinciaDecodificata = "Pavia";
                    break;
                case "PA":
                    provinciaDecodificata = "Palermo";
                    break;
                case "BZ":
                    provinciaDecodificata = "Bolzano";
                    break;
                case "AR":
                    provinciaDecodificata = "Arezzo";
                    break;
                case "LI":
                    provinciaDecodificata = "Livorno";
                    break;
                case "RM":
                    provinciaDecodificata = "Roma";
                    break;
                case "PT":
                    provinciaDecodificata = "Potenza";
                    break;
                case "AN":
                    provinciaDecodificata = "Ancona";
                    break;
                case "RE":
                    provinciaDecodificata = "Reggio Emilia";
                    break;
                case "VR":
                    provinciaDecodificata = "Verona";
                    break;
                case "PG":
                    provinciaDecodificata = "Perugia";
                    break;
                case "LU":
                    provinciaDecodificata = "Lucca";
                    break;
                case "GR":
                    provinciaDecodificata = "Grosseto";
                    break;
                case "FC":
                    provinciaDecodificata = "Forli Cesena";
                    break;
                case "MS":
                    provinciaDecodificata = "Messina";
                    break;
                case "MO":
                    provinciaDecodificata = "Modena";
                    break;
                case "AQ":
                    provinciaDecodificata = "L'Aquila";
                    break;
                case "RN":
                    provinciaDecodificata = "Riccione";
                    break;
                case "VE":
                    provinciaDecodificata = "Venezia";
                    break;
                case "TR":
                    provinciaDecodificata = "Trapani";
                    break;
                case "BG":
                    provinciaDecodificata = "Bergamo";
                    break;
                case "LT":
                    provinciaDecodificata = "Latina";
                    break;
                case "PO":
                    provinciaDecodificata = "Prato";
                    break;
                case "BS":
                    provinciaDecodificata = "Brescia";
                    break;
                case "SI":
                    provinciaDecodificata = "Siena";
                    break;
                case "MC":
                    provinciaDecodificata = "Macerata";
                    break;
                case "MN":
                    provinciaDecodificata = "Mantova";
                    break;
                case "VI":
                    provinciaDecodificata = "Vicenza";
                    break;
                case "SP":
                    provinciaDecodificata = "La Spezia";
                    break;
                case "VT":
                    provinciaDecodificata = "Viterbo";
                    break;
                case "LE":
                    provinciaDecodificata = "Lecce";
                    break;
                case "SR":
                    provinciaDecodificata = "Siracusa";
                    break;
                case "TN":
                    provinciaDecodificata = "Trento";
                    break;
                case "FE":
                    provinciaDecodificata = "Ferrara";
                    break;
                case "TV":
                    provinciaDecodificata = "Treviso";
                    break;
                case "LO":
                    provinciaDecodificata = "Lodi";
                    break;
                case "GE":
                    provinciaDecodificata = "Genova";
                    break;
                case "SO":
                    provinciaDecodificata = "Sondrio";
                    break;
                case "SV":
                    provinciaDecodificata = "Savona";
                    break;
                case "MB":
                    provinciaDecodificata = "Monza Brianza";
                    break;
                case "PN":
                    provinciaDecodificata = "Pordenone";
                    break;
                case "RA":
                    provinciaDecodificata = "Ravenna";
                    break;
                case "PE":
                    provinciaDecodificata = "Pesaro";
                    break;
                case "CH":
                    provinciaDecodificata = "Chieti";
                    break;
                case "PR":
                    provinciaDecodificata = "Parma";
                    break;
                case "PD":
                    provinciaDecodificata = "Padova";
                    break;
                case "AL":
                    provinciaDecodificata = "Alessandria";
                    break;
                case "TE":
                    provinciaDecodificata = "Teramo";
                    break;
                case "PU":
                    provinciaDecodificata = "Pesaro";
                    break;



                default:
                    break;
            }

            return provinciaDecodificata;


        }

        private string GetFormaGiuridica(string ragioneSociale)
        {
            string temp = ragioneSociale.Replace(".", "").ToUpper();

            if (temp.Contains("SPA"))
                return "IT00001";

            if (temp.Contains("SRL"))
                return "IT00002";

            if (temp.Contains("SNC"))
                return "IT00004";

            if (temp.Contains("SOCCOOP") || (temp.Contains("SOC") && temp.Contains("COOP")))
                return "IT00008";

            if (temp.Contains("SCARL"))
                return "IT00014";

            if (temp.Contains("SOCCOOPPA") || (temp.Contains("SOC") && temp.Contains("COOP")))
                return "IT00015";

            if (temp.Contains("CONSORTILE"))
                return "IT00025";

            return "NOT_DEF";

        }

        private object GetCellValue(ICell cell)
        {
            object valorCell = null;

            switch (cell.CellType)
            {
                case CellType.Blank: valorCell = DBNull.Value; break;
                case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                case CellType.String: valorCell = cell.StringCellValue; break;
                case CellType.Numeric:
                    if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                    else { valorCell = cell.NumericCellValue; }
                    break;
                case CellType.Formula:
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.Blank: valorCell = DBNull.Value; break;
                        case CellType.String: valorCell = cell.StringCellValue; break;
                        case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                        case CellType.Numeric:
                            if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                            else { valorCell = cell.NumericCellValue; }
                            break;
                    }
                    break;
                default: valorCell = cell.StringCellValue; break;
            }

            return valorCell;
        }





        /// <summary>Abre un archivo de Excel (xls o xlsx) y lo convierte en un DataTable.
        /// LA PRIMERA FILA DEBE CONTENER LOS NOMBRES DE LOS CAMPOS.</summary>
        /// <param name="pRutaArchivo">Ruta completa del archivo a abrir.</param>
        /// <param name="pHojaIndex">Número (basado en cero) de la hoja que se desea abrir. 0 es la primera hoja.</param>
        private DataTable Excel_To_DataTable(string pRutaArchivo, int pHojaIndex)
        {
            // --------------------------------- //
            /* REFERENCIAS:
             * NPOI.dll
             * NPOI.OOXML.dll
             * NPOI.OpenXml4Net.dll */
            // --------------------------------- //
            /* USING:
             * using NPOI.SS.UserModel;
             * using NPOI.HSSF.UserModel;
             * using NPOI.XSSF.UserModel; */
            // AUTOR: Ing. Jhollman Chacon R. 2015
            // --------------------------------- //
            DataTable Tabla = null;
            try
            {
                if (System.IO.File.Exists(pRutaArchivo))
                {

                    IWorkbook workbook = null;  //IWorkbook determina si es xls o xlsx              
                    ISheet worksheet = null;
                    string first_sheet_name = "";

                    using (FileStream FS = new FileStream(pRutaArchivo, FileMode.Open, FileAccess.Read))
                    {
                        workbook = WorkbookFactory.Create(FS);          //Abre tanto XLS como XLSX
                        worksheet = workbook.GetSheetAt(pHojaIndex);    //Obtener Hoja por indice
                        first_sheet_name = worksheet.SheetName;         //Obtener el nombre de la Hoja

                        Tabla = new DataTable(first_sheet_name);
                        Tabla.Rows.Clear();
                        Tabla.Columns.Clear();

                        // Leer Fila por fila desde la primera
                        for (int rowIndex = 0; rowIndex <= worksheet.LastRowNum; rowIndex++)
                        {
                            DataRow NewReg = null;
                            IRow row = worksheet.GetRow(rowIndex);
                            IRow row2 = null;
                            IRow row3 = null;

                            if (rowIndex == 0)
                            {
                                row2 = worksheet.GetRow(rowIndex + 1); //Si es la Primera fila, obtengo tambien la segunda para saber el tipo de datos
                                row3 = worksheet.GetRow(rowIndex + 2); //Y la tercera tambien por las dudas
                            }

                            if (row != null) //null is when the row only contains empty cells 
                            {
                                if (rowIndex > 0) NewReg = Tabla.NewRow();

                                int colIndex = 0;
                                //Leer cada Columna de la fila
                                foreach (ICell cell in row.Cells)
                                {
                                    object valorCell = null;
                                    string cellType = "";
                                    string[] cellType2 = new string[2];

                                    if (rowIndex == 0) //Asumo que la primera fila contiene los titlos:
                                    {
                                        for (int i = 0; i < 2; i++)
                                        {
                                            ICell cell2 = null;
                                            if (i == 0) { cell2 = row2.GetCell(cell.ColumnIndex); }
                                            else { cell2 = row3.GetCell(cell.ColumnIndex); }

                                            if (cell2 != null)
                                            {
                                                switch (cell2.CellType)
                                                {
                                                    case CellType.Blank: break;
                                                    case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                    case CellType.String: cellType2[i] = "System.String"; break;
                                                    case CellType.Numeric:
                                                        if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                        else
                                                        {
                                                            cellType2[i] = "System.Double";  //valorCell = cell2.NumericCellValue;
                                                        }
                                                        break;

                                                    case CellType.Formula:
                                                        bool continuar = true;
                                                        switch (cell2.CachedFormulaResultType)
                                                        {
                                                            case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                            case CellType.String: cellType2[i] = "System.String"; break;
                                                            case CellType.Numeric:
                                                                if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                                else
                                                                {
                                                                    try
                                                                    {
                                                                        //DETERMINAR SI ES BOOLEANO
                                                                        if (cell2.CellFormula == "TRUE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                        if (continuar && cell2.CellFormula == "FALSE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                        if (continuar) { cellType2[i] = "System.Double"; continuar = false; }
                                                                    }
                                                                    catch { }
                                                                }
                                                                break;
                                                        }
                                                        break;
                                                    default:
                                                        cellType2[i] = "System.String"; break;
                                                }
                                            }
                                        }

                                        //Resolver las diferencias de Tipos
                                        if (cellType == null)
                                            cellType = ""; //Merk
                                        if (cellType2[0] == cellType2[1]) { cellType = cellType2[0]; }
                                        else
                                        {
                                            if (cellType2[0] == null) cellType = cellType2[1];
                                            if (cellType2[1] == null) cellType = cellType2[0];
                                            if (cellType == "") cellType = "System.String";
                                        }

                                        //Obtener el nombre de la Columna
                                        string colName = "Column_{0}";
                                        try { colName = cell.StringCellValue; }
                                        catch { colName = string.Format(colName, colIndex); }

                                        //Verificar que NO se repita el Nombre de la Columna
                                        foreach (DataColumn col in Tabla.Columns)
                                        {
                                            if (col.ColumnName == colName) colName = string.Format("{0}_{1}", colName, colIndex);
                                        }

                                        //Agregar el campos de la tabla:
                                        DataColumn codigo = new DataColumn(colName, System.Type.GetType(cellType));
                                        Tabla.Columns.Add(codigo); colIndex++;
                                    }
                                    else
                                    {
                                        //Las demas filas son registros:
                                        switch (cell.CellType)
                                        {
                                            case CellType.Blank: valorCell = DBNull.Value; break;
                                            case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                            case CellType.String: valorCell = cell.StringCellValue; break;
                                            case CellType.Numeric:
                                                if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                                else { valorCell = cell.NumericCellValue; }
                                                break;
                                            case CellType.Formula:
                                                switch (cell.CachedFormulaResultType)
                                                {
                                                    case CellType.Blank: valorCell = DBNull.Value; break;
                                                    case CellType.String: valorCell = cell.StringCellValue; break;
                                                    case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                                    case CellType.Numeric:
                                                        if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                                        else { valorCell = cell.NumericCellValue; }
                                                        break;
                                                }
                                                break;
                                            default: valorCell = cell.StringCellValue; break;
                                        }
                                        //Agregar el nuevo Registro
                                        if (cell.ColumnIndex <= Tabla.Columns.Count - 1) NewReg[cell.ColumnIndex] = valorCell;
                                    }
                                }
                            }
                            if (rowIndex > 0) Tabla.Rows.Add(NewReg);
                        }
                        Tabla.AcceptChanges();
                    }
                }
                else
                {
                    throw new Exception("ERROR 404: El archivo especificado NO existe.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Tabla;
        }


        /// <summary>Convierte un DataTable en un archivo de Excel (xls o Xlsx) y lo guarda en disco.</summary>
        /// <param name="pDatos">Datos de la Tabla a guardar. Usa el nombre de la tabla como nombre de la Hoja</param>
        /// <param name="pFilePath">Ruta del archivo donde se guarda.</param>
        private void DataTable_To_Excel(DataTable pDatos, string pFilePath)
        {
            try
            {
                if (pDatos != null && pDatos.Rows.Count > 0)
                {
                    IWorkbook workbook = null;
                    ISheet worksheet = null;

                    using (FileStream stream = new FileStream(pFilePath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        string Ext = System.IO.Path.GetExtension(pFilePath); //<-Extension del archivo
                        switch (Ext.ToLower())
                        {
                            case ".xls":
                                HSSFWorkbook workbookH = new HSSFWorkbook();
                                NPOI.HPSF.DocumentSummaryInformation dsi = NPOI.HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
                                dsi.Company = "Cutcsa"; dsi.Manager = "Departamento Informatico";
                                workbookH.DocumentSummaryInformation = dsi;
                                workbook = workbookH;
                                break;

                            case ".xlsx": workbook = new XSSFWorkbook(); break;
                        }

                        worksheet = workbook.CreateSheet(pDatos.TableName); //<-Usa el nombre de la tabla como nombre de la Hoja

                        //CREAR EN LA PRIMERA FILA LOS TITULOS DE LAS COLUMNAS
                        int iRow = 0;
                        if (pDatos.Columns.Count > 0)
                        {
                            int iCol = 0;
                            IRow fila = worksheet.CreateRow(iRow);
                            foreach (DataColumn columna in pDatos.Columns)
                            {
                                ICell cell = fila.CreateCell(iCol, CellType.String);
                                cell.SetCellValue(columna.ColumnName);
                                iCol++;
                            }
                            iRow++;
                        }

                        //FORMATOS PARA CIERTOS TIPOS DE DATOS
                        ICellStyle _doubleCellStyle = workbook.CreateCellStyle();
                        _doubleCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.###");

                        ICellStyle _intCellStyle = workbook.CreateCellStyle();
                        _intCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");

                        ICellStyle _boolCellStyle = workbook.CreateCellStyle();
                        _boolCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("BOOLEAN");

                        ICellStyle _dateCellStyle = workbook.CreateCellStyle();
                        _dateCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy");

                        ICellStyle _dateTimeCellStyle = workbook.CreateCellStyle();
                        _dateTimeCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy HH:mm:ss");

                        //AHORA CREAR UNA FILA POR CADA REGISTRO DE LA TABLA
                        foreach (DataRow row in pDatos.Rows)
                        {
                            IRow fila = worksheet.CreateRow(iRow);
                            int iCol = 0;
                            foreach (DataColumn column in pDatos.Columns)
                            {
                                ICell cell = null; //<-Representa la celda actual                               
                                object cellValue = row[iCol]; //<- El valor actual de la celda

                                switch (column.DataType.ToString())
                                {
                                    case "System.Boolean":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Boolean);

                                            if (Convert.ToBoolean(cellValue)) { cell.SetCellFormula("TRUE()"); }
                                            else { cell.SetCellFormula("FALSE()"); }

                                            cell.CellStyle = _boolCellStyle;
                                        }
                                        break;

                                    case "System.String":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.String);
                                            cell.SetCellValue(Convert.ToString(cellValue));
                                        }
                                        break;

                                    case "System.Int32":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt32(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Int64":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt64(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Decimal":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;
                                    case "System.Double":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;

                                    case "System.DateTime":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDateTime(cellValue));

                                            //Si No tiene valor de Hora, usar formato dd-MM-yyyy
                                            DateTime cDate = Convert.ToDateTime(cellValue);
                                            if (cDate != null && cDate.Hour > 0) { cell.CellStyle = _dateTimeCellStyle; }
                                            else { cell.CellStyle = _dateCellStyle; }
                                        }
                                        break;
                                    default:
                                        break;
                                }
                                iCol++;
                            }
                            iRow++;
                        }

                        workbook.Write(stream);
                        stream.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
