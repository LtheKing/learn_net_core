using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TesExportImport.Model
{
    public class KertasKerjaModel
    {

        public long nomor { get; set; }
        public string indikator { get; set; }
        public double bobot { get; set; }
        public int level { get; set; }
        public string parameter { get; set; }
        public string aktivitas { get; set; }
        public string faktor_verifikatif { get; set; }
        public int checklist { get; set; }
        public double bobot_fv { get; set; }

    }

    public class KerangkaKerjaModel
    {
        public long nomor { get; set; }
        public string indikator { get; set; }
        public double bobot { get; set; }
        public int level { get; set; }
        public string parameter { get; set; }
        public string aktivitas { get; set; }
        public string faktor_verifikatif { get; set; }
        public int checklist { get; set; }
        public double bobot_fv { get; set; }
    }

    public class SDMModel
    {
        public long nomor { get; set; }
        public string indikator { get; set; }
        public double bobot { get; set; }
        public int level { get; set; }
        public string parameter { get; set; }
        public string aktivitas { get; set; }
        public string faktor_verifikatif { get; set; }
        public int checklist { get; set; }
        public double bobot_fv { get; set; }
    }
    public class ProsesModel
    {
        public long nomor { get; set; }
        public string indikator { get; set; }
        public double bobot { get; set; }
        public int level { get; set; }
        public string parameter { get; set; }
        public string aktivitas { get; set; }
        public string faktor_verifikatif { get; set; }
        public int checklist { get; set; }
        public double bobot_fv { get; set; }
    }
}

//- no
//- indikator
//- bobot(%)
//- level
//- parameter
//- aktivitas
//- faktor verifikatif
//- Checklist
//- bobot per FV(%)