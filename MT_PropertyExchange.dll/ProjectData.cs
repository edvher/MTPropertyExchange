using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace CPi.iDimAT.Common
{
    public enum CalculationType //Berechnungsart
    {
        Kabel = 1, 
        Parallelkabel = 2, 
        Versorgungsleitung = 3
    }

    [Serializable] public class OfficeProperties : ICloneable
    {
        public Projekt()
        {
            projektBearbeiter = new Bearbeiter();
            
            eingabeSpannung = new ListValue(3, "Drehstrom", 400);
            eingabeLeiter = new ListValue(2, "Kupfer", 57);
            eingabeIsolierung = new AddValue("PVC-Isolierte Kabel", 1);
            eingabeKennlinie = new AddValue("LSS B", 1);

            eingabeStromkreis = new AddValue("Endstromkreis", 1);
            eingabeVerlegeart = new AddValue("in Gebäuden (in/auf Wand)", 1);
            eingabeVerlegung = new AddValue("A - Aderleitungen in Wärmedämmung", 1);
            eingabeVerlegeAnordnung = new AddValue("gebündelt", 1);
            eingabeAnordnung = new AddValue("in Rohr/Kanal in Wand", 2);
            eingabeParaKabel = new double[5] { 0, 0, 0, 0, 0 };
        }

        private double m_leistung = 0;
        private double m_bStrom = 0;
        public bool showMessage = false;
        public string message = String.Empty;

        #region Projektangaben
        public long projektId = -1;
        public string projektBezeichnung = String.Empty;
        public Bearbeiter projektBearbeiter;
        public long projektNormId = 2009;
        public string projektBerechnungsArt = CalculationType.Kabel.ToString();
        #endregion

        #region Einstellungen
        public ListValue eingabeSpannung;
        public ListValue eingabeLeiter;
        public AddValue eingabeIsolierung;
        public double eingabeSpannAbf = 1.5;
        public double eingabeSpannAbfVolt
        {
            get { return (double)eingabeSpannung.value / 100 * eingabeSpannAbf; }
            set { }
        }
        public AddValue eingabeKennlinie;
        #endregion

        #region Eingaben
        public double eingabeLeitungslaenge = 0;
        public double eingabeCosPhi = 1;
        public bool eingabeCheckKurz = false;
        public double eingabeSpannungKurz = 230;
        public int eingabeVorwiderstand = 200;
        public AddValue eingabeStromkreis;
        public bool eingabeCheckBrand = false;
        public string eingabeBrandKlasse = "E 30";
        public double eingabeBrandLaenge = 0;
        public double eingabeLeistung
        {
            get
            {
                if (eingabeNachLeistung)
                {
                    if (projektBerechnungsArt == CalculationType.Versorgungsleitung.ToString())
                        m_leistung = v_leistungGes;
                    return m_leistung;
                }
                else
                {
                    switch (eingabeSpannung.id)
                    {
                        case 1: //Wechselstrom
                            m_leistung = m_bStrom * eingabeSpannung.value * eingabeCosPhi;
                            break;
                        case 2: //Gleichstrom
                            m_leistung = m_bStrom * eingabeSpannung.value;
                            break;
                        case 3: //Drehstrom
                            m_leistung = m_bStrom * eingabeSpannung.value * eingabeCosPhi * Math.Sqrt(3);
                            break;
                        default:
                            m_leistung = -1;
                            break;
                    }
                    return m_leistung;
                }
            }
            set { m_leistung = value; }
        }
        public double eingabeBStrom
        {
            get
            {
                if (eingabeNachLeistung)
                {
                    if (projektBerechnungsArt == CalculationType.Versorgungsleitung.ToString())
                        m_leistung = v_leistungGes;
                    switch (eingabeSpannung.id)
                    {
                        case 1: //Wechselstrom
                            m_bStrom = m_leistung / (eingabeSpannung.value * eingabeCosPhi);
                            break;
                        case 2: //Gleichstrom
                            m_bStrom = m_leistung / eingabeSpannung.value;
                            break;
                        case 3: //Drehstrom
                            m_bStrom = m_leistung / (eingabeSpannung.value * eingabeCosPhi * Math.Sqrt(3));
                            break;
                        default:
                            m_bStrom = -1;
                            break;
                    }
                    if(projektBerechnungsArt != CalculationType.Versorgungsleitung.ToString())
                        m_bStrom *= eingabeGZF;
                    return m_bStrom;
                }
                else
                    return m_bStrom;
            }
            set { m_bStrom = value; }
        }
        public double eingabeGZF = 1;
        public bool eingabeNachLeistung = true;
        //Versorgungsleitung
        public int eingabeAnzVWE = 0;
        public double eingabeLeistungVWE = 0;
        public double eingabeGzfVWE = 0;

        public double eingabeLeistungAWE = 0;
        public double eingabeGzfAWE = 0;
        public double eingabeAnzAWE = 0;

        public double eingabeLeistungDLE = 0;
        public double eingabeGzfDLE = 0;
        public double eingabeAnzDLE = 0;

        public double eingabeLeistungNacht = 0;
        public double eingabeGzfNacht = 0;
        public double eingabeAnzNacht = 0;
        public double v_leistungGes
        {
            get 
            { 
                return eingabeAnzVWE * eingabeLeistungVWE * eingabeGzfVWE 
                    + eingabeLeistungAWE * eingabeGzfAWE
                    + eingabeLeistungDLE * eingabeGzfDLE 
                    + eingabeLeistungNacht * eingabeGzfNacht;
            }
            set { }
        }
        #endregion

        #region Verlegung
        public int eingabeAnzStromkreise = 1;
        public AddValue eingabeVerlegeart;
        public AddValue eingabeVerlegung;
        public AddValue eingabeVerlegeAnordnung;
        public AddValue eingabeAnordnung;
        public int eingabeAnzWannen = 0;
        public int eingabeUmgTemperatur = 25;
        public double[] eingabeParaKabel;
        #endregion

        #region Zwischenergebnisse/Werte
        public int v_anzAdern
        {
            get
            {
                switch (eingabeSpannung.id)
                {
                    case 1: //Wechselstrom
                        return 2;
                    case 2: //Gleichstrom
                        return 2;
                    case 3: //Drehstrom
                        return 3;
                    default:
                        return -1;
                }
            }
            set { }
        }
        public double v_faktorTemp = 1;
        public double v_faktorHaeufung = 1; //bei Parallelkabeln entspricht dieser dem neuen Häufungsfaktor
        public double v_bemessungsstrom = 0;
        public double v_impedanz = 0;
        //public double v_resistanz = 0;
        //public double v_reaktanz = 0;
        //Parallelkabel
        public double v_maxA = 0;
        public double v_maxAneu = 0;
        public int v_anzStromkreise_vorhanden
        {
            get
            {
                int count = 0;
                for (int i = 0; i < eingabeParaKabel.Length; i++)
                    if (eingabeParaKabel[i] > 0)
                        count++;
                return count + eingabeAnzStromkreise;
            }
        }
        public double v_faktorHaeufung_vorhanden = 1;
        public double v_bemessungsstromAmax = 0;
        public double v_bemessungsstromAmax_neu = 0;
        public double v_izvAmax //Izv(Amax)
        {
            get
            {
                return v_faktorHaeufung_vorhanden * v_faktorTemp * v_bemessungsstromAmax;
            }
        }
        public double v_iz_vorhanden // Izv
        {
            get
            {
                if (v_maxA != 0)
                    return v_izvAmax * (eingabeParaKabel[0] / v_maxA 
                                    + eingabeParaKabel[1] / v_maxA 
                                    + eingabeParaKabel[2] / v_maxA
                                    + eingabeParaKabel[3] / v_maxA 
                                    + eingabeParaKabel[4] / v_maxA);
                else
                    return 0;
            }
        }
        public double v_iznAmax
        {
            get
            {
                if (v_faktorHaeufung_vorhanden != 0)
                    return v_izvAmax * v_faktorHaeufung / v_faktorHaeufung_vorhanden;
                else
                    return 0;
            }
        }
        public double v_iznAmax_neu
        {
            get
            {
                return v_bemessungsstromAmax_neu * v_faktorTemp * v_faktorHaeufung;
            }
        }
        public double v_Izn
        {
            get
            {
                if (v_maxA != 0 && ergParaQuerBestimmt != 0 && showMessage != true)
                    return v_iznAmax * (eingabeParaKabel[0] + eingabeParaKabel[1] + eingabeParaKabel[2] + eingabeParaKabel[3] + eingabeParaKabel[4] + ergParaQuerBestimmt) / v_maxA;
                else
                    return 0;
            }
        }
        public double v_Izn_neu
        {
            get
            {
                if (v_maxAneu != 0 && ergParaQuerBestimmt != 0 && showMessage != true)
                    return v_iznAmax_neu * (eingabeParaKabel[0] + eingabeParaKabel[1] + eingabeParaKabel[2] + eingabeParaKabel[3] + eingabeParaKabel[4] + ergParaQuerBestimmt) / v_maxAneu;
                else
                    return 0;
            }
        }

        public double v_FaktorNutzungAWE = 0;
        public double v_FaktorNutzungVWE = 0;
        public double v_FaktorNutzungDLE = 0;
        public double v_FaktorNutzungNacht = 0;

        public double v_FaktorKurz = 0;
        public double v_VergleichsspannungEndstromkreis
        {
            get { return ergAbschaltStrom * ergZs / 1000; }
            set { }
        }
        public double v_VergleichsspannungVerteilnetz
        {
            get
            {
                if (eingabeSpannungKurz > 400)
                    return ergZs / 1000 * 2.5 * ergIn;
                else
                    return ergZs / 1000 * 1.6 * ergIn;
            }
            set { }
        }
        public double v_VergleichsspannungHauptversorgung
        {
            get { return ergZs / 1000 * 3.5 * ergIn; }
            set { }
        }
        public int v_VerBrand
        {
            get
            {
                if (eingabeLeitungslaenge != 0)
                    //return Math.Round((brandLaenge / leitungslaenge * 100), 0, MidpointRounding.AwayFromZero);
                    return (int)(eingabeBrandLaenge / eingabeLeitungslaenge * 100);
                else
                    return  0;
            }
            set { }
        }
        public double v_SpannAbfBrand
        {
            get { return (double)eingabeSpannung.value / 100 * 4.5; }
            set { }
        }
        public double v_leistungSpann
        {
            get
            {
                switch (eingabeSpannung.id)
                {
                    case 1: //Wechselstrom
                        return eingabeSpannung.value * ergIn;
                    case 2: //Gleichstrom
                        return eingabeSpannung.value * ergIn * eingabeCosPhi;
                    case 3: //Drehstrom
                        return eingabeSpannung.value * ergIn * eingabeCosPhi * Math.Sqrt(3);
                    default:
                        return 0;
                }
            }
            set { }
        }
        public double v_querSpannVorhPara
        {
            get
            {
                switch (eingabeSpannung.id)
                {
                    case 1: //Wechselstrom
                        return (eingabeBStrom * 2 * eingabeLeitungslaenge * eingabeCosPhi) / (eingabeLeiter.value * eingabeSpannAbfVolt);
                    case 2: //Gleichstrom
                        return (eingabeBStrom * 2 * eingabeLeitungslaenge) / (eingabeLeiter.value * eingabeSpannAbfVolt);
                    case 3: //Drehstrom
                        return (eingabeBStrom * Math.Sqrt(3) * eingabeLeitungslaenge * eingabeCosPhi) / (eingabeLeiter.value * eingabeSpannAbfVolt);
                    default:
                        return 0;
                }
            }
            set { }
        }
        #endregion

        #region Ergebnisse
            public double ergIz
            {
                get {
                        return v_bemessungsstrom * v_faktorTemp * v_faktorHaeufung; 
                }
                set { }
            }
            public double ergIzPara
        {
            get
            {
                if (v_Izn_neu > 0)
                    return v_Izn_neu;
                else
                    return v_Izn;
            }
            set { }
        }
            public int ergIn = 0;
            public double ergSpannAbfRel
            {
                get
                {
                    if (eingabeSpannung.value != 0)
                        return (double)ergSpannAbfVolt / eingabeSpannung.value * 100;
                    else
                        return 0;
                }
                set { }
            }
            public double ergSpannAbfVolt
            {
                get
                {
                    if (ergQuerNorm != 0 && eingabeSpannung.value != 0)
                    {
                        //return (leistung * leitungslaenge * (resistanz + reaktanz * Math.Acos(cosPhi))) / spannung.value;
                        if (eingabeSpannung.id == 3)
                            return (v_impedanz * v_leistungSpann * eingabeLeitungslaenge) / (1000 * eingabeSpannung.value);
                        else
                            return (v_impedanz * v_leistungSpann * eingabeLeitungslaenge * 2) / (1000 * eingabeSpannung.value);
                    }
                    return 0;
                    //if (ergQuerNorm != 0 && eingabeLeiter.value != 0)
                    //{
                    //    switch (eingabeSpannung.id)
                    //    {
                    //        case 1: //Wechselstrom
                    //            return (eingabeBStrom * 2 * eingabeLeitungslaenge * eingabeCosPhi) / (eingabeLeiter.value * ergQuerNorm);
                    //        case 2: //Gleichstrom
                    //            return (eingabeBStrom * 2 * eingabeLeitungslaenge) / (eingabeLeiter.value * ergQuerNorm);
                    //        case 3: //Drehstrom
                    //            return (eingabeBStrom * Math.Sqrt(3) * eingabeLeitungslaenge * eingabeCosPhi) / (eingabeLeiter.value * ergQuerNorm);
                    //        default:
                    //            return 0;
                    //    }
                    //}
                    //return 0;
                }
                set { }
            }
            public double ergZs
            {
                get {
                    if (!showMessage && eingabeCheckKurz)
                        return (2 * (eingabeVorwiderstand + v_impedanz));
                    else
                        return 0;
                }
                set { }
            }
            public double ergAbschaltStrom
            {
                get { return ergIn * v_FaktorKurz; }
                set { }
            }
            public string ergVerhaeltnis
            {
                get
                {
                    if (eingabeBrandLaenge == 0 || ergQuerSpann == 0)
                        return String.Empty;
                    else
                        return ((100 - v_VerBrand).ToString() + "/" + v_VerBrand);
                }
                set { }
            }
            //public double ergIk
            //{
            //    get
            //    {
            //        if (eingabeLeitungslaenge != 0 && eingabeCheckKurz)
            //            return (eingabeSpannungKurz * eingabeLeiter.value * ergQuerNorm) / (2 * eingabeLeitungslaenge);
            //        else
            //            return 0;
            //    }
            //    set { }
            //}y
        #endregion

        #region Querschnittergebnisse
        public double ergQuerSpann = 0;
            //{
            //    get
            //    {
            //        switch (eingabeSpannung.id)
            //        {
            //            case 1: //Wechselstrom
            //                return (eingabeBStrom * 2 * eingabeLeitungslaenge * eingabeCosPhi) / (eingabeLeiter.value * eingabeSpannAbfVolt);
            //            case 2: //Gleichstrom
            //                return (eingabeBStrom * 2 * eingabeLeitungslaenge) / (eingabeLeiter.value * eingabeSpannAbfVolt);
            //            case 3: //Drehstrom
            //                return (eingabeBStrom * Math.Sqrt(3) * eingabeLeitungslaenge * eingabeCosPhi) / (eingabeLeiter.value * eingabeSpannAbfVolt);
            //            default:
            //                return 0;
            //        }
            //    }
            //    set { }
            //}
            public double ergQuerSpannPara = 0;
            public double ergQuerBstrom = 0;
            public double ergQuerSicherung = 0;
            public double ergQuerBrand
            {
                get
                {
                    if (eingabeBrandLaenge != 0 && eingabeLeiter.value != 0 && v_SpannAbfBrand != 0 && ergQuerSpann != 0)
                    {
                        double brandFaktor = 0;
                        double m_quer = 0;

                        switch (eingabeSpannung.id)
                        {
                            case 1: //Wechselstrom
                                m_quer = (eingabeBStrom * 2 * eingabeLeitungslaenge * eingabeCosPhi) / (eingabeLeiter.value * v_SpannAbfBrand);
                                break;
                            case 2: //Gleichstrom
                                m_quer = (eingabeBStrom * 2 * eingabeLeitungslaenge) / (eingabeLeiter.value * v_SpannAbfBrand);
                                break;
                            case 3: //Drehstrom
                                m_quer = (eingabeBStrom * Math.Sqrt(3) * eingabeLeitungslaenge * eingabeCosPhi) / (eingabeLeiter.value * v_SpannAbfBrand);
                                break;
                            default:
                                m_quer = 0;
                                break;
                        }
                        if (eingabeBrandKlasse == "E 30")
                            brandFaktor = v_VerBrand * 0.0157 + 1;
                        else
                            brandFaktor = v_VerBrand * 0.036 + 1;
                        return m_quer * brandFaktor;
                    }
                    else
                        return 0;
                }
                set { }
            }
            public double ergQuerKurz = 0;
            public double ergQuerNorm = 0;
            public double ergParaQuerGerechnet
            {
                get
                {
                    if (v_iznAmax != 0)
                        return eingabeBStrom / v_iznAmax * v_maxA - eingabeParaKabel[0] - eingabeParaKabel[1] - eingabeParaKabel[2] - eingabeParaKabel[3] - eingabeParaKabel[4];
                    else
                        return 0;
                }
                set { }
            }
            public double ergParaQuerBestimmt = 0;
        #endregion

        private double getQuerschnittMax()
        {
            double max = 0;
            for (int i = 0; i < eingabeParaKabel.Length; i++)
                if (eingabeParaKabel[i] > max)
                    max = eingabeParaKabel[i];
            return max;
        }
        private static bool checkVersion(string fileVersion)
        {
            bool rc = true;
            //Projekt obj = new Projekt();
            Version ver = System.Reflection.Assembly.GetEntryAssembly().GetName().Version; // (obj.GetType()).GetName().Version;
            string[] parts = fileVersion.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            if (Int32.Parse(parts[0].ToString()) == ver.Major)
            {
                if (Int32.Parse(parts[1].ToString()) > ver.Minor)
                    rc = false;
            }
            else
                rc = false;
            return rc;
        }

        #region Clone function
        public object Clone()
        {
            MemoryStream ms = new MemoryStream();
            BinaryFormatter bf = new BinaryFormatter();
            bf.Serialize(ms, this);
            ms.Position = 0;
            object obj = bf.Deserialize(ms);
            ms.Close();
            return obj;
        }
        #endregion

        #region Override of Equals function
        private static bool inEquals = false;
        public override bool Equals(System.Object obj)
        {
            if (inEquals) return (obj == this);
            // If parameter is null return false.
            if (obj == null) return false;

            // If parameter cannot be cast to Point return false.
            Projekt p = obj as Projekt;
            if ((System.Object)p == null) return false;

            // Return true if the fields match:
            return Equals(p);
        }
        public bool Equals(Projekt p)
        {
            if (inEquals) return true;
            if ((object)p == null) return false;

            inEquals = true;
            //XmlSerializer mySerializer = new XmlSerializer(typeof(Projekt));
            //StringWriter xml1 = new StringWriter();
            //StringWriter xml2 = new StringWriter();
            //mySerializer.Serialize(xml1, this);
            //mySerializer.Serialize(xml2, p);

            bool rc = this.ToXml().Equals(p.ToXml());
            inEquals = false;
            return rc;
        }
        public override int GetHashCode()
        {
            return 0;
        }
        #endregion

        #region Serialization
        public string ToXml()
        {
            XmlSerializer mySerializer = new XmlSerializer(typeof(Projekt));
            MemoryStream memStream = new MemoryStream();
            StreamWriter myWriter = new StreamWriter(memStream);
            UTF8Encoding utf8e = new UTF8Encoding();
            try
            {
                inEquals = true;
                mySerializer.Serialize(myWriter, this);
                inEquals = false;
                myWriter.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Fehler beim Prüfen des Projekts.\n\n" + e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return String.Empty;
            }
            return utf8e.GetString(memStream.ToArray());
        }
        public static Projekt FromXml(string xml)
        {
            Projekt loaded;
            XmlSerializer mySerializer = new XmlSerializer(typeof(Projekt));
            StringReader myReader = new StringReader(xml);
            try
            {
                loaded = (Projekt)mySerializer.Deserialize(myReader);
                //if (loaded.version.Equals(String.Empty)) loaded.SetActVersion();
            }
            catch (Exception e)
            {
                MessageBox.Show("Fehler beim Lesen des Projekts.\n\n" + e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                loaded = null;
            }
            finally
            {
                myReader.Close();
            }
            return loaded;
        }
        public bool SaveToFile(string fileName)
        {
            try
            {
                string xml = ToXml();
                if (String.IsNullOrEmpty(xml)) return false;
                File.WriteAllText(fileName, xml);
            }
            catch (Exception e)
            {
                MessageBox.Show("Fehler beim Speichern des Projekts.\n\n" + e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }
        public static Projekt LoadFromFile(string fileName)
        {
            string xml;
            try
            {
                xml = File.ReadAllText(fileName);
            }
            catch
            {
                MessageBox.Show(String.Format("Fehler beim Laden der Datei '{0}'", fileName), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new Projekt();
            }
            return FromXml(xml);
        }
        public static Projekt LoadFromZip(string fileName)
        {
            ZipStorer zip = null;
            string tempDir = String.Empty;
            string tempXmlFile = String.Empty;
            string tempZipFile = String.Empty;
            string versionString = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString(3);
            List<ZipStorer.ZipFileEntry> dir;
            const string fmtInvalidVer = "Inkompatible Dateiversion.\r\n\r\nDie Datei wurde mit der Version {0} erstellt und kann nicht mit der aktuellen Version {1} gelesen werden.";

            try
            {
                try
                {
                    tempDir = Path.GetTempFileName(); File.Delete(tempDir); Directory.CreateDirectory(tempDir);
                    tempXmlFile = Path.GetTempFileName(); File.Delete(tempXmlFile);
                    tempZipFile = Path.Combine(tempDir, Path.GetFileName(tempDir));
                    tempXmlFile = Path.Combine(tempDir, Path.GetFileName(tempXmlFile));
                    File.Copy(fileName, tempZipFile, true);
                    zip = ZipStorer.Open(tempZipFile, FileAccess.ReadWrite);
                    dir = zip.ReadCentralDir();
                    if (checkVersion(dir[0].ExtraField))
                    {
                        zip.ExtractStoredFile(dir[0], tempXmlFile);
                    }
                    else
                    {
                        MessageBox.Show(string.Format(fmtInvalidVer, dir[0].ExtraField.ToString(), versionString),Settings.ProgramType,MessageBoxButtons.OK,MessageBoxIcon.Error);
                        return null;
                    }
                    return LoadFromFile(tempXmlFile);
                }
                catch (System.Exception)
                {
                    throw new Exception(string.Format("Datei konnte nicht geöffnet werden."));
                }

                //try
                //{
                //    return LoadFromFile(tempXmlFile);
                //}
                //catch (System.Exception)
                //{
                //    throw new Exception(string.Format(fmtInvalidVer, dir[0].ExtraField.ToString(), versionString));
                //}
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (zip != null) zip.Close();
                Directory.Delete(tempDir, true);
            }
            return null;
        }
        public bool SaveAsZip(string fileName)
        {
            if (SaveToFile(fileName))
                return ZipFile(fileName);
            else
                return false;
        }
        public bool ZipFile(string fileName)
        {
            ZipStorer zip = null;
            string tempPath = String.Empty;
            string tempDir = String.Empty;
            string versionString = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString(3);


            try
            {
                tempDir = Path.GetTempFileName(); File.Delete(tempDir); Directory.CreateDirectory(tempDir);
                tempPath = Path.Combine(tempDir, Path.GetFileName(fileName));
                zip = new ZipStorer();
                zip = ZipStorer.Create(tempPath, "", versionString);
                zip.AddFile(fileName, Path.GetFileName(fileName), "", versionString);
                zip.Close();
                File.Copy(tempPath, fileName, true);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Laden der Datei\n'" + fileName.ToString() + "':" + ex.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                try { zip.Close(); }
                catch { }
                Directory.Delete(tempDir, true);
            }
        }
        #endregion
    };

    #region Bearbeiter
    [Serializable] public class Bearbeiter
    {
        public Bearbeiter() : this(-1, String.Empty) { }
        public Bearbeiter(long id, string name)
        {
            this.id = id;
            this.name = name;
        }
        public long id;
        public string name;

        #region Clone function
        public object Clone()
        {
            MemoryStream ms = new MemoryStream();
            BinaryFormatter bf = new BinaryFormatter();
            bf.Serialize(ms, this);
            ms.Position = 0;
            object obj = bf.Deserialize(ms);
            ms.Close();
            return obj;
        }
        #endregion

        #region Override of Equals function
        private static bool inEquals = false;
        public override bool Equals(System.Object obj)
        {
            if (inEquals) return true;
            // If parameter is null return false.
            if (obj == null) return false;

            // If parameter cannot be cast to Point return false.
            Bearbeiter p = obj as Bearbeiter;
            if ((System.Object)p == null) return false;

            // Return true if the fields match:
            return Equals(p);
        }
        public bool Equals(Bearbeiter p)
        {
            if (inEquals) return true;
            if ((object)p == null) return false;

            inEquals = true;
            XmlSerializer mySerializer = new XmlSerializer(typeof(Bearbeiter));
            StringWriter xml1 = new StringWriter();
            StringWriter xml2 = new StringWriter();
            mySerializer.Serialize(xml1, this);
            mySerializer.Serialize(xml2, p);

            inEquals = false;
            return (xml1.ToString().Equals(xml2.ToString()));
        }
        public override int GetHashCode()
        {
            return 0;
        }
        #endregion
    };

    [Serializable] public class BearbeiterCollection : CollectionBase
    {
        #region Constructors
        public BearbeiterCollection() { }
        #endregion

        #region  Implementation of CollectionBase
        public int Add(Bearbeiter value)
        {
            return this.List.Add(value);
        }

        public void AddRange(Bearbeiter[] value)
        {
            for (int i = 0; i < value.Length; i++)
            {
                this.Add(value[i]);
            }
        }

        public void AddRange(BearbeiterCollection value)
        {
            for (int i = 0; i < value.Count; i++)
            {
                this.Add(value[i]);
            }
        }

        public bool Contains(Bearbeiter value)
        {
            return this.List.Contains(value);
        }

        public void CopyTo(Bearbeiter[] layArray, int intIndex)
        {
            this.List.CopyTo(layArray, intIndex);
        }

        //public new BearbeiterEnumerator GetEnumerator()
        //{
        //    return new BearbeiterEnumerator(this);
        //}

        public int IndexOf(Bearbeiter value)
        {
            return this.List.IndexOf(value);
        }

        public int IndexOf(long id)
        {
            int num3 = this.List.Count - 1;
            for (int i = 0; i <= num3; i++)
            {
                if (((Bearbeiter)this.List[i]).id.Equals(id))
                {
                    return i;
                }
            }
            return -1;
        }

        public int IndexOf(string bezeichnung)
        {
            for (int i = 0; i < this.List.Count; i++)
            {
                if (((Bearbeiter)this.List[i]).name.Equals(bezeichnung.Trim()))
                    return i;
            }
            return -1;
        }

        public void Insert(int intIndex, Bearbeiter value)
        {
            this.List.Insert(intIndex, value);
        }

        protected override void OnClear()
        {
            while (this.List.Count > 0)
            {
                Bearbeiter value = (Bearbeiter)this.List[0];
                this.Remove(value);
                value = null;
            }
        }

        protected override void OnInsert(int intIndex, object objValue)
        {
        }

        protected override void OnRemove(int intIndex, object objValue)
        {
        }

        protected override void OnSet(int intIndex, object objOldValue, object objNewValue)
        {
        }

        protected override void OnValidate(object objValue)
        {
        }

        public void Remove(Bearbeiter value)
        {
            this.List.Remove(value);
        }
        #endregion

        public SortedList<string, long> GetSorted()
        {
            SortedList<string, long> slist = new SortedList<string, long>();
            foreach (Bearbeiter v in this.List)
            {
                slist.Add(v.name, v.id);
            }
            return slist;
        }

        #region User Properties
        public Bearbeiter GetById(long id)
        {
            foreach (Bearbeiter v in this.List)
            {
                if (v.id.Equals(id)) return v;
            }
            return null;
        }
        public Bearbeiter Latest
        {
            get
            {
                Bearbeiter rc = null;
                long currId = long.MinValue;

                foreach (Bearbeiter v in this.List)
                {
                    if (v.id > currId) { currId = v.id; rc = v; }
                }
                return rc;
            }
        }
        #endregion

        #region Default accessor
        public Bearbeiter this[int intIndex]
        {
            get
            {
                return (Bearbeiter)this.List[intIndex];
            }
            set
            {
                this.List[intIndex] = value;
            }
        }
        #endregion

        #region Clone function
        public object Clone()
        {
            MemoryStream ms = new MemoryStream();
            BinaryFormatter bf = new BinaryFormatter();
            bf.Serialize(ms, this);
            ms.Position = 0;
            object obj = bf.Deserialize(ms);
            ms.Close();
            return obj;
        }
        #endregion

        #region Override of Equals function
        private static bool inEquals = false;
        public override bool Equals(System.Object obj)
        {
            if (inEquals) return true;
            // If parameter is null return false.
            if (obj == null) return false;

            // If parameter cannot be cast to Point return false.
            BearbeiterCollection p = obj as BearbeiterCollection;
            if ((System.Object)p == null) return false;

            // Return true if the fields match:
            return Equals(p);
        }
        public bool Equals(BearbeiterCollection p)
        {
            if (inEquals) return true;
            if ((object)p == null) return false;

            inEquals = true;
            XmlSerializer mySerializer = new XmlSerializer(typeof(BearbeiterCollection));
            StringWriter xml1 = new StringWriter();
            StringWriter xml2 = new StringWriter();
            mySerializer.Serialize(xml1, this);
            mySerializer.Serialize(xml2, p);

            inEquals = false;
            return (xml1.ToString().Equals(xml2.ToString()));
        }
        public override int GetHashCode()
        {
            return 0;
        }
        #endregion

        #region Serialization
        public static BearbeiterCollection FromXml(string xml)
        {
            BearbeiterCollection loaded;
            XmlSerializer mySerializer = new XmlSerializer(typeof(BearbeiterCollection));
            StringReader myReader = new StringReader(xml);
            try
            {
                loaded = (BearbeiterCollection)mySerializer.Deserialize(myReader);
                myReader.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Fehler beim Lesen der Bearbeiterdaten.\n\n" + e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                loaded = null;
            }
            return loaded;
        }
        public string ToXml()
        {
            XmlSerializer mySerializer = new XmlSerializer(typeof(BearbeiterCollection));
            MemoryStream memStream = new MemoryStream();
            StreamWriter myWriter = new StreamWriter(memStream);
            UTF8Encoding utf8e = new UTF8Encoding();
            try
            {
                mySerializer.Serialize(myWriter, this);
                myWriter.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Fehler beim Prüfen der Bearbeiterdaten.\n\n" + e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return String.Empty;
            }
            return utf8e.GetString(memStream.ToArray());
        }
        public static bool ZipFile(string fileName)
        {
            ZipStorer zip = null;
            string tempPath = String.Empty;
            string tempDir = String.Empty;
            string versionString = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString(3);

            try
            {
                tempDir = Path.GetTempFileName(); File.Delete(tempDir); Directory.CreateDirectory(tempDir);
                tempPath = Path.Combine(tempDir, Path.GetFileName(fileName));
                zip = ZipStorer.Create(tempPath, "", versionString);
                zip.AddFile(fileName, Path.GetFileName(fileName), "", versionString);
                zip.Close();
                File.Copy(tempPath, fileName, true);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler beim Laden der Datei\n'" + fileName.ToString() + "':" + ex.Message.ToString(),Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                try { zip.Close(); }
                catch { }
                Directory.Delete(tempDir, true);
            }
        }
        public static BearbeiterCollection LoadFromFile(string fileName)
        {
            string xml;
            try
            {
                xml = File.ReadAllText(fileName);
            }
            catch
            {
                MessageBox.Show(String.Format("Fehler beim Laden der Datei '{0}'", fileName), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            return FromXml(xml);
        }
        public static BearbeiterCollection LoadFromZip(string fileName)
        {
            ZipStorer zip = null;
            string tempDir = String.Empty;
            string tempXmlFile = String.Empty;
            string tempZipFile = String.Empty;
            string versionString = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString(3);
            List<ZipStorer.ZipFileEntry> dir;
            const string fmtInvalidVer = "Inkompatible Dateiversion.\r\n\r\nDie Datei wurde mit der Version {0} erstellt und kann nicht mit der aktuellen Version {1} gelesen werden.";

            try
            {
                try
                {
                    tempDir = Path.GetTempFileName(); File.Delete(tempDir); Directory.CreateDirectory(tempDir);
                    tempXmlFile = Path.GetTempFileName(); File.Delete(tempXmlFile);
                    tempZipFile = Path.Combine(tempDir, Path.GetFileName(tempDir));
                    tempXmlFile = Path.Combine(tempDir, Path.GetFileName(tempXmlFile));
                    File.Copy(fileName, tempZipFile, true);
                    zip = ZipStorer.Open(tempZipFile, FileAccess.ReadWrite);
                    dir = zip.ReadCentralDir();
                    if (checkVersion(dir[0].ExtraField))
                    {
                        zip.ExtractStoredFile(dir[0], tempXmlFile);
                    }
                    else
                    {
                        MessageBox.Show(string.Format(fmtInvalidVer, dir[0].ExtraField.ToString(), versionString), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return new BearbeiterCollection();
                    }
                    return LoadFromFile(tempXmlFile);
                }
                catch (System.Exception)
                {
                    throw new Exception(string.Format("Datei konnte nicht geöffnet werden."));
                }

                //try
                //{
                //    return LoadFromFile(tempXmlFile);
                //}
                //catch (System.Exception)
                //{
                //    throw new Exception(string.Format(fmtInvalidVer, dir[0].ExtraField.ToString(), versionString));
                //}
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (zip != null) zip.Close();
                Directory.Delete(tempDir, true);
            }
            return new BearbeiterCollection();
        }
        public bool SaveAsZip(string fileName)
        {
            if (SaveToFile(fileName))
                return ZipFile(fileName);
            else
                return false;
        }
        public bool SaveToFile(string fileName)
        {
            bool rc = false;
            try
            {
                string xml = ToXml();
                if (String.IsNullOrEmpty(xml)) return rc;
                File.WriteAllText(fileName, xml);
                rc = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Fehler beim Speichern der Normdaten.\n\n" + e.Message.ToString(), Settings.ProgramType, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return rc;
            }
            return rc;
        }

        private static bool checkVersion(string fileVersion)
        {
            bool rc = true;
            //Bearbeiter obj = new Bearbeiter();
            Version ver = System.Reflection.Assembly.GetEntryAssembly().GetName().Version; 
            //Version ver = System.Reflection.Assembly.GetAssembly(obj.GetType()).GetName().Version;
            string[] parts = fileVersion.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            if (Int32.Parse(parts[0].ToString()) == ver.Major)
            {
                if (Int32.Parse(parts[1].ToString()) > ver.Minor)
                    rc = false;
            }
            else
                rc = false;
            return rc;
        }
        #endregion
    };

    public class BearbeiterEnumerator : IEnumerator
    {
        #region Fields
        private IEnumerator iEnBaseEnumerator;
        #endregion

        #region Implementation of IEnumerator Methods
        public BearbeiterEnumerator(BearbeiterCollection mappings)
        {
            this.iEnBaseEnumerator = mappings.GetEnumerator();
        }

        public bool IEnumerator_MoveNext()
        {
            return this.iEnBaseEnumerator.MoveNext();
        }

        public void IEnumerator_Reset()
        {
            this.iEnBaseEnumerator.Reset();
        }

        public bool MoveNext()
        {
            return this.iEnBaseEnumerator.MoveNext();
        }

        public void Reset()
        {
            this.iEnBaseEnumerator.Reset();
        }
        #endregion

        #region Implementation of IEnumerator Properties
        public Bearbeiter Current
        {
            get
            {
                return (Bearbeiter)this.iEnBaseEnumerator.Current;
            }
        }

        public object IEnumerator_Current
        {
            get
            {
                return this.iEnBaseEnumerator.Current;
            }
        }

        object System.Collections.IEnumerator.Current
        {
            get
            {
                return this.iEnBaseEnumerator.Current;
            }
        }
        #endregion

    };
    #endregion
}