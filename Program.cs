using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Reflection;
using System.Diagnostics;

namespace dpr_smjenski
{
    class Program
    {
        private static object oMissing;

        static void Main(string[] args)
        {
            //Console.WriteLine("11111111111111111111111" + DateTime.Now);
            //Console.ReadKey();

            string connectionString = @"Data Source=192.168.0.3;Initial Catalog=FeroApp;User ID=sa;Password=AdminFX9.";
            string connectionStringRFIND = @"Data Source=192.168.0.3;Initial Catalog=RFIND;User ID=sa;Password=AdminFX9.";


            string c7 = "", c8 = "", c9 = "", c10 = "", c11 = "", c12 = "", c13 = "", c14 = "", c15 = "0", c151="", c16 = "", c17="",c19 = "", c21 = ""; // 
            string c7do = "", c8do = "", c9do = "", c10do = "", c11do = "", c12do = "", c13do = "", c14do = "", c15do = "", c151do ="", c16do = "", c17do = "", c19do = "", c21do = ""; // 
            double i7 = 0.0, i8 = 0.0, i9 = 0.0, i10 = 0.0, i11 = 0.0, i12 = 0.0, i13 = 0.0, i14 = 0.0, i15 = 0.0, i151=0, i16 = 0.0, i17=0.0, i18=0.0, i19 = 0.0, i21 = 0.0,i22=0; // 
            double f7 = 0.0, f8 = 0.0, f9 = 0.0, f10 = 0.0, f11 = 0.0, f12 = 0.0, f13 = 0.0, f14 = 0.0, f15 = 0.0, f151=0, f16 = 0.0, f17 = 0.0, f18 = 0.0, f19 = 0.0, f21 = 0.0, f22 = 0; // 
            double g7 = 0.0, g8 = 0.0, g9 = 0.0, g10 = 0.0, g11 = 0.0, g12 = 0.0, g13 = 0.0, g14 = 0.0, g15 = 0.0, g151=0, g16 = 0.0, g17 = 0.0, g18 = 0.0, g19 = 0.0, g21 = 0.0, g22 = 0; // 
            string j7 = "0", j8 = "0", j9 = "0", j10 = "0", j11 = "0", j12 = "0", j13 = "0", j14 = "0", j15 = "0", j151="0" ,j16 = "0", j17="0",j19 = "0", j21 = "0"; // 
            string sql1 = "";
            string dat1 = "2017-06-20", dat2 = "2017-06-20", dat3 = "", dat3p = "";

            // Process.Start("C:\\brisi\\_raspored djelatnika21.xlsm");

            DateTime d1 = new DateTime(2017, 6, 20);
            DateTime d10 = new DateTime(2017, 7, 27);
            //DateTime d2 = DateTime.Now.AddDays( -1 )  ; // smjena 2
          //d2 = d2.AddHours(-8);

            DateTime d3 = DateTime.Now;
            DateTime d3p = DateTime.Now;
            DateTime d2 = DateTime.Now;
//            d2 = new DateTime(2019, 3, 2,23, 20, 0);

            if (d2.Hour < 14 && 1 == 1)
            {
                d2 = d2.AddDays(-1);
//                d2 = d2.AddHours(-7);   // brisi
                d3p = d2;
            }

                 //druga smjena od prethodnog dana            
            d2 = d2.AddHours(15);

            // test on date
            //
            
            d3 = d2;

            string dan1 = d2.Day.ToString();
            string m1 = d2.Month.ToString();
            string g1 = d2.Year.ToString();
            int smjenaz = 2;                        // smjena ???

            DateTime input = d2;
            int delta = DayOfWeek.Monday - input.DayOfWeek;
            if (d2.DayOfWeek == DayOfWeek.Sunday)
            {
                delta = -6;
            }

            DateTime monday = input.AddDays(delta);
            delta = DayOfWeek.Sunday - input.DayOfWeek + 7;
            if (d2.DayOfWeek == DayOfWeek.Sunday)
            {
                delta = 0;
            }

            DateTime sunday = input.AddDays(delta);
            //monday= new DateTime(2017, 7, 31);
            sunday = monday.AddDays(6);
            //sunday = d2;

            string mm1 = "";
            if (d2.Month <= 9)
                mm1 = "0";

            dat1 = d2.Year.ToString() + '-' + mm1+d2.Month.ToString() + '-' + d2.Day.ToString();
            DateTime d30 = DateTime.Now;
            mm1 = "";
            if (d30.Month <= 9)
                mm1 = "0";
            string dats = d30.Year.ToString() + '-' + mm1 + d30.Month.ToString() + '-' + d30.Day.ToString();  // današnji datum
            dat2 = dat1;
            mm1 = "";
            if (d3.Month <= 9)
                mm1 = "0";
            dat3 = d3.Year.ToString() + '-' + mm1+d3.Month.ToString() + '-' + d3.Day.ToString();
            dat3p = d3p.Year.ToString() + '-' + d3p.Month.ToString() + '-' + d3p.Day.ToString();
            TimeSpan t = d2 - d1;
            int dana = t.Days;
            string nuland = "", nulanm = "", dnuland = "", dnulanm = "";
            if (d2.Day <= 9)
            {
                nuland = "0";
            }
            if (d2.Month <= 9)
            {
                nulanm = "0";
            }

            if (d3.Day <= 9)
            {
                dnuland = "0";
            }
            if (d3.Month <= 9)
            {
                dnulanm = "0";
            }

            string dat10 = nuland + d2.Day.ToString() + '.' + nulanm + d2.Month.ToString() + '.' + d2.Year.ToString();
            string dat30 = dnuland + d3.Day.ToString() + '.' + dnulanm + d3.Month.ToString() + '.' + d3.Year.ToString();   // današnji datum
            string dat13 = dat1 + " 6:00:00";
            DateTime d23 = d2.AddDays(1);
            //string dat23 = d2.Year.ToString() + '-' + d2.Month.ToString() + '-' + d2.Day.ToString() + " 6:00:00";

             mm1 = "";
            if (d2.Month <= 9)
                mm1 = "0";

            string dat23 = d2.Year.ToString() + '-' + mm1+d2.Month.ToString() + '-' + d2.Day.ToString();

            //string fileName = @"L:\izvještaji\dsr\dprs" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + ".xlsm";
            Console.WriteLine("Od datuma: " + dat1 + " - " + dat2 + " trenutno vrijeme " + DateTime.Now);
            string smjenanaziv = "", smj = "";
            DateTime datrep = d3; ;
            if (d2.Hour < 14 && d2.Hour >= 6)   // daj rezultat od 3 smjene
            {
                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',1,3";
                smjenanaziv = "3.smjena";
                smj = "3";
                datrep = d3;
                d23 = d2.AddDays(1);
                mm1 = "";
                if (d23.Month <= 9)
                    mm1 = "0";

                dat23 = d23.Year.ToString() + '-' + mm1 + d23.Month.ToString() + '-' + d23.Day.ToString() + " 6:00:00";
                dat23 = d23.Year.ToString() + '-' + mm1 + d23.Month.ToString() + '-' + d23.Day.ToString() ;
            }
            else
            {
                if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                {
                    sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',1,1";
                    smjenanaziv = "1.smjena";
                    smj = "1";
                    datrep = d3;
                }

                if ((d2.Hour > 21) || (d2.Hour < 8))  // u 24 sati daj komade od smjene 2
                {
                    sql1 = "rfind.dbo.realizacija2'" + dat1 + "','" + dat2 + "',1,2";
                    smjenanaziv = "2.smjena";
                    smj = "2";
                    datrep = d3;
                }
            }

            string fileName = @"l:\izvještaji\dsr\dprs" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + "_" + smj + ".xlsm";
            string fileNamebv = @"l:\izvještaji\dsr\_dprs" + nuland + d2.Day.ToString() + nulanm + d2.Month.ToString() + d2.Year.ToString() + "_" + smj + ".xlsm";
            
            if (1 == smjenaz)
            {
                sql1 = "rfind.dbo.realizacija2'" + dat1 + "','" + dat2 + "',1,2";
                smjenanaziv = "2.smjena";
                smj = "2";
                datrep = d3.AddDays(-1);
            }

            if (datrep.Day <= 9)
            {
                dnuland = "0";
            }
            if (datrep.Month <= 9)
            {
                dnulanm = "0";
            }
            string dat10sp = dat10;
            string dat1sp = dat1;
            string dat2sp = dat2;
            string datreps = dnuland + datrep.Day.ToString() + '.' + dnulanm + datrep.Month.ToString() + '.' + datrep.Year.ToString();   // današnji datum
            string datLDP = dat1;
            Application excel = new Application();
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            //Workbook workbook = excel.Workbooks.Open(@"c:\izvještaji\dsr\dprs_template1.xlsm", ReadOnly: false, Editable: true);
            // Workbook workbook = excel.Workbooks.Open(@"c:\brisi\dprs_template1.xlsm", ReadOnly: false, Editable: true);
//            Workbook workbook = excel.Workbooks.Open(@"c:\brisi\dprs_template1601.xlsm", ReadOnly: false, Editable: true);

            //            Workbook workbook = excel.Workbooks.Open(@"c:\brisi\dprs_template1304.xlsm", ReadOnly: false, Editable: true);
//            Workbook workbook = excel.Workbooks.Open(@"c:\brisi\dprs_template1907.xlsm", ReadOnly: false, Editable: true);
//            Workbook workbook = excel.Workbooks.Open(@"c:\brisi\dprs_template2007.xlsm", ReadOnly: false, Editable: true);
            Workbook workbook = excel.Workbooks.Open(@"c:\brisi\dprs_template1403.xlsm", ReadOnly: false, Editable: true);


            for (int ws = 1; ws <= 4; ws++)
            {
                double ukupv = 0.0, ukupk = 0.0;

                if (ws == 2)
                {
                    dat10 = monday.Day.ToString() + "." + monday.Month.ToString() + "." + monday.Year.ToString() + " - " + sunday.Day.ToString() + "." + sunday.Month.ToString() + "." + sunday.Year.ToString();
                    dat1 = monday.Year.ToString() + "-" + monday.Month.ToString() + "-" + monday.Day.ToString();
                    dat2 = sunday.Year.ToString() + "-" + sunday.Month.ToString() + "-" + sunday.Day.ToString();
                   
                }
                else
                {
                    dat10 = dat10sp;
                    dat1 = dat1sp;
                    dat2 = dat2sp;
                }


                // samo SONA
                // select kupac, VrstaNarudzbe, sum((obradaa * CijenaObradaA + obradab * CijenaObradab + obradac * CijenaObradac) * kolicinaok) vrijednost,sum(kolicinaok) kolicinaok
                // from feroapp.dbo.evidnormi(@dat1, @dat2, 0)
                // where kupac LIKE '%SONA%' AND OBRADAA = 1 and smjena = @smjena
                // group by kupac,VrstaNarudzbe

                c7 = ""; c8 = ""; c9 = ""; c10 = ""; c11 = ""; c12 = ""; c13 = ""; c14 = ""; c15 = "0"; c16 = ""; c19 = ""; c21 = "";// 
                ukupk = 0.0;

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as iddoubl;[ime x] as ime;[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (d2.Hour < 14 && d2.Hour > 5)   // daj rezultat od 3 smjene
                    {
                        sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',11,3";
                        smjenanaziv = "3.smjena";
                        smj = "3";
                    }
                    else
                    {
                        if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                        {
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',11,1";
                            smjenanaziv = "1.smjena";
                            smj = "1";
                        }

                        if ((d2.Hour > 21) || (d2.Hour < 6))  // u 24 sati daj komade od smjene 2
                        {
                            sql1 = "rfind.dbo.realizacija2' " + dat1 + "','" + dat2 + "',11,2";
                            smjenanaziv = "2.smjena";
                            smj = "2";
                        }

                    }
                    if (1 == smjenaz)
                    {
                        sql1 = "rfind.dbo.realizacija2'" + dat1 + "','" + dat2 + "',11,2";
                        smjenanaziv = "2.smjena";
                        smj = "2";
                        datrep = d3.AddDays(-1);
                    }


                    if (ws == 2)
                        sql1 = "rfind.dbo.realizacija'" + dat1 + "','" + dat2 + "',111";


                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap;


                  
                    while (reader.Read())
                    {
                        string idpar = reader["id_par"].ToString();

                        if (reader["kolicinaok"] != DBNull.Value)
                        {
                            if (idpar == "121301")   // SONA M
                            {
                                c15 = (reader["kolicinaok"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }
                            else if (idpar=="121302")   // SONA R
                            {
                                c151 = (reader["kolicinaok"].ToString());
                                ukupk = ukupk + (double.Parse)(reader["kolicinaok"].ToString());
                            }
                        }

                    }
                }

                // svi ostali osim sone
                // kupac, komada iz sELECT * FROM EvidencijaProizvedenoFakturirano_Zbirno_SG('2017-08-02', '2017-08-02')
                c7 = ""; c8 = ""; c9 = ""; c10 = ""; c11 = ""; c12 = ""; c13 = ""; c14 = ""; c16 = ""; c17 = ""; c19 = ""; // 
                c7do = ""; c8do = ""; c9do = ""; c10do = ""; c11do = ""; c12do = ""; c13do = ""; c17do = ""; c14do = ""; c15 = "0";c151 = "0"; c151do = "0"; c15do = "0"; c16do = ""; c19do = ""; // 
                g7 = 0.0; g8 = 0.0; g9 = 0.0; g10 = 0.0; g11 = 0.0; g12 = 0.0; g13 = 0.0; g14 = 0.0; g15 = 0.0; g16 = 0.0; g17 = 0.0; g18 = 0.0; g19 = 0.0; g21 = 0.0; g22 = 0; // 
                f7 = 0.0; f8 = 0.0; f9 = 0.0; f10 = 0.0; f11 = 0.0; f12 = 0.0; f13 = 0.0; f14 = 0.0; f15 = 0.0; f16 = 0.0; f17 = 0.0; f18 = 0.0; f19 = 0.0; f21 = 0.0; f22 = 0; // 
                ukupk = (int.Parse)(c15);
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as iddoubl;[ime x] as ime;[prezime x] as prezime;rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws < 2)
                        sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',1";
                    else
                        sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',31";


                    if (d2.Hour < 14 && d2.Hour > 5)   // daj rezultat od 3 smjene
                    {
                        sql1 = sql1 + ",3";
                        smjenanaziv = "3.smjena";
                        smj = "3";
                    }
                    else
                    {
                        if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                        {
                            sql1 = sql1 + ",1";
                            smjenanaziv = "1.smjena";
                            smj = "1";
                        }

                        if ((d2.Hour > 21) || (d2.Hour < 6))  // u 24 sati daj komade od smjene 2
                        {
                            sql1 = sql1 + ",2";
                            smjenanaziv = "2.smjena";
                            smj = "2";
                        }

                    }
                    if (1 == smjenaz)
                    {
                        sql1 = "rfind.dbo.realizacija2'" + dat1 + "','" + dat2 + "',1,2";
                        smjenanaziv = "2.smjena";
                        smj = "2";
                        datrep = d3.AddDays(-1);
                    }


                    if (ws == 2)
                    {
                        sql1 = "rfind.dbo.realizacija2' " + dat1 + "','" + dat2 + "',13,0";
                        //    sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',1";
                    }


                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac,idpar, vrstap, vo, hala1;

                    while (reader.Read())
                    {
                        kupac = reader["Kupac"].ToString();
                        idpar = reader["id_par"].ToString();
                        vrstap = reader["VrstaPro"].ToString();
                        vo = reader["vo"].ToString();
                        if (ws >= 3)
                        {
                            hala1 = reader["Hala"].ToString().TrimEnd();
                            if (hala1 != "3" && hala1 != "1")
                            {
                                continue;
                            }

                            if (((ws == 3) && (hala1 == "3")) || ((ws == 4) && (hala1 == "1")))
                            {
                                continue;
                            }
                        }

                        if (kupac.Contains("Austria"))
                        {
                            if (vo == "T")
                            {
                                c7 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                f7 = f7 + (double.Parse)( reader["vrijednost"].ToString() );
                            }
                            else
                            {
                                c7do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                g7 = g7 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                           
                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {
                            if (vo == "T")
                            {
                                c8 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                f8 = f8 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                c8do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                g8 = g8+ (double.Parse)(reader["vrijednost"].ToString());
                            }
                            
                        }

                        if (kupac.Contains("FAG"))
                        {
                            if (vo == "T")
                            {
                                c9 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                f9 = f9 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                c9do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                g9 = g9 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            
                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if (vo == "T")
                            {
                                c10 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                f10 = f10+ (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                c10do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                g10 = g10 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vrstap.Contains("Prsten"))  // wupertal
                            {
                                if (vo == "T")
                                {
                                    c11 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                    f11 = f11 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    c11do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                    g11 = g11 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                                
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            if (vrstap.Contains("Valj"))  // wupertal
                            {

                                if (vo == "T")
                                {
                                    c12 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                    f12 = f12 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    c12do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                    g12 = g12 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                                
                            }
                        }

                        if (kupac.Contains("KYSUCE"))  // 
                        {
                            if (vrstap.Contains("Ku"))  // Kysuce kučište
                            {
                                if (vo == "T")
                                {
                                    c13 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                    f13 = f13+ (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    c13do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                    g13 = g13 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                                

                            }
                            else  // kučište prsten
                            {
                                if (vo == "T")
                                {
                                    c17 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                    f17 = f17 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    c17do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                    g17 = g17 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                                
                            }

                        }

                        if (kupac.Contains("FERRO"))  // SKF
                        {
                            if (vo == "T")
                            {
                                c14 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                f14 = f14 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                c14do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                g14 = g14 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            
                        }


                        if (kupac.Contains("SONA"))  // sona
                        {

                            if (idpar == "121301")
                            {
                                if (vo == "T")
                                {
                                    c15 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                    f15 = f15 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    c15do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                    g15 = g15 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                            }
                            else
                            {
                                if (vo == "T")
                                {
                                    c151 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                    f151 = f151 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                                else
                                {
                                    c151do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                    g151 = g151 + (double.Parse)(reader["vrijednost"].ToString());
                                }
                            }
                            
                        }
                        

                        //                        if (kupac.Contains("SONA"))
                        //                        {
                        //                            c15 = (reader["Proizvedeno"].ToString());
                        //                           ukupk = ukupk + (double.Parse)(reader["Proizvedeno"].ToString());
                        //                       }

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            if (vo == "T")
                            {
                                c16 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                f16 = f16 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                c16do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                g16 = g16 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            
                        }

                        if (kupac.Contains("Brasil"))  // Brasil
                        {
                            if (vo == "T")
                            {
                                c19 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                f19 = f19 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                c19do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                g19 = g19 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            
                        }

                        if (kupac.Contains("NEU"))  // NSK
                        {
                            if (vo == "T")
                            {
                                c21 = (reader["kolicina"].ToString());   // proizvedeno tokarenjem, sve                                    
                                f21 = f21 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            else
                            {
                                c21do = (reader["kolicina"].ToString());    // proizvedeno dodatnom obradom                                    
                                g21   = g21 + (double.Parse)(reader["vrijednost"].ToString());
                            }
                            
                        }

                    }
                }

                // vrijednost obrade sve hale 112, pogoni 113
                // svi ostali osim sone

                /////////// sg 12.09.2017.

                // planirano = norma
                string b7 = "", b8 = "", b9 = "", b10 = "", b11 = "", b12 = "", b13 = "", b14 = "", b15 = "", b151="",b16 = "", b17 = "", b18 = "", b19 = "", b20 = "", b21 = "";
                string d70 = "", d80 = "", d90 = "", d100 = "", d110 = "", d120 = "", d130 = "", d140 = "", d150 = "",d1510="", d160 = "", d170 = "", d180 = "", d190 = "", d210 = "";
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws <= 2)
                        sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat3 + "',12";
                    else
                        sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat3 + "',121";

                    if (d2.Hour < 14 && d2.Hour > 5)   // daj rezultat od 3 smjene
                    {
                        sql1 = sql1 + ",3";
                        smjenanaziv = "3.smjena";
                        smj = "3";
                    }
                    else
                    {
                        if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                        {
                            sql1 = sql1 + ",1";
                            smjenanaziv = "1.smjena";
                            smj = "1";
                        }

                        if ((d2.Hour > 21) || (d2.Hour < 6))  // u 24 sati daj komade od smjene 2
                        {
                            sql1 = sql1 + ",2";
                            smjenanaziv = "2.smjena";
                            smj = "2";
                        }

                    }
                    if (1 == smjenaz)
                    {
                        sql1 = "rfind.dbo.realizacija2' " + dat1 + "','" + dat2 + "',12,2";
                        smjenanaziv = "2.smjena";
                        smj = "2";
                        datrep = d3.AddDays(-1);
                    }

                    if (ws == 2)
                    {
                        //sql1 = "rfind.dbo.ldp_recalc' " + dat1 + "','" + dat3 + "',101";
                        sql1 = "rfind.dbo.realizacija2' " + dat1 + "','" + dat3p + "',101," + smj;
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap, vnar = "", vo = "", hala1 = "";

                    while (reader.Read())
                    {
                        kupac = reader["Kupac"].ToString();
                        string idpar = reader["id_par"].ToString();


                        if (ws >= 3)
                        {
                            hala1 = reader["Hala"].ToString().TrimEnd();
                            if (hala1 != "3" && hala1 != "1")
                            {
                                continue;
                            }

                            if (((ws == 3) && (hala1 == "3")) || ((ws == 4) && (hala1 == "1")))
                            {
                                continue;
                            }
                        }

                        vnar = reader["vrstanarudzbe"].ToString();
                        vo = reader["vo"].ToString();
                        if (kupac.Contains("Austria"))
                        {
                            if (vo == "T")
                            {
                                b7 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d70 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {
                            if (vo == "T")
                            {
                                b8 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d80 = (reader["norma"].ToString());
                            }


                        }

                        if (kupac.Contains("FAG"))
                        {
                            if (vo == "T")
                            {
                                b9 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d90 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if (vo == "T")
                            {
                                b10 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d100 = (reader["norma"].ToString());
                            }



                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vnar.Contains("Prsteni"))  // wupertal
                            {
                                if (vo == "T")
                                {
                                    b11 = (reader["norma"].ToString());
                                }
                                else
                                {
                                    d110 = (reader["norma"].ToString());
                                }


                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            if (vnar.Contains("Valjci"))  // wupertal
                            {
                                if (vo == "T")
                                {
                                    b12 = (reader["norma"].ToString());
                                }
                                else
                                {
                                    d120 = (reader["norma"].ToString());
                                }


                            }
                        }

                        if (kupac.Contains("KYSUCE"))  // 
                        {
                            if (vnar.Contains("Ku"))  // 
                            {
                                if (vo == "T")
                                {
                                    b13 = (reader["norma"].ToString());
                                }
                                else
                                {
                                    d130 = (reader["norma"].ToString());
                                }
                            }
                        }
                        if (kupac.Contains("KYSUCE"))  // 
                        {
                            if (vnar.Contains("Prsten"))  // 
                            {

                                if (vo == "T")
                                {
                                    b17 = (reader["norma"].ToString());
                                }
                                else
                                {
                                    d170 = (reader["norma"].ToString());
                                }
                            }
                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {
                            if (vo == "T")
                            {
                                b14 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d140 = (reader["norma"].ToString());
                            }
                        }

                        if (kupac.Contains("SONA"))
                        {
                            if (idpar == "121301")    // sona m
                            {
                                if (vo == "T")
                                {
                                    b15 = (reader["norma"].ToString());
                                }
                                else
                                {
                                    d150 = (reader["norma"].ToString());
                                }
                            }
                            else if (idpar == "121302")   // sona r
                            {
                                if (vo == "T")
                                {
                                    b151 = (reader["norma"].ToString());
                                }
                                else
                                {
                                    d1510 = (reader["norma"].ToString());
                                }

                            }

                        }
                                                                       

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            if (vo == "T")
                            {
                                b16 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d160 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("Brasil"))  // Brasil
                        {
                            if (vo == "T")
                            {
                                b19 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d190 = (reader["norma"].ToString());
                            }

                        }

                        if (kupac.Contains("NEU"))  // NSK
                        {
                            if (vo == "T")
                            {
                                b21 = (reader["norma"].ToString());
                            }
                            else
                            {
                                d210 = (reader["norma"].ToString());
                            }

                        }

                    }
                    cn.Close();

                }
                Console.WriteLine("Izračunat broj štelanja " + DateTime.Now);
                j7 = "0"; j8 = "0"; j9 = "0"; j10 = "0"; j11 = "0"; j12 = "0"; j13 = "0"; j14 = "0"; j15 = "0";j151 = "0"; j16 = "0"; j17 = "0"; j19 = "0"; j21 = "0"; // 

                /////////// sg 030.08
                // broj stelanja
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                    if (d2.Hour < 14 && d2.Hour > 5)   // daj rezultat od 3 smjene
                    {
                        if (ws <= 2)
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat23 + "',10,3";
                        else
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat23 + "',102,3";

                        smjenanaziv = "3.smjena";
                        smj = "3";
                    }
                    else
                    {
                        if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                        {
                            if (ws <= 2)
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat23 + "',10,1";
                            else
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat23 + "',102,1";

                            smjenanaziv = "1.smjena";
                            smj = "1";
                        }

                        if ((d2.Hour > 21) || (d2.Hour < 6))  // u 24 sati daj komade od smjene 2
                        {
                            if (ws <= 2)
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat23 + "',10,2";
                            else
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat23 + "',102,2";

                            smjenanaziv = "2.smjena";
                            smj = "2";
                        }

                    }
                    if (1 == smjenaz)
                    {
                        sql1 = "rfind.dbo.realizacija2' " + dat1 + "','" + dat23 + "',10,2";
                        smjenanaziv = "2.smjena";
                        smj = "2";
                        datrep = d3.AddDays(-1);
                    }


                    if (ws == 2)
                        sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat23 + "',102";

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vrstap, hala1;

                    while (reader.Read())
                    {
                        kupac = reader["Kupac"].ToString();
                        string idpar = reader["id_par"].ToString();
                        vrstap = reader["Vrstanarudzbe"].ToString();
                        if (ws != 2)
                            hala1 = reader["Hala"].ToString().TrimEnd();
                        else
                            hala1 = "A";

                        if (kupac.Contains("Austria"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                j7 = (reader["broj_stelanja"].ToString());  // BDF
                            }

                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                j8 = (reader["broj_stelanja"].ToString());  // BMW
                            }

                        }

                        if (kupac.Contains("FAG"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                j9 = (reader["broj_stelanja"].ToString());  // DEbrecin
                            }

                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                j10 = (reader["broj_stelanja"].ToString());  // Brašov
                            }

                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vrstap.Contains("Prsteni"))  // wupertal
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    j11 = (reader["broj_stelanja"].ToString());  // 
                                }

                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vrstap.Contains("Valjci"))  // wupertal
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    j12 = (reader["broj_stelanja"].ToString());  // 
                                }

                            }
                        }

                        if (kupac.Contains("KYSUCE"))  // 
                        {
                            if (vrstap.Contains("Ku"))  // 
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    j13 = (reader["broj_stelanja"].ToString());  // 
                                }
                            }
                        }

                        if (kupac.Contains("KYSUCE"))  // 
                        {
                            if (vrstap.Contains("Prsten"))  //
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    j17 = (reader["broj_stelanja"].ToString());  // 
                                }
                            }

                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                j14 = (reader["broj_stelanja"].ToString());  // 
                            }

                        }

                        if (kupac.Contains("SONA"))
                        {
                            if (idpar == "121301")
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    j15 = (reader["broj_stelanja"].ToString());  // 
                                }
                            }
                            else if (idpar == "121302")
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    j151 = (reader["broj_stelanja"].ToString());  // 
                                }
                            }

                        }

                        

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                j16 = (reader["broj_stelanja"].ToString());  // 
                            }

                        }

                        if (kupac.Contains("Brasil"))  // Brasil
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                j19 = (reader["broj_stelanja"].ToString());  // 
                            }

                        }
                        if (kupac.Contains("NEU"))  // NSK
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                j21 = (reader["broj_stelanja"].ToString());  // 
                            }

                        }

                    }
                    cn.Close();

                }
                Console.WriteLine("Izračunat broj štelanja " + DateTime.Now);

                #region             // kupac, broj linija koje ne rade
                int n7 = 0, n8 = 0, n9 = 0, n10 = 0, n11 = 0, n12 = 0, n13 = 0, n14 = 0, n15 = 0, n151=0, n16 = 0, n17 = 0, n18 = 0, n19 = 0, n20 = 0, n21 = 0;
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);

                    if (d2.Hour < 14 && d2.Hour > 5)   // daj rezultat od 3 smjene
                    {
                        if (ws <= 2)
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',61,3";
                        else
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',62,3";

                        smjenanaziv = "3.smjena";
                        smj = "3";
                    }
                    else
                    {
                        if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                        {
                            if (ws <= 2)
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',61,1";
                            else
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',62,1";
                            smjenanaziv = "1.smjena";
                            smj = "1";
                        }

                        if ((d2.Hour > 21) || (d2.Hour < 6))  // u 24 sati daj komade od smjene 2
                        {
                            if (ws <= 2)
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',61,2";
                            else
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',62,2";

                            smjenanaziv = "2.smjena";
                            smj = "2";
                        }

                    }
                    if (1 == smjenaz)
                    {
                        if (ws <= 2)
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',61,2";
                        else
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',62,2";

                        smjenanaziv = "2.smjena";
                        smj = "2";
                        datrep = d3.AddDays(-1);
                    }


                    //sql1 = "rfind.dbo.ldp_recalc '" + dat1 + "','" + dat2 + "',61";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vnar = "", hala1;

                    while (reader.Read())
                    {

                        kupac = reader["Kupac"].ToString();
                        string idpar = reader["id_par"].ToString();
                        vnar  = reader["vrstanarudzbe"].ToString();
                        hala1 = reader["Hala"].ToString().TrimEnd();

                        if (kupac.Contains("Austria"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                n7 = (int.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                n8 = (int.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("FAG"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                n9 = (int.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                n10 = (int.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if (vnar.Contains("Prsteni"))  // wupertal
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    n11 = (int.Parse)(reader["broj_linija"].ToString());  // 
                                }
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {

                            if (vnar.Contains("Valjci"))  // wupertal
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    n12 = (int.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                        }

                        if (kupac.Contains("KYSUCE"))  // wupertal
                        {
                            if (vnar.Contains("Ku"))  // 
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    n13 = (int.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                        }

                        if (kupac.Contains("KYSUCE"))  // 
                        {
                            if (vnar.Contains("Prsten"))  // 
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    n17 = (int.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                n14 = (int.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("SONA BLW"))
                        {
                            if (idpar == "121301")
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    n15 = (int.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                            else if (idpar == "121302")
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    n151 = (int.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                        }

                        
                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                n16 = (int.Parse)(reader["broj_linija"].ToString());  //
                            }
                        }

                        if (kupac.Contains("Brasil"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                n19 = (int.Parse)(reader["broj_linija"].ToString());  //
                            }
                        }

                        if (kupac.Contains("NEUWEG"))  // nsk
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                n21 = (int.Parse)(reader["broj_linija"].ToString());  //
                            }
                        }

                    }
                    cn.Close();

                }
                #endregion  // broj linija koje ne rade


                // kupac, broj linija
                i7 = 0.0; i8 = 0.0; i9 = 0.0; i10 = 0.0; i11 = 0.0; i12 = 0.0; i13 = 0.0; i14 = 0.0; i15 = 0.0;i151 = 0.0; i16 = 0.0; i17 = 0; i18 = 0; i19 = 0.0; i21 = 0.0; i22 = 0; // 

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (d2.Hour < 14 && d2.Hour > 5)   // daj rezultat od 3 smjene
                    {
                        if (ws <= 2)
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',2,3";
                        else
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',21,3";

                        smjenanaziv = "3.smjena";
                        smj = "3";
                    }
                    else
                    {
                        if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                        {
                            if (ws <= 2)
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',2,1";
                            else
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',21,1";
                            smjenanaziv = "1.smjena";
                            smj = "1";
                        }

                        if ((d2.Hour > 21) || (d2.Hour < 6))  // u 24 sati daj komade od smjene 2
                        {
                            if (ws <= 2)
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',2,2";
                            else
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',21,2";
                            smjenanaziv = "2.smjena";
                            smj = "2";
                        }

                    }
                    if (1 == smjenaz)
                    {
                        sql1 = "rfind.dbo.realizacija2'" + dat1 + "','" + dat2 + "',2,2";
                        smjenanaziv = "2.smjena";
                        smj = "2";
                        datrep = d3.AddDays(-1);
                    }

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vnar = "", hala1;

                    while (reader.Read())
                    {
                        hala1 = reader["hala"].ToString().TrimEnd();
                        if (reader["Kupac"] == DBNull.Value)
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i22 = (double.Parse)(reader["broj_linija"].ToString());
                                continue;
                            }
                        }
                        kupac = reader["Kupac"].ToString();
                        string idpar = reader["id_par"].ToString();
                        vnar = reader["vrstanarudzbe"].ToString().Trim();

                        if (kupac.Contains("Austria"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i7 = (double.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i8 = (double.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("FAG"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i9 = (double.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i10 = (double.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                if (vnar.Contains("Prsteni"))  // wupertal
                                {
                                    i11 = (double.Parse)(reader["broj_linija"].ToString());  // 
                                }
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                if (vnar.Contains("Valjci"))  // wupertal
                                {
                                    i12 = (double.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                        }

                        if (kupac.Contains("KYSUCE"))  //
                        {
                            if (vnar.Contains("Ku"))  // 
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    i13 = (double.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                        }

                        if (kupac.Contains("KYSUCE"))  //
                        {
                            if (vnar.Contains("Prsten"))  // 
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    i17 = (double.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i14 = (double.Parse)(reader["broj_linija"].ToString());
                            }
                        }

                        if (kupac.Contains("SONA"))
                        {
                            if (idpar == "121301")
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    i15 = (double.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                            else if (idpar == "121302")
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    i151 = (double.Parse)(reader["broj_linija"].ToString());
                                }
                            }
                        }
                                                

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i16 = (double.Parse)(reader["broj_linija"].ToString());  //
                            }
                        }

                        if (kupac.Contains("TECHNOLOGIES"))  // eltmann 121987
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i18 = (double.Parse)(reader["broj_linija"].ToString());  //
                            }
                        }



                        if (kupac.Contains("Brasil"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i19 = (double.Parse)(reader["broj_linija"].ToString());  //
                            }
                        }


                        if (kupac.Contains("NEU"))  // NSK
                        {
                            if (ws == 3)
                            {
                                hala1 = hala1;
                            }

                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i21 = (double.Parse)(reader["broj_linija"].ToString());  //
                            }
                        }

                    }
                    cn.Close();

                }

                // kupac, korekcija broja linija
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (d2.Hour < 14 && d2.Hour > 5)   // daj rezultat od 3 smjene
                    {
                        if (ws <= 2)
                            sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',22,3";
                        else
                            sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',221,3";

                        smjenanaziv = "3.smjena";
                        smj = "3";
                    }
                    else
                    {
                        if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                        {
                            if (ws <= 2)
                                sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',22,1";
                            else
                                sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',221,1";

                            smjenanaziv = "1.smjena";
                            smj = "1";
                        }

                        if ((d2.Hour > 21) || (d2.Hour < 6))  // u 24 sati daj komade od smjene 2
                        {
                            if (ws <= 2)
                                sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',22,2";
                            else
                                sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',221,2";

                            smjenanaziv = "2.smjena";
                            smj = "2";
                        }

                    }
                    if (1 == smjenaz)
                    {
                        if (ws <= 2)
                            sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',22,2";
                        else
                            sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',221,2";

                        smjenanaziv = "2.smjena";
                        smj = "2";
                        datrep = d3.AddDays(-1);
                    }


                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, vnar = "", hala1;
                    double ukupno = 0.0, ukupno1 = 0.0;

                    while (reader.Read())
                    {

                        kupac = reader["Kupac"].ToString();
                        string idpar = reader["id_par"].ToString();
                        hala1 = reader["Hala"].ToString().TrimEnd();

                        if (kupac.Contains("_Ukupno"))
                        {
                            if (reader["ukupnotr"] != DBNull.Value)
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    ukupno = (double.Parse)(reader["ukupnotr"].ToString());
                                }
                            }
                        }

                        if (reader["ukupnotr"] != DBNull.Value)
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                ukupno1 = (double.Parse)(reader["ukupnotr"].ToString());
                            }
                        }
                        else
                        {
                            ukupno1 = 0.0;
                        }

                        if (kupac.Contains("Austria"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                if (reader["ukupnotr"] != DBNull.Value)
                                {
                                    i7 = i7 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                            }
                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                if (reader["ukupnotr"] != DBNull.Value)
                                {
                                    i8 = i8 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                            }
                        }

                        if (kupac.Contains("FAG"))
                        {
                            if (reader["ukupnotr"] != DBNull.Value)
                            { 
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    i9 = i9 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                            }
                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i10 = i10 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                if (reader["ukupnotr"] != DBNull.Value)
                                    i11 = i11 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }
                        }



                        if (kupac.Contains("KYSUCE"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i13 = i13 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }
                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i14 = i14 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }
                        }

                        if (kupac.Contains("SONA BLW"))
                        {
                            if (idpar == "121301")
                            {

                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    i15 = i15 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                            }
                            else if (idpar == "121302")
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {
                                    i151 = i151 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                            }
                        }

                        

                        if (kupac.Contains("SIGMA"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i16 = i16 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }
                        }

                        if (kupac.Contains("TECHNOLOGIES"))  // eltmann
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i18 = i18 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }
                        }



                        if (kupac.Contains("Brasil"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i19 = i19 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }
                        }

                        if (kupac.Contains("NEU"))  // NSK
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                i21 = i21 - 1.0 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }
                        }

                    }
                    cn.Close();

                }


                if (ws <= 4)   // nova verzija od 26.03.
                {
                    // kupac, korekcija broja linija
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        i7 = 0.0; i8 = 0.0; i9 = 0.0; i10 = 0.0; i11 = 0.0; i12 = 0.0; i13 = 0.0; i14 = 0.0; i15 = 0.0;i151 = 0; i16 = 0.0; i17 = 0.0; i19 = 0.0; i18 = 0.0; i21 = 0.0; i22 = 0; // 
                        if (ws<=3)
                        {
                            i22 = 69;
                        }
                        if (ws == 3)
                        {
                            i22 = 41;
                        }
                        if (ws == 4)
                        {
                            i22 = 28;
                        }

                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        //if (ws <= 0)
                        //{
                        //    //sql1 = "rfind.dbo.realizacija '" + dat1 + "','" + dat2 + "',22";
                        //}
                        //else
                        //{
                        //    sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',221";
                        //}
                        if (d2.Hour < 14 && d2.Hour > 5)   // daj rezultat od 3 smjene
                        {
                            if (ws <= 4)
                                sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',223,3";
                            //else
                            //    sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',221,3";

                            smjenanaziv = "3.smjena";
                            smj = "3";
                        }
                        else
                        {
                            if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                            {
                                if (ws <= 4)
                                    sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',223,1";
                                //else
                                //    sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',221,1";

                                smjenanaziv = "1.smjena";
                                smj = "1";
                            }

                            if ((d2.Hour > 21) || (d2.Hour < 6))  // u 24 sati daj komade od smjene 2
                            {
                                if (ws <= 4)
                                    sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',223,2";
                                //else
                                //    sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',221,2";

                                smjenanaziv = "2.smjena";
                                smj = "2";
                            }

                        }
                        if (1 == smjenaz)
                        {
                            if (ws <= 4)
                                sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',223,2";
                            //else
                            //    sql1 = "rfind.dbo.realizacija2 ' " + dat1 + "','" + dat2 + "',221,2";

                            smjenanaziv = "2.smjena";
                            smj = "2";
                            datrep = d3.AddDays(-1);
                        }



                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string kupac, vnar = "", Linijanaziv = "", hala = "";
                        double ukupno = 0.0;
                        int red11 = 1;
                        double var1 = 0.0;

                        while (reader.Read())
                        {

                            hala = reader["hala"].ToString();
                            if (ws >= 3)
                            {
                                if ((ws == 3) && hala.Contains("3")) // pušta samo halu 1
                                {
                                    continue;
                                }
                                if ((ws == 4) && hala.Contains("1"))  // pušta samo 3
                                {
                                    continue;
                                }
                            }
                            kupac = reader["Kupac"].ToString();
                            string idpar = reader["id_par"].ToString();
                            if (reader["ukupnotr"] != DBNull.Value)
                            {

                            }
                            else
                            {
                                continue;
                            }


                            Linijanaziv = reader["naziv"].ToString();

                            if (kupac.Contains("_Ukupno"))
                            {
                                if (reader["ukupnotr"] != DBNull.Value)
                                {
                                    ukupno = (double.Parse)(reader["ukupnotr"].ToString());
                                }
                                else
                                {
                                    ukupno = 0.0;
                                }
                            }

                            if (kupac.Contains("Austria"))
                            {
                                if (ws == 3)
                                {
                                    int hj = 0;
                                }
                                if (reader["ukupnotr"] != DBNull.Value)
                                {
                                    var1 = (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                                else
                                {
                                    var1 = 0.0;
                                }
                                i7 = i7 + var1;
                                Console.WriteLine("Linija" + Linijanaziv + " udio " + var1.ToString() + "i7 " + i7.ToString());
                                //if (ws == 3)
                                //{
                                //    works10.Cells[red11, 1].Value = Linijanaziv;
                                //    works10.Cells[red11, 2].Value = hala;
                                //    works10.Cells[red11, 3].Value = var1;
                                //    works10.Cells[red11, 4].Value = i7;
                                //    works10.Cells[red11, 5].Value = ws;
                                //    red11++;
                                //}

                            }

                            if (kupac.Contains("SCHWEINFURT"))
                            {
                                if (reader["ukupnotr"] != DBNull.Value)
                                {
                                    i8 = i8 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }

                            }

                            if (kupac.Contains("FAG"))
                            {
                                if (reader["ukupnotr"] != DBNull.Value)
                                {
                                    i9 = i9 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                                else
                                {

                                }

                            }

                            if (kupac.Contains("ROMANIA"))
                            {
                                i10 = i10 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("Technologies"))  // wupertal
                            {

                                i11 = i11 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("KYSUCE"))  // wupertal
                            {
                                i13 = i13 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("FERRO"))  // wupertal
                            {

                                i14 = i14 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("SONA"))
                            {
                                if (idpar == "121301")
                                {
                                    i15 = i15 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                                else if (idpar == "121302")
                                {
                                    i151 = i151 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                            }
                                                        
                            if (kupac.Contains("SIGMA"))  // wupertal
                            {
                                i16 = i16 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("TECHNOLOGIES"))  // eltman
                            {
                                i18 = i18 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("Brasil"))  // wupertal
                            {
                                i19 = i19 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                            }

                            if (kupac.Contains("NEUWEG"))  // NSK
                            {
                                if (reader["ukupnotr"] != DBNull.Value)
                                {
                                    i21 = i21 + (double.Parse)(reader["ukupnotr"].ToString()) / ukupno;
                                }
                                else
                                {
                                    i21 = 1;
                                }




                            }
                            if (ws <= 3)
                            {
                               // i22 = 1;
                            }

                        }
                        cn.Close();

                    }
                }

                Console.WriteLine("Izračunate linije   trenutno vrijeme " + DateTime.Now);
                // kupac, škart
                double e7tk = 0.00, e7ts = 0.00, post7t = 0.00, e13tk = 0.00, e13ts = 0.00, post13t = 0.00, e17ts = 0.0, post17tt, e8tk = 0.00, e8ts = 0.00, post8t = 0.00, e9tk = 0.00, e9ts = 0.00, post9t = 0.00, e10tk = 0.00, e10ts = 0.00, post10t = 0.00;
                double e11tk = 0.00, e11ts = 0.00, post11t = 0.00, e12tk = 0.00, e12ts = 0.00, post12t = 0.00, e14tk = 0.00, e14ts = 0.00, post14t = 0.00, e15tk = 0.00, e15ts = 0.00, e151tk = 0.00, e151ts = 0.00, post15t = 0.00,post151t=0.00, e16tk = 0.00, e16ts = 0.00, post16t = 0.00, e17tk = 0.0, post17t = 0.0, e19ts = 0.0, e19tk = 0.00, post19t = 0.00, e21tk = 0.00, post21t = 0.00, e21ts = 0.00;
                double post7tm = 0.00, post8tm = 0.00, post9tm = 0.00, post10tm = 0.00, post11tm = 0.00, post12tm = 0.00, post13tm = 0.00, post14tm = 0.00, post15tm = 0.00, post151tm=0.00, post16tm = 0.00, post17tm = 0.0, post19tm = 0.00, post21tm = 0.00;
                double e7ko = 0.00, e8ko = 0.00, e9ko = 0.00, e10ko = 0.00, e11ko = 0.00, e12ko = 0.00, e13ko = 0.00, e14ko = 0.00, e15ko = 0.00,e151ko=0.00, e16ko = 0.00, e17ko = 0.00, e18ko = 0.00, e19ko = 0.00, e21ko = 0.00;

                // škart

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (d2.Hour < 14 && d2.Hour > 5)   // daj rezultat od 3 smjene
                    {
                        if (ws <= 2)
                        {
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',3,3";
                        }
                        else
                        {
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',33,3";
                        }

                        smjenanaziv = "3.smjena";
                        smj = "3";
                    }
                    else
                    {
                        if (d2.Hour < 22 && d2.Hour >= 14)  // u 16 sati daj komade od smjene 1
                        {
                            if (ws <= 2)
                            {
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',3,1";
                            }
                            else
                            {
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',33,1";
                            }

                            smjenanaziv = "1.smjena";
                            smj = "1";
                        }

                        if ((d2.Hour > 21) || (d2.Hour < 6))  // u 24 sati daj komade od smjene 2
                        {
                            if (ws <= 2)
                            {
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',3,2";
                            }
                            else
                            {
                                sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',33,2";
                            }

                            smjenanaziv = "2.smjena";
                            smj = "2";
                        }

                    }
                    if (1 == smjenaz)
                    {
                        if (ws <= 2)
                        {
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',3,2";
                        }
                        else
                        {
                            sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',33,2";
                        }

                        smjenanaziv = "2.smjena";
                        smj = "2";
                        datrep = d3.AddDays(-1);
                    }


                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string kupac, turm, hala1;

                    while (reader.Read())
                    {
                        kupac = reader["Kupac"].ToString();
                        string idpar = reader["id_par"].ToString();
                        turm = reader["Turm"].ToString();
                        hala1 = reader["hala"].ToString().TrimEnd();
                        if (kupac.Contains("Austria"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (turm.Contains("Da"))
                                {
                                    e7tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e7ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e7ko = e7ts;
                                    if (e7tk > 0)
                                        post7t = post7t + e7ts / (e7tk + e7ts);

                                    e7ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e7tk > 0)
                                        post7tm = post7tm + e7ts / (e7tk + e7ts);

                                }
                                else
                                {
                                    e7tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e7ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e7ko = e7ts;
                                    if (e7tk > 0)
                                        post7t = post7t + e7ts / (e7tk + e7ts);

                                    e7ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e7tk > 0)
                                        post7tm = post7tm + e7ts / (e7tk + e7ts);

                                }
                            }
                        }

                        if (kupac.Contains("SCHWEINFURT"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (turm.Contains("Da"))
                                {
                                    e8tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e8ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e8ko = e8ts;
                                    if (e8tk > 0)
                                        post8t = post8t + e8ts / (e8tk + e8ts);

                                    e8ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e8tk > 0)
                                        post8tm = post8tm + e8ts / (e8tk + e8ts);

                                }
                                else
                                {
                                    e8tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e8ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e8ko = e8ts;
                                    if (e8tk > 0)
                                        post8t = post8t + e8ts / (e8tk + e8ts);

                                    e8ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e8tk > 0)
                                        post8tm = post8tm + e8ts / (e8tk + e8ts);

                                }

                            }
                        }

                        if (kupac.Contains("FAG"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (turm.Contains("Da"))
                                {
                                    e9tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e9ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e9ko = e9ts;
                                    if (e9tk > 0)
                                        post9t = post9t + e9ts / (e9tk + e9ts);

                                    e9ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e9tk > 0)
                                        post9tm = post9tm + e9ts / (e9tk + e9ts);

                                }
                                else
                                {
                                    e9tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e9ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e9ko = e9ts;
                                    if (e9tk > 0)
                                        post9t = post9t + e9ts / (e9tk + e9ts);

                                    e9ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e9tk > 0)
                                        post9tm = post9tm + e9ts / (e9tk + e9ts);

                                }
                            }
                        }

                        if (kupac.Contains("ROMANIA"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (turm.Contains("Da"))
                                {
                                    e10tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e10ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e10ko = e10ts;
                                    if (e10tk > 0)
                                        post10t = post10t + e10ts / (e10tk + e10ts);

                                    e10ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e10tk > 0)
                                        post10tm = post10tm + e10ts / (e10tk + e10ts);

                                }
                                else
                                {
                                    e10tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e10ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e10ko = e10ts;
                                    if (e10tk > 0)
                                        post10t = post10t + e10ts / (e10tk + e10ts);

                                    e10ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e10tk > 0)
                                        post10tm = post10tm + e10ts / (e10tk + e10ts);

                                }
                            }

                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (kupac.Contains("Prsteni"))  // wupertal
                                {
                                    if (turm.Contains("Da"))
                                    {
                                        if (reader["kolicinaok"] != DBNull.Value)
                                        {
                                            e11tk = (double.Parse)(reader["kolicinaok"].ToString());
                                            e11ts = (double.Parse)(reader["otpadobrada"].ToString());
                                            e11ko = e11ts;
                                            if (e11tk > 0)
                                                post11t = post11t + e11ts / (e11tk + e11ts);

                                            e11ts = (double.Parse)(reader["otpadmat"].ToString());
                                            if (e11tk > 0)
                                                post11tm = post11tm + e11ts / (e11tk + e11ts);
                                        }

                                    }
                                    else
                                    {
                                        if (reader["kolicinaok"] != DBNull.Value)
                                        {
                                            e11tk = (double.Parse)(reader["kolicinaok"].ToString());
                                            e11ts = (double.Parse)(reader["otpadobrada"].ToString());
                                            e11ko = e11ts;
                                            if (e11tk > 0)
                                                post11t = post11t + e11ts / (e11tk + e11ts);

                                            e11ts = (double.Parse)(reader["otpadmat"].ToString());
                                            if (e11tk > 0)
                                                post11tm = post11tm + e11ts / (e11tk + e11ts);
                                        }

                                    }

                                }
                            }
                        }

                        if (kupac.Contains("Technologies"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (kupac.Contains("Valjci"))  // wupertal
                                {
                                    if (turm.Contains("Da"))
                                    {
                                        e12tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e12ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e12ko = e12ts;
                                        if (e12tk > 0)
                                            post12t = post12t + e12ts / (e12tk + e12ts);

                                        e12ts = (double.Parse)(reader["otpadmat"].ToString());
                                        post12tm = post12tm + e12ts / (e12tk + e12ts);
                                        if ((e12tk + e12ts) <= 0.001)
                                            post12tm = 0.0;

                                    }
                                    else
                                    {
                                        e12tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e12ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e12ko = e12ts;
                                        if (e12tk > 0)
                                            post12t = post12t + e12ts / (e12tk + e12ts);

                                        e12ts = (double.Parse)(reader["otpadmat"].ToString());
                                        post12tm = post12tm + e12ts / (e12tk + e12ts);

                                        if ((e12tk + e12ts) <= 0.001)
                                            post12tm = 0.0;

                                    }

                                }
                            }
                        }


                        if (kupac.Contains("KYSUCE"))  // 
                        {
                            if (ws == 4)
                            {
                                int jk = 0;
                            }
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (kupac.Contains("Ku"))  // 
                                {
                                    if (turm.Contains("Da"))
                                    {
                                        e13tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e13ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e13ko = e13ts;
                                        if (e13tk > 0)
                                            post13t = post13t + e13ts / (e13tk + e13ts);

                                        e13ts = (double.Parse)(reader["otpadmat"].ToString());
                                        if (e13tk > 0)
                                            post13tm = post13tm + e13ts / (e13tk + e13ts);

                                    }
                                    else
                                    {
                                        e13tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e13ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e13ko = e13ts;
                                        if (e13tk > 0)
                                            post13t = post13t + e13ts / (e13tk + e13ts);

                                        e13ts = (double.Parse)(reader["otpadmat"].ToString());
                                        if (e13tk > 0)
                                            post13tm = post13tm + e13ts / (e13tk + e13ts);

                                    }
                                }
                            }
                        }

                        if (kupac.Contains("KYSUCE"))  //
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                if (kupac.Contains("Prsten"))  // 
                                {
                                    if (turm.Contains("Da"))
                                    {
                                        e17tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e17ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e17ko = e17ts;
                                        if (e17tk > 0)
                                            post17t = post17t + e17ts / (e17tk + e17ts);

                                        e17ts = (double.Parse)(reader["otpadmat"].ToString());
                                        if (e17tk > 0)
                                            post17tm = post17tm + e17ts / (e17tk + e17ts);

                                    }
                                    else
                                    {
                                        e17tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e17ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e17ko = e17ts;
                                        if (e17tk > 0)
                                            post17t = post17t + e17ts / (e17tk + e17ts);

                                        e17ts = (double.Parse)(reader["otpadmat"].ToString());
                                        if (e17tk > 0)
                                            post17tm = post17tm + e17ts / (e17tk + e17ts);

                                    }
                                }
                            }

                        }

                        if (kupac.Contains("FERRO"))  // wupertal
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (turm.Contains("Da"))
                                {
                                    e14tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e14ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e14ko = e14ts;
                                    if (e14tk > 0)
                                        post14t = post14t + e14ts / (e14tk + e14ts);

                                    e14ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e14tk > 0)
                                        post14tm = post14tm + e14ts / (e14tk + e14ts);

                                }
                                else
                                {
                                    e14tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e14ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e14ko = e14ts;
                                    if (e14tk > 0)
                                        post14t = post14t + e14ts / (e14tk + e14ts);

                                    e14ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e14tk > 0)
                                        post14tm = post14tm + e14ts / (e14tk + e14ts);

                                }
                            }

                        }

                        if (kupac.Contains("SONA BLW"))
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {
                                if (idpar == "121301")
                                {
                                    if (turm.Contains("Da"))
                                    {
                                        e15tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e15ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e15ko = e15ts;
                                        if (e15tk > 0)
                                            post15t = post15t + e15ts / (e15tk + e15ts);

                                        e15ts = (double.Parse)(reader["otpadmat"].ToString());
                                        if (e15tk > 0)
                                            post15tm = post15tm + e15ts / (e15tk + e15ts);

                                    }
                                    else
                                    {
                                        e15tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e15ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e15ko = e15ts;
                                        if (e15tk > 0)
                                            post15t = post15t + e15ts / (e15tk + e15ts);

                                        e15ts = (double.Parse)(reader["otpadmat"].ToString());
                                        if (e15tk > 0)
                                            post15tm = post15tm + e15ts / (e15tk + e15ts);

                                    }
                                }
                                else if (idpar == "121302")
                                {                                    
                                        if (turm.Contains("Da"))
                                        {
                                            e151tk = (double.Parse)(reader["kolicinaok"].ToString());
                                            e151ts = (double.Parse)(reader["otpadobrada"].ToString());
                                            e151ko = e151ts;
                                            if (e151tk > 0)
                                                post151t = post151t + e151ts / (e151tk + e151ts);

                                            e151ts = (double.Parse)(reader["otpadmat"].ToString());
                                            if (e151tk > 0)
                                                post151tm = post151tm + e151ts / (e151tk + e151ts);

                                        }
                                        else
                                        {
                                            e151tk = (double.Parse)(reader["kolicinaok"].ToString());
                                            e151ts = (double.Parse)(reader["otpadobrada"].ToString());
                                            e151ko = e151ts;
                                            if (e151tk > 0)
                                                post151t = post151t + e151ts / (e151tk + e151ts);

                                            e151ts = (double.Parse)(reader["otpadmat"].ToString());
                                            if (e151tk > 0)
                                                post151tm = post151tm + e151ts / (e151tk + e151ts);

                                        }

                                    
                                }

                            }
                        }

                                                

                        if (kupac.Contains("SIGMA"))  // 
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (turm.Contains("Da"))
                                {
                                    e16tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e16ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e16ko = e16ts;
                                    if (e16tk > 0)
                                        post16t = post16t + e16ts / (e16tk + e16ts);

                                    e16ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e16tk > 0)
                                        post16tm = post16tm + e16ts / (e16tk + e16ts);

                                }
                                else
                                {
                                    e16tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e16ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e16ko = e16ts;
                                    if (e16tk > 0)
                                        post16t = post16t + e16ts / (e16tk + e16ts);

                                    e16ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e16tk > 0)
                                        post16tm = post16tm + e16ts / (e16tk + e16ts);

                                }
                            }

                        }

                        if (kupac.Contains("Brasil"))  // 
                        {
                            if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                            {

                                if (turm.Contains("Da"))
                                {
                                    e19tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e19ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e19ko = e19ts;
                                    if (e19tk > 0)
                                        post19t = post19t + e19ts / (e19tk + e19ts);

                                    e19ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e19tk > 0)
                                        post19tm = post19tm + e19ts / (e19tk + e19ts);

                                }
                                else
                                {
                                    e19tk = (double.Parse)(reader["kolicinaok"].ToString());
                                    e19ts = (double.Parse)(reader["otpadobrada"].ToString());
                                    e19ko = e19ts;
                                    if (e19tk > 0)
                                        post19t = post19t + e19ts / (e19tk + e19ts);

                                    e19ts = (double.Parse)(reader["otpadmat"].ToString());
                                    if (e19tk > 0)
                                        post19tm = post19tm + e19ts / (e19tk + e19ts);


                                }
                            }

                            if (kupac.Contains("NEU"))  //  NSK
                            {
                                if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") || (ws <= 2 && hala1 == "A"))
                                {

                                    if (turm.Contains("Da"))
                                    {
                                        e21tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e21ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e21ko = e21ts;
                                        if (e21tk > 0)
                                            post21t = post21t + e21ts / (e21tk + e21ts);

                                        e21ts = (double.Parse)(reader["otpadmat"].ToString());
                                        if (e21tk > 0)
                                            post21tm = post21tm + e21ts / (e21tk + e21ts);

                                    }
                                    else
                                    {
                                        e21tk = (double.Parse)(reader["kolicinaok"].ToString());
                                        e21ts = (double.Parse)(reader["otpadobrada"].ToString());
                                        e21ko = e21ts;
                                        if (e21tk > 0)
                                            post21t = post21t + e21ts / (e21tk + e21ts);

                                        e21ts = (double.Parse)(reader["otpadmat"].ToString());
                                        if (e21tk > 0)
                                            post21tm = post21tm + e21ts / (e21tk + e21ts);

                                    }
                                }

                            }
                        }

                    }
                    cn.Close();

                }
                

                /////////////
                // kaliona
                double b31 = 0.0, e31 = 0.0;
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws == 1)
                    {
                        sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',4," + smj; ;  // smjenski izvještaj
                    }
                    else
                    {
                        sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',40," + smj;   // tjedni izvještaj
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;

                    while (reader.Read())
                    {
                        pec = reader["pec"].ToString();
                        if (reader["pec"] != DBNull.Value)
                        {
                            if (pec.Contains("CODERE"))
                            {
                                b31 = (double.Parse)(reader["tezina"].ToString());
                            }
                            else
                            {
                                e31 = (double.Parse)(reader["tezina"].ToString());
                            }
                        }
                    }
                    cn.Close();
                }
                // broj šarži
                int g31 = 0;
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    if (ws == 1)
                    {
                        sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',41," + smj;
                    }
                    else
                    {
                        sql1 = "rfind.dbo.realizacija2 '" + dat1 + "','" + dat2 + "',410," + smj;
                    }
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;

                    while (reader.Read())
                    {

                        g31 = g31 + (int.Parse)(reader["brojsarzi"].ToString());

                    }
                    cn.Close();
                }

                Console.WriteLine("Izračunat broj šarži  trenutno vrijeme " + DateTime.Now);
                if (ws == 1)
                {

                }
                Worksheet worksheet = workbook.Worksheets.Item[ws] as Worksheet;
                if (worksheet == null)
                    return;

                // planiranje share
                if (ws == 1 || ws == 3 || ws == 4)
                {
                    worksheet.Rows.Cells[1, 1].value = worksheet.Rows.Cells[1, 1].value + " " + dat10 + "     " + smjenanaziv;
                }
                else if (ws == 2)
                {
                    worksheet.Rows.Cells[1, 1].value = worksheet.Rows.Cells[1, 1].value + " " + dat10;
                }

                worksheet.Rows.Cells[7, 2].value = b7;  // planirano=norma, samo tokarenjem
                worksheet.Rows.Cells[8, 2].value = b8;
                worksheet.Rows.Cells[9, 2].value = b9;
                worksheet.Rows.Cells[10, 2].value = b10;
                worksheet.Rows.Cells[11, 2].value = b11;
                worksheet.Rows.Cells[12, 2].value = b12;
                worksheet.Rows.Cells[13, 2].value = b13;
                worksheet.Rows.Cells[14, 2].value = b14;
                worksheet.Rows.Cells[15, 2].value = b15;     // sona blw (m)
                worksheet.Rows.Cells[16, 2].value = b151;    // sona auto (r)
                worksheet.Rows.Cells[17, 2].value = b16;
                worksheet.Rows.Cells[18, 2].value = b17;
                worksheet.Rows.Cells[20, 2].value = b19;
                worksheet.Rows.Cells[22, 2].value = b21;

                worksheet.Rows.Cells[7, 4].value = d70;  // planirano=norma  dodatna obrada
                worksheet.Rows.Cells[8, 4].value = d80;
                worksheet.Rows.Cells[9, 4].value = d90;
                worksheet.Rows.Cells[10, 4].value = d100;
                worksheet.Rows.Cells[11, 4].value = d110;
                worksheet.Rows.Cells[12, 4].value = d120;
                worksheet.Rows.Cells[13, 4].value = d130;
                worksheet.Rows.Cells[14, 4].value = d140;
                worksheet.Rows.Cells[15, 4].value = d150;
                worksheet.Rows.Cells[16, 4].value = d1510;
                worksheet.Rows.Cells[17, 4].value = d160;
                worksheet.Rows.Cells[18, 4].value = d170;
                worksheet.Rows.Cells[20, 4].value = d190;
                worksheet.Rows.Cells[22, 4].value = d210;



                //          worksheet.Rows.Cells[26, 2].value = worksheet.Rows.Cells[22, 2].value;


                worksheet.Rows.Cells[7, 3].value = c7;  // količina
                worksheet.Rows.Cells[8, 3].value = c8;
                worksheet.Rows.Cells[9, 3].value = c9;
                worksheet.Rows.Cells[10, 3].value = c10;
                worksheet.Rows.Cells[11, 3].value = c11;
                worksheet.Rows.Cells[12, 3].value = c12;
                worksheet.Rows.Cells[13, 3].value = c13;
                worksheet.Rows.Cells[14, 3].value = c14;
                worksheet.Rows.Cells[15, 3].value = c15;
                worksheet.Rows.Cells[16, 3].value = c151;
                worksheet.Rows.Cells[17, 3].value = c16;
                worksheet.Rows.Cells[18, 3].value = c17;
                worksheet.Rows.Cells[20, 3].value = c19;
                worksheet.Rows.Cells[22, 3].value = c21;

                worksheet.Rows.Cells[7, 5].value = c7do;  // količina obrađena dodatnom obradom
                worksheet.Rows.Cells[8, 5].value = c8do;
                worksheet.Rows.Cells[9, 5].value = c9do;
                worksheet.Rows.Cells[10, 5].value = c10do;
                worksheet.Rows.Cells[11, 5].value = c11do;
                worksheet.Rows.Cells[12, 5].value = c12do;
                worksheet.Rows.Cells[13, 5].value = c13do;
                worksheet.Rows.Cells[14, 5].value = c14do;
                worksheet.Rows.Cells[15, 5].value = c15do;     // sona m
                worksheet.Rows.Cells[16, 5].value = c151do;   // sona r
                worksheet.Rows.Cells[17, 5].value = c16do;
                worksheet.Rows.Cells[18, 5].value = c17do;
                worksheet.Rows.Cells[20, 5].value = c19do;
                worksheet.Rows.Cells[22, 5].value = c21do;

                if (ws <= 4)
                {
                    worksheet.Rows.Cells[7, 6].value = f7;  // vrijednost obrade tokarenje
                    worksheet.Rows.Cells[8, 6].value = f8;
                    worksheet.Rows.Cells[9, 6].value = f9;
                    worksheet.Rows.Cells[10, 6].value = f10;
                    worksheet.Rows.Cells[11, 6].value = f11;
                    worksheet.Rows.Cells[12, 6].value = f12;
                    worksheet.Rows.Cells[13, 6].value = f13;
                    worksheet.Rows.Cells[14, 6].value = f14;
                    worksheet.Rows.Cells[15, 6].value = f15;
                    worksheet.Rows.Cells[16, 6].value = f151;
                    worksheet.Rows.Cells[17, 6].value = f16;
                    worksheet.Rows.Cells[18, 6].value = f17;
                    worksheet.Rows.Cells[20, 6].value = f19;
                    worksheet.Rows.Cells[22, 6].value = f21;

                    worksheet.Rows.Cells[7, 7].value = g7;  // vrijednost obrade, dodatne operacije
                    worksheet.Rows.Cells[8, 7].value = g8;
                    worksheet.Rows.Cells[9, 7].value = g9;
                    worksheet.Rows.Cells[10, 7].value = g10;
                    worksheet.Rows.Cells[11, 7].value = g11;
                    worksheet.Rows.Cells[12, 7].value = g12;
                    worksheet.Rows.Cells[13, 7].value = g13;
                    worksheet.Rows.Cells[14, 7].value = g14;
                    worksheet.Rows.Cells[15, 7].value = g15;
                    worksheet.Rows.Cells[16, 7].value = g151;
                    worksheet.Rows.Cells[17, 7].value = g16;
                    worksheet.Rows.Cells[18, 7].value = g17;
                    worksheet.Rows.Cells[20, 7].value = g19;
                    worksheet.Rows.Cells[22, 7].value = g21;
                }
                if (ws <= 4)
                {

                    worksheet.Rows.Cells[7, 15].value = j7;   // broj štelanja
                    worksheet.Rows.Cells[8, 15].value = j8;
                    worksheet.Rows.Cells[9, 15].value = j9;
                    worksheet.Rows.Cells[10, 15].value = j10;
                    worksheet.Rows.Cells[11, 15].value = j11;
                    worksheet.Rows.Cells[12, 15].value = j12;
                    worksheet.Rows.Cells[13, 15].value = j13;
                    worksheet.Rows.Cells[14, 15].value = j14;
                    worksheet.Rows.Cells[15, 15].value = j15;
                    worksheet.Rows.Cells[16, 15].value = j151;
                    worksheet.Rows.Cells[17, 15].value = j16;
                    worksheet.Rows.Cells[18, 15].value = j17;
                    worksheet.Rows.Cells[20, 15].value = j19;
                    worksheet.Rows.Cells[22, 3].value  = c21;

                    worksheet.Rows.Cells[7, 13].value = i7;   // broj linija
                    worksheet.Rows.Cells[8, 13].value = i8;
                    worksheet.Rows.Cells[9, 13].value = i9;
                    worksheet.Rows.Cells[10, 13].value = i10;
                    worksheet.Rows.Cells[11, 13].value = i11;
                    worksheet.Rows.Cells[12, 13].value = i12;
                    worksheet.Rows.Cells[13, 13].value = i13;
                    worksheet.Rows.Cells[14, 13].value = i14;
                    worksheet.Rows.Cells[15, 13].value = i15;
                    worksheet.Rows.Cells[16, 13].value = i151;

                    worksheet.Rows.Cells[17, 13].value = i16;
                    worksheet.Rows.Cells[18, 13].value = i17;
                    worksheet.Rows.Cells[19, 13].value = i18;
                    worksheet.Rows.Cells[20, 13].value = i19;
                    worksheet.Rows.Cells[22, 13].value = i21;
                    worksheet.Rows.Cells[23, 13].value = i22 - worksheet.Rows.Cells[25, 13].value;

                    worksheet.Rows.Cells[7, 14].value = n7;   // broj linija koje ne rade
                    worksheet.Rows.Cells[8, 14].value = n8;
                    worksheet.Rows.Cells[9, 14].value = n9;
                    worksheet.Rows.Cells[10, 14].value = n10;
                    worksheet.Rows.Cells[11, 14].value = n11;
                    worksheet.Rows.Cells[12, 14].value = n12;
                    worksheet.Rows.Cells[13, 14].value = n13;
                    worksheet.Rows.Cells[14, 14].value = n14;
                    worksheet.Rows.Cells[15, 14].value = n15;

                    worksheet.Rows.Cells[16, 14].value = n151; 
                    
                    worksheet.Rows.Cells[17, 14].value = n16;
                    worksheet.Rows.Cells[18, 14].value = n17;
                    worksheet.Rows.Cells[20, 14].value = n19;
                    worksheet.Rows.Cells[22, 14].value = n21;
                    worksheet.Rows.Cells[23, 14].value = worksheet.Rows.Cells[23, 13].value;


                    worksheet.Rows.Cells[7, 9].value = Math.Round(post7t, 2);   // škart obrade
                    worksheet.Rows.Cells[8, 9].value = Math.Round(post8t, 2);
                    worksheet.Rows.Cells[9, 9].value = Math.Round(post9t, 2);
                    worksheet.Rows.Cells[10, 9].value = Math.Round(post10t, 2);
                    worksheet.Rows.Cells[11, 9].value = Math.Round(post11t, 2);
                    worksheet.Rows.Cells[12, 9].value = Math.Round(post12t, 2);
                    worksheet.Rows.Cells[13, 9].value = Math.Round(post13t, 2);   // kysuce
                    worksheet.Rows.Cells[14, 9].value = Math.Round(post14t, 2);
                    worksheet.Rows.Cells[15, 9].value = Math.Round(post15t, 2);     // sona m
                    worksheet.Rows.Cells[16, 9].value = Math.Round(post151t, 2);     // sona¸r
                    worksheet.Rows.Cells[17, 9].value = Math.Round(post16t, 2);
                    worksheet.Rows.Cells[18, 9].value = Math.Round(post17t, 2);
                    worksheet.Rows.Cells[20, 9].value = Math.Round(post19t, 2);

                    worksheet.Rows.Cells[7, 10].value = e7ko;   // škart obrade komada
                    worksheet.Rows.Cells[8, 10].value = e8ko;
                    worksheet.Rows.Cells[9, 10].value = e9ko;
                    worksheet.Rows.Cells[10, 10].value = e10ko;
                    worksheet.Rows.Cells[11, 10].value = e11ko;
                    worksheet.Rows.Cells[12, 10].value = e12ko;
                    worksheet.Rows.Cells[13, 10].value = e13ko;    // kysuce
                    worksheet.Rows.Cells[14, 10].value = e14ko;
                    worksheet.Rows.Cells[15, 10].value = e15ko;      // sona m
                    worksheet.Rows.Cells[16, 10].value = e151ko;      // sona r
                    worksheet.Rows.Cells[17, 10].value = e16ko;
                    worksheet.Rows.Cells[18, 10].value = e17ko;
                    worksheet.Rows.Cells[20, 10].value = e19ko;

                    worksheet.Rows.Cells[7, 11].value = Math.Round(post7tm, 2);   // škart materijala
                    worksheet.Rows.Cells[8, 11].value = Math.Round(post8tm, 2);
                    worksheet.Rows.Cells[9, 11].value = Math.Round(post9tm, 2);
                    worksheet.Rows.Cells[10, 11].value = Math.Round(post10tm, 2);
                    worksheet.Rows.Cells[11, 11].value = Math.Round(post11tm, 2);
                    worksheet.Rows.Cells[12, 11].value = Math.Round(post12tm, 2);
                    worksheet.Rows.Cells[13, 11].value = Math.Round(post13tm, 2);
                    worksheet.Rows.Cells[14, 11].value = Math.Round(post14tm, 2);
                    worksheet.Rows.Cells[15, 11].value = Math.Round(post15tm, 2);    // sona m
                    worksheet.Rows.Cells[16, 11].value = Math.Round(post151tm, 2);   // sona r
                    worksheet.Rows.Cells[17, 11].value = Math.Round(post16tm, 2);
                    worksheet.Rows.Cells[18, 11].value = Math.Round(post17tm, 2);
                    worksheet.Rows.Cells[20, 11].value = Math.Round(post19tm, 2);

                    worksheet.Rows.Cells[7, 12].value = e7ts;   // škart materijala komada
                    worksheet.Rows.Cells[8, 12].value = e8ts;
                    worksheet.Rows.Cells[9, 12].value = e9ts;
                    worksheet.Rows.Cells[10, 12].value = e10ts;
                    worksheet.Rows.Cells[11, 12].value = e11ts;
                    worksheet.Rows.Cells[12, 12].value = e12ts;
                    worksheet.Rows.Cells[13, 12].value = e13ts;    // kysuce
                    worksheet.Rows.Cells[14, 12].value = e14ts;
                    worksheet.Rows.Cells[15, 12].value = e15ts;      // sona m
                    worksheet.Rows.Cells[16, 12].value = e151ts;      // sona r
                    worksheet.Rows.Cells[17, 12].value = e16ts;
                    worksheet.Rows.Cells[18, 12].value = e17ts;
                    worksheet.Rows.Cells[20, 12].value = e19ts;
                    worksheet.Rows.Cells[22, 12].value = e21ts;


                    if (ws <= 2)
                    {
                        //            worksheet.Rows.Cells[26, 2].value = b26;    // Realizacija ukupno: planirano
                        worksheet.Rows.Cells[29, 2].value = worksheet.Rows.Cells[25, 2].value + worksheet.Rows.Cells[25, 4].value; //
                        worksheet.Rows.Cells[29, 3].value = worksheet.Rows.Cells[25, 3].value + worksheet.Rows.Cells[25, 5].value; //-- ukupk;    // ukupna količina
                        worksheet.Rows.Cells[30, 3].value = worksheet.Rows.Cells[25, 6].value + worksheet.Rows.Cells[25, 7].value; //-- ukupk;    // ukupna vrijednost
                                                                                                                                   //           worksheet.Rows.Cells[27, 2].value = b27;    // Realizacija ukupno: realizirano
                                                                                                                                   //worksheet.Rows.Cells[27, 3].value = c27;    // ukupna vrijednost

                        worksheet.Rows.Cells[34, 2].value = b31;    // kaliona
                        worksheet.Rows.Cells[34, 5].value = e31;
                        worksheet.Rows.Cells[34, 7].value = g31;   // broj šarži
                    }
                    //worksheet.Rows.Cells[31, 9].value = (e31 + b31) * 0.33;

                    //worksheet.Rows.Cells[26, 6].value = ul24;    // Transporti 24 t
                    //worksheet.Rows.Cells[26, 7].value = iz24;
                    //worksheet.Rows.Cells[27, 6].value = ul15;    // Transporti 1.5T
                    //worksheet.Rows.Cells[27, 7].value = iz15;

                }
                Console.WriteLine("Napunjen dprsddmmyyyy   trenutno vrijeme " + DateTime.Now);
                Worksheet worksheet4 = workbook.Worksheets.Item[6] as Worksheet;  // po linijama
                if (ws == 1)  // na kraju popuni ldp radnika,linija
                {

                    Worksheet worksheet3 = workbook.Worksheets.Item[5] as Worksheet;  // djelatnici

                    if (worksheet == null)
                        return;

                    // LDP
                    worksheet3.Rows.Cells[1, 2].value = worksheet3.Rows.Cells[1, 2].value + " " + datreps;
                    worksheet3.Rows.Cells[2, 6].value = smjenanaziv;

                    worksheet4.Rows.Cells[1, 3].value = worksheet4.Rows.Cells[1, 3].value + " " + datreps;
                    worksheet4.Rows.Cells[2, 6].value = smjenanaziv;
                    double ostvareno = 0; double planirano = 0; double uostvareno = 0.0;

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        //sql1 = "rfind.dbo.ldp_recalc '" + dat13 + "','" + dat23 + "'," + smj; ;
                        //SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        //SqlDataReader reader = sqlCommand.ExecuteReader();
                        //string pec;
                        int i = 0;
                        int ukupn = 0, ukupp = 0, ukupkol = 0;
                        int k1 = 0, n1 = 0;
                        double minuta = 0.0, p1 = 0.0, ukupl = 0.0;
                        sql1 = "select * from rfind.dbo.evidnormiradad('" + datLDP + "','" + datLDP + "')  where smjena=" + smj + " order by vrsta1,vrsta,radnik";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string pec;


                        while (reader.Read())
                        {
                            string sk1 = reader["kolicinaok"].ToString();
                            minuta = 0.0;


                            if (reader["kolicinaok"] == DBNull.Value)
                            {
                                k1 = 0;
                            }
                            else
                            {
                                k1 = (int.Parse)(reader["kolicinaok"].ToString());
                            }

                            if (reader["norma"] == DBNull.Value)
                            {
                                n1 = 0;
                                minuta = (double.Parse)((reader["minutaradaradnika"].ToString()).Replace(',', '.')) / 100;
                                p1 = 0;
                            }
                            else
                            {
                                n1 = (int.Parse)(reader["norma"].ToString());
                                if (reader["minutaradaradnika"] == DBNull.Value)
                                {
                                    minuta = 0.0;
                                }
                                else
                                {
                                    minuta = (double.Parse)((reader["minutaradaradnika"].ToString()).Replace(',', '.')) / 100;
                                }

                                if (minuta != 480)
                                    p1 = n1 * minuta / (480 + 0.001);
                                else
                                    p1 = n1;

                            }

                            double post1 = 0.0;
                            if (k1 > 0)
                            {
                                post1 = (k1 + 0.00001) / (p1 + 0.0001);
                            }
                            else
                            {
                                post1 = 0.0;
                            }

                            worksheet3.Rows.Cells[5 + i, 1].value = reader["radnik"];
                            worksheet3.Rows.Cells[5 + i, 2].value = (reader["vrsta"].ToString());
                            worksheet3.Rows.Cells[5 + i, 3].value = (reader["hala"].ToString());
                            worksheet3.Rows.Cells[5 + i, 4].value = (reader["smjena"].ToString());
                            worksheet3.Rows.Cells[5 + i, 5].value = (reader["linija"].ToString());
                            worksheet3.Rows.Cells[5 + i, 6].value = reader["nazivpar"];
                            worksheet3.Rows.Cells[5 + i, 7].value = reader["brojrn"];
                            worksheet3.Rows.Cells[5 + i, 8].value = (reader["proizvod"].ToString());
                            worksheet3.Rows.Cells[5 + i, 9].value = (reader["norma"].ToString());
                            worksheet3.Rows.Cells[5 + i, 11].value = (reader["kolicinaok"].ToString());

                            worksheet3.Rows.Cells[5 + i, 10].value = (int)p1;
                            worksheet3.Rows.Cells[5 + i, 12].value = post1;
                            worksheet3.Rows.Cells[5 + i, 13].value = (reader["otpadobrada"].ToString());
                            worksheet3.Rows.Cells[5 + i, 14].value = (reader["kolicinaporozno"].ToString());
                            worksheet3.Rows.Cells[5 + i, 15].value = (reader["otpadmat"].ToString());

                            worksheet3.Rows.Cells[5 + i, 16].value = minuta;
                            worksheet3.Rows.Cells[5 + i, 19].value = (reader["napomena1"].ToString());
                            worksheet3.Rows.Cells[5 + i, 20].value = (reader["napomena2"].ToString());
                            worksheet3.Rows.Cells[5 + i, 21].value = (reader["napomena3"].ToString());

                            string smjena2     = (reader["smjena"].ToString());
                            string hala2       = (reader["hala"].ToString());
                            string linija2     = (reader["linija"].ToString());
                            string radninalog2 = (reader["brojrn"].ToString());
                            int trajanje11 = 0; int trajanje_orad = 0;
                            string krajj ="" ,napomena_or="";

                            /// begin aktivnosti
                            using (SqlConnection cna = new SqlConnection(connectionString))
                            {
                                cna.Open();
                                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                                dat13 = datLDP + " 06:00:00";
                                string dat33 = dat3 + " 05:59:00";

                                if (smjena2 == "1")
                                {
                                    dat13 = datLDP + " 06:00:00";
                                    dat33 = datLDP + " 13:59:00";
                                }
                                if (smjena2 == "2")
                                {
                                    dat13 = datLDP + " 14:00:00";
                                    dat33 = datLDP + " 21:59:00";
                                }
                                if (smjena2 == "3")
                                {
                                    dat13 = datLDP + " 22:00:00";
                                    dat33 = dats + " 05:59:00";
                                }

                                string sqla = "select * from rfind.dbo.pregled_po_liniji( '" + dat13 + "','" + dat33 + "','" + hala2 + "','" + linija2.TrimEnd() + "')";
                                if (linija2.TrimEnd()=="07")
                                {
                                    int l = 0;
                                }
                                SqlCommand sqlCommanda = new SqlCommand(sqla, cna);
                                SqlDataReader readera = sqlCommanda.ExecuteReader();
                                string aktivnost1 = "";
                                string trajanje1 = "";
                                string napomena1 = "", rna1 = "";
                             

                                while (readera.Read())
                                {
                                    if (linija2 == "16")
                                    {
                                        int tt211 = 1;
                                    }

                                    aktivnost1 = readera["aktivnost"].ToString();
                                    trajanje1 = readera["trajanje"].ToString();
                                    krajj = readera["kraj"].ToString().TrimEnd();
                                    napomena1 = readera["napomena"].ToString();

                                    if (krajj == "")
                                        napomena1 = "U toku:" + napomena1;

                                    
                                    rna1 = readera["brojrn"].ToString().TrimEnd();
                                    if (rna1 != "")
                                    {
                                        if (rna1 != radninalog2)
                                        {
//                                            continue;
                                        }
                                    }

                                    if (aktivnost1.Contains("Otežan"))   // Otežan rad
                                    {
                                        //worksheet3.Rows.Cells[5 + i, 23].value = napomena1 + " ( RN: " + rna1 + " )";

                                        if (napomena1.Contains("RBT1"))                 // rad bez transportera2  10 minuta u 8 sati
                                        {
                                            trajanje_orad = (int.Parse)(trajanje1) / 480 * 10 ;
                                            napomena_or =  " - "+trajanje_orad.ToString();
                                        }
                                        if (napomena1.Contains("RBT2"))                 // rad bez transportera2  15 minuta
                                        {
                                            trajanje_orad = (int.Parse)(trajanje1) / 480 * 15;
                                            napomena_or = " - " + trajanje_orad.ToString();
                                        }
                                        if (napomena1.Contains("KONZ+"))            //konzerviranje
                                        {
                                            trajanje_orad = (int.Parse)(trajanje1) / 480 * 10;
                                            napomena_or = " - " + trajanje_orad.ToString();
                                        }
                                        if (napomena1.Contains("1KANAL+"))                 // RAD NA JEDNOM KANALU
                                        {
                                            trajanje_orad = (int.Parse)(trajanje1) / 480 * 200;
                                            napomena_or = " - " + trajanje_orad.ToString();
                                        }

                                        if (trajanje_orad>0)
                                            worksheet3.Rows.Cells[5 + i, 22].value = trajanje1.ToString() + " ( " + napomena_or+" )";
                                        else
                                            worksheet3.Rows.Cells[5 + i, 22].value = trajanje1.ToString() ;

                                        worksheet3.Rows.Cells[5 + i, 23].value = napomena1 + " ( RN: " + rna1 + " )";
                                    }


                                    if (aktivnost1.Contains("djelatnika"))  // Nedostatak djelatnika
                                    {
                                        worksheet3.Rows.Cells[5 + i, 24].value = trajanje1;
                                        worksheet3.Rows.Cells[5 + i, 25].value = napomena1 + " ( RN: " + rna1 + " )";

                                    }
                                    if (aktivnost1.Contains("sirovca"))    // Nedostatak sirovca
                                    {
                                        worksheet3.Rows.Cells[5 + i, 26].value = trajanje1;
                                        worksheet3.Rows.Cells[5 + i, 27].value = napomena1 + " ( RN: " + rna1 + " )";

                                    }
                                    if (aktivnost1.Contains("Dorada"))    // Dorada
                                    {
                                        worksheet3.Rows.Cells[5 + i, 28].value = trajanje1;
                                        worksheet3.Rows.Cells[5 + i, 29].value = napomena1 + " ( RN: " + rna1 + " )";

                                    }

                                }
                            }

                            ///111

                            /// end aktivnosti 

                            p1 = n1 * (minuta -  trajanje_orad) / 480;


                            //int n1 = (int.Parse)(reader["norma"].ToString());
                            double posto1 = 0.00;
                            if (n1 > 0)
                            {

                                if (k1 == p1)
                                {
                                    posto1 = 1;
                                }
                                else
                                {
                                    posto1 = k1 / (p1 + 0.0000001);
                                }

                                if (posto1 >= 1.00)
                                {
                                    worksheet3.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbGreenYellow;
                                }

                                if (posto1 >= 0.900 && posto1 < 1.00)
                                {
                                    worksheet3.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbOrange;
                                }

                                if (posto1 >= 0.99 && posto1 < 1.00 && k1 > 0)
                                {
                                    //worksheet3.Rows.Cells[5 + i, 11].Interior.Color = XlRgbColor.rgbYellowGreen;
                                }
                                if (posto1 < 0.900)
                                {
                                    worksheet3.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbRed;
                                }

                            }
                                                       
                            ukupn = ukupn + n1;
                            ukupl = ukupl + p1;
                            if (reader["norma"] != DBNull.Value)
                            {
                                ukupp = ukupp + n1;
                                ukupkol = ukupkol + k1;
                            }

                            i++;

                        }
                        worksheet3.Rows.Cells[5 + i, 8].value = "Ukupno :";
                        worksheet3.Rows.Cells[5 + i, 9].value = ukupn;
                        worksheet3.Rows.Cells[5 + i, 11].value = ukupkol;
                        worksheet3.Rows.Cells[5 + i, 10].value = (int)ukupl;
                        worksheet3.Rows.Cells[6 + i, 8].value = "Postotak realizacije";

                        worksheet3.Rows.Cells[6 + i, 9].value = Math.Round(((ukupkol + 0.001) / (ukupp + 0.001) * 100), 2);
                        worksheet3.Rows.Cells[6 + i, 10].value = Math.Round(((ukupkol + 0.001) / (ukupl + 0.001) * 100), 2);

                        cn.Close();
                    }
                }
                // sumarno izostanci ws=1
                if ((ws == 1) || (ws == 3) || (ws == 4))
                {
                    using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        if (ws == 1)
                            sql1 = "rfind.dbo.realizacija2 '" + datLDP + "','" + datLDP + "',72," + smj;
                        else              // izostanci po mt,smjenama i pogonima
                            sql1 = "rfind.dbo.realizacija2 '" + datLDP + "','" + datLDP + "',720," + smj;

                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        //worksizostanci.Rows.Cells[4, 3].value = worksizostanci.Rows.Cells[4, 3].value + " " + datreps;
                        string vrstao = "", mt = "", hala1 = "";
                        int broj = 0;

                        int i = 0;
                        while (reader.Read())
                        {

                            mt = reader["mtroska"].ToString();
                            if (ws >= 3)
                            {
                                hala1 = reader["hala"].ToString().Trim();
                                if (((ws == 3) && (hala1 == "1")) || ((ws == 4) && (hala1 == "3")))
                                {
                                }
                                else
                                {
                                    continue;
                                }

                            }

                            broj = (int.Parse)(reader["broj"].ToString());
                            vrstao = (reader["vrsta"].ToString());



                            if (mt.Contains("Tokarenje"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[7, 19].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[7, 20].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[7, 21].value = broj;
                                }

                            }
                            if (mt.Contains("Kaljenje"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[8, 19].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[8, 20].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[8, 21].value = broj;
                                }

                            }
                            if (mt.Contains("Alatnica"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[9, 19].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[9, 20].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[9, 21].value = broj;
                                }

                            }
                            if (mt.Contains("Održavanje"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[10, 19].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[10, 20].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[10, 21].value = broj;
                                }

                            }
                            if (mt.Contains("Šteleri"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[11, 19].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[11, 20].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[11, 21].value = broj;
                                }

                            }
                            if (mt.Contains("Kvaliteta"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[12, 19].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[12, 20].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[12, 21].value = broj;
                                }

                            }
                            if (mt.Contains("Zaštita"))
                            {
                                if (vrstao.Contains("Bolovanje"))
                                {
                                    worksheet.Rows.Cells[13, 19].value = broj;
                                }
                                if (vrstao.Contains("Godi"))
                                {
                                    worksheet.Rows.Cells[13, 20].value = broj;
                                }
                                if (vrstao.Contains("Nije"))
                                {
                                    worksheet.Rows.Cells[13, 21].value = broj;
                                }

                            }

                        }
                        i++;

                        cn.Close();
                    }
                    Console.WriteLine("Izostanci po mjestu troška trenutno vrijeme " + DateTime.Now);
                    
                    // pregled izostanka, popis imena,mt

                    Worksheet worksizostanci = workbook.Worksheets.Item[8] as Worksheet;    // Pregled izostanaka

                    using (SqlConnection cn = new SqlConnection(connectionStringRFIND))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        sql1 = "rfind.dbo.realizacija2 '" + datLDP + "','" + datLDP + "',71," + smj;
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        if (ws == 1)
                        {
                            worksizostanci.Rows.Cells[4, 3].value = worksizostanci.Rows.Cells[4, 3].value + " " + datreps;
                            worksizostanci.Rows.Cells[5, 3].value = smjenanaziv;
                        }

                        string vrstao = "";
                        int i = 0;
                        while (reader.Read())
                        {
                            if (1 == 1)
                            {
                                worksizostanci.Rows.Cells[8 + i, 1].value = reader["mtroska"];
                                worksizostanci.Rows.Cells[8 + i, 2].value = reader["hala"];
                                worksizostanci.Rows.Cells[8 + i, 3].value = reader["smjena"];
                                worksizostanci.Rows.Cells[8 + i, 4].value = reader["linija"].ToString();
                                worksizostanci.Rows.Cells[8 + i, 5].value = reader["prezime"].ToString();
                                worksizostanci.Rows.Cells[8 + i, 6].value = reader["ime"].ToString();
                                worksizostanci.Rows.Cells[8 + i, 7].value = reader["vrsta"].ToString();
                            }
                            i++;

                        }

                        cn.Close();
                    }
                    Console.WriteLine("Izostanci po imenima trenutno vrijeme " + DateTime.Now);

                    dat13 = datLDP;
                    //dat23 = d3;
                    Worksheet workshstelanje = workbook.Worksheets.Item[7] as Worksheet;    // Pregled aktivnosti

                    // aktivnosti
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        sql1 = "rfind.dbo.realizacija2 '" + dat13 + "','" + dat23 + "',6," + smj;
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string pec;
                        int i = 0;
                        int ukupn = 0, ukupp = 0, ukupkol = 0, erv = 0;
                        if (ws == 1)
                        {
                            workshstelanje.Rows.Cells[4, 3].value = workshstelanje.Rows.Cells[4, 3].value + "  " + datreps;
                            workshstelanje.Rows.Cells[5, 3].value = smjenanaziv;
                        }

                        while (reader.Read())
                        {
                            workshstelanje.Rows.Cells[9 + i, 1].value = (reader["hala"].ToString());
                            workshstelanje.Rows.Cells[9 + i, 2].value = (reader["naziv"].ToString());
                            workshstelanje.Rows.Cells[9 + i, 3].value = (reader["aktivnost"].ToString());
                            workshstelanje.Rows.Cells[9 + i, 4].value = reader["pocetak"]; ;  // norma
                            workshstelanje.Rows.Cells[9 + i, 5].value = reader["kraj"]; ;  // komada                    
                            workshstelanje.Rows.Cells[9 + i, 6].value = reader["trajanje_minuta"];
                            workshstelanje.Rows.Cells[9 + i, 7].value = reader["brojrn"];
                            workshstelanje.Rows.Cells[9 + i, 8].value = reader["napomena"];
                            workshstelanje.Rows.Cells[9 + i, 9].value = reader["username"];
                            workshstelanje.Rows.Cells[9 + i, 10].value = reader["vrstanarudzbe"];
                            workshstelanje.Rows.Cells[9 + i, 11].value = reader["nazivpar"];
                            i++;

                        }

                        cn.Close();
                    }
                }
                // linijeeeeeeeeeeeeeeeeeeeeeeeeeeeee1
                //smj = "5";  // po linijama

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                    // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                    string smjen1 = ((int.Parse)(smj) + 10).ToString();  // smjene su 11,12,13  , u proceduri oduzima 10
                    sql1 = "rfind.dbo.ldp_recalc '" + datLDP + "','" + datLDP + "'," + smjen1; ;
//                    sql1 = "rfind.dbo.ldp_recalc '" + datLDP + "','" + datLDP + "',1234" ;

                    int h1 = ws;
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string pec;
                    int i = 0;
                    int ukupn = 0, ukupp = 0, ukupkol = 0;
                    string hala1 = "", smjena1 = "", proiz1 = "", linija1 = "", radninalog1 = "", radninalog2 = "";
                    string hala2 = "", smjena2 = "", proiz2 = "", linija2 = "";
                    string kupac1 = "", kupac2 = "";
                    int norma1 = 0, kolicina1 = 0, kolicina2 = 0, norma2 = 0, norma = 0, kolicina = 0;
                    int prvii = 1, normaukup1 = 0, normaukup2 = 0, erv1 = 0, erv2 = 0, erv = 0, normaukupp = 0;
                    double cijena1 = 0.0, cijena2 = 0.0, norm = 0.0, norm1 = 0.0;
                    double uostvareno = 0.0, ostvareno = 0.0, planirano = 0.0, uos1 = 0.0, uos3 = 0.0;
                    
                    string naziv1 = "";
                    int jos = 1,imared=0;
                    //var imah=reader.Read();
                    reader.Read();

                    while (jos==1)
                    {

                        //if ((ws == 3 && hala1 == "1") || (ws == 4 && hala1 == "3") ) // || (ws <= 2 && (hala1=="1" || hala1=="3") ))
                        //{
                        //}
                        //else
                        //{
                        //      continue;
                        //}

                        //if (reader.Read())
                        //    {
                        //}
                        //else
                        //{
                        //    jos=0;
                        //    continue;
                        //}

                        //if (!reader.IsDBNull(3))
                        //{
                        //}
                        //else
                        //{                             
                        //    continue;
                        //}

                        //if (reader["hala"]== DBNull.Value)
                        //{
                        //    continue;
                        //}


                        hala1 = (reader["hala"].ToString());
                        if (hala1 == "")
                        {
                            hala1 = (reader["halal"].ToString());
                            naziv1 = (reader["naziv"].ToString());
                            worksheet4.Rows.Cells[5 + i, 1].value = dat1;
                            worksheet4.Rows.Cells[5 + i, 2].value = hala1;
                            worksheet4.Rows.Cells[5 + i, 3].value = naziv1;
                            i++;
                            var loop= reader.Read();
                            //if (i>=77)
                            //{
                            //    i = i;
                            //}
                            if (loop)
                                {                                
                                continue;
                            }
                            else
                            {
                                if (imared == 0)
                                {
                                    jos = 0;
                                    continue;
                                }
                            }
                            
                        }

                        imared = 1;
                        smjena1 = (reader["smjena"].ToString());
                        proiz1 = (reader["nazivpro"].ToString());
                        linija1 = (reader["linija"].ToString());
                        if (linija1.Contains("GT600"))
                            {
                            int tttt1 = 1;
                        }
                        string tt = (reader["obrada3"].ToString()).Trim();
                        if (tt == "1")
                        {
                            tt = tt;
                        }
                        radninalog1 = (reader["brojrn"].ToString());
                        kolicina1 = (int.Parse)(reader["kolicinaok"].ToString());
                        normaukup1 = (int.Parse)(reader["normukup"].ToString());
                        erv1 = (int.Parse)(reader["erv"].ToString());
                        norma1 = (int.Parse)(reader["norma"].ToString());
                        cijena1 = (double.Parse)(reader["cijena"].ToString());
                        kupac1 = (reader["kupac"].ToString());
                        

                        double normaukup = 0.0, n1 = 0.0, k1 = 0.0, posto1 = 0.0;
                        int trajanje11 = 0;  // ukupno vrijeme aktivnosti ( štelenaje,kvar stroja,nedostatak materijala,dorada)

                        if (prvii==1)
                        {
                            norma = 0;
                            kolicina = 0;
                            erv = 0;
                            normaukupp = 0;
                            norm = 0;
                            hala2 = (reader["hala"].ToString());
                            linija2 = (reader["linija"].ToString()).TrimEnd();
                            proiz2 = (reader["nazivpro"].ToString());
                            kupac2 = (reader["kupac"].ToString());
                            smjena2 = (reader["smjena"].ToString());
                            radninalog2 = (reader["brojrn"].ToString());
                            cijena2 = (double.Parse)(reader["cijena"].ToString());
                        }

                        while (((hala1 == hala2 && smjena1 == smjena2 && linija1 == linija2 && proiz1 == proiz2 && kupac2 == kupac1) || (prvii == 1)) && (1==1))
                        {

                            //111
                            using (SqlConnection cna = new SqlConnection(connectionString))
                            {
                                cna.Open();
                                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                                dat13 = datLDP + " 06:00:00";
                                string dat33 = dat3 + " 05:59:00";

                                if (smjena2 == "1")
                                {
                                    dat13 = datLDP + " 06:00:00";
                                    dat33 = datLDP + " 13:59:00";
                                }
                                if (smjena2 == "2")
                                {
                                    dat13 = datLDP + " 14:00:00";
                                    dat33 = datLDP + " 21:59:00";
                                }
                                if (smjena2 == "3")
                                {
                                    dat13 = datLDP + " 22:00:00";
                                    dat33 = dats + " 05:59:00";
                                }

                                string sqla = "select * from rfind.dbo.pregled_po_liniji( '" + dat13 + "','" + dat33 + "','" + hala2 + "','" + linija2.TrimEnd() + "')";
                                    if (linija2.TrimEnd()=="07")
                                {
                                    int l = 1;
                                }
                                SqlCommand sqlCommanda = new SqlCommand(sqla, cna);
                                SqlDataReader readera = sqlCommanda.ExecuteReader();
                                string aktivnost1 = "";
                                string trajanje1 = "";
                                string napomena1 = "", rna1 = "";

                                while (readera.Read())
                                {
                                    if (linija2 == "16")
                                    {
                                        int tt211 = 1;
                                    }

                                    aktivnost1 = readera["aktivnost"].ToString();
                                    trajanje1 = readera["trajanje"].ToString();
                                    
                                    napomena1 = readera["napomena"].ToString();
                                    rna1 = readera["brojrn"].ToString().TrimEnd();
                                    if (rna1 != "")
                                    {
                                        if (rna1 != radninalog2)
                                        {
                                            continue;
                                        }
                                    }

                                    if (aktivnost1.Contains("telanje"))   // štelanje
                                    {
                                        worksheet4.Rows.Cells[5 + i, 16].value = trajanje1;
                                        worksheet4.Rows.Cells[5 + i, 17].value = napomena1 + " ( RN: " + rna1 + " )";
                                        trajanje11 = trajanje11 + (int.Parse)(trajanje1);
                                    }
                                    if (aktivnost1.Contains("Otežan"))   // Otežan rad
                                    {
                                        worksheet4.Rows.Cells[5 + i, 18].value = trajanje1;
                                        worksheet4.Rows.Cells[5 + i, 19].value = napomena1 + " ( RN: " + rna1 + " )";
                                        
                                    }
                                    if (aktivnost1.Contains("Kvar"))    // Kvar stroja
                                    {
                                        worksheet4.Rows.Cells[5 + i, 20].value = trajanje1;
                                        worksheet4.Rows.Cells[5 + i, 21].value = napomena1 + " ( RN: " + rna1 + " )";
                                        trajanje11 = trajanje11 + (int.Parse)(trajanje1);
                                    }
                                    if (aktivnost1.Contains("djelatnika"))  // Nedostatak djelatnika
                                    {
                                        worksheet4.Rows.Cells[5 + i, 22].value = trajanje1;
                                        worksheet4.Rows.Cells[5 + i, 23].value = napomena1 + " ( RN: " + rna1 + " )";
                                        
                                    }
                                    if (aktivnost1.Contains("sirovca"))    // Nedostatak sirovca
                                    {
                                        worksheet4.Rows.Cells[5 + i, 20].value = trajanje1;
                                        worksheet4.Rows.Cells[5 + i, 21].value = napomena1 + " ( RN: " + rna1 + " )";
                                        
                                    }
                                    if (aktivnost1.Contains("Dorada"))    // Dorada
                                    {
                                        worksheet4.Rows.Cells[5 + i, 22].value = trajanje1;
                                        worksheet4.Rows.Cells[5 + i, 23].value = napomena1 + " ( RN: " + rna1 + " )";
                                        
                                    }

                                }
                            }

                            ///111

                            worksheet4.Rows.Cells[5 + i, 1].value = dat1; ;
                            worksheet4.Rows.Cells[5 + i, 2].value = hala2;
                            worksheet4.Rows.Cells[5 + i, 3].value = linija2;
                            if (linija2.Contains("VC510") || linija2.Contains("HURCO") || linija2.Contains("INDEX") || linija2.Contains("GT600") || tt == "1")
                            {
                                worksheet4.Rows.Cells[5 + i, 4].value = "Dodatne operacije";
                            }
                            else
                            {
                                worksheet4.Rows.Cells[5 + i, 4].value = "Tokarenje";
                            }

                            worksheet4.Rows.Cells[5 + i, 5].value = proiz2;
                            worksheet4.Rows.Cells[5 + i, 6].value = norma1; //  reader["norma"]; ;  // norma
                            worksheet4.Rows.Cells[5 + i, 7].value = kolicina1; // reader["kolicinaok"]; ;  // komada                    

                            norm1 = norma1 * (normaukup1 - trajanje11) / normaukup1 ;  // preračunata norma, obzirom na vrijeme trajanja aktivnosti
                            worksheet4.Rows.Cells[5 + i, 8].value = norm1; // planirano komada
                            worksheet4.Rows.Cells[5 + i, 9].value = cijena2;

                            ostvareno = kolicina1 * cijena2;
                            planirano = norma1 * cijena2;  //  za normu od 480 minuta

                            k1 = ostvareno;
                            //n1 = (planirano) * (normaukup1-trajanje11) / normaukup1;    // korigirina norma, normaukup = ukupno normirano vrijeme
                            n1 = norm1 * cijena2;
                            
                            //posto1 = (k1 + 0.0000001) / (n1 + 0.0000001);
                            posto1 = (kolicina1+0.000001) / (norm1+0.000001);  // stvarno/planirano

                            if (kolicina1 <= 0.0001 || n1 <= 0.0001)
                            {
                                //posto1 = 0.0;
                            }

                            if (posto1 >= 1.00)
                            {
                                worksheet4.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbGreenYellow;
                            }
                            if (posto1 < 0.900)
                            {

                                worksheet4.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbRed;
                            }
                            if (posto1 >= 0.900 && posto1 < 1.00)
                            {

                                worksheet4.Rows.Cells[5 + i, 12].Interior.Color = XlRgbColor.rgbOrange;
                            }


                            worksheet4.Rows.Cells[5 + i, 10].value = ostvareno;
                            uostvareno = uostvareno + ostvareno;

                            if (hala1 == "1")
                                uos1 = uos1 + ostvareno;
                            if (hala1 == "3")
                                uos3 = uos3 + ostvareno;

                            worksheet4.Rows.Cells[5 + i, 11].value = n1;   // korigirana norma za srtvarno vrijeme rada
                            worksheet4.Rows.Cells[5 + i, 12].value = posto1;
                            worksheet4.Rows.Cells[5 + i, 13].value = ostvareno - planirano;
                            //   worksheet4l.Rows.Cells[5 + i, 12].value = (reader["napomena"].ToString());
                            worksheet4.Rows.Cells[5 + i, 15].value = kupac2;
                            worksheet4.Rows.Cells[5 + i, 1].value = dat1; ;
                            worksheet4.Rows.Cells[5 + i, 2].value = hala2;
                            worksheet4.Rows.Cells[5 + i, 3].value = linija2;

                            if (linija2.Contains("VC510") || linija2.Contains("HURCO") || linija2.Contains("INDEX") || linija2.Contains("GT600") || tt == "1")
                            {
                                worksheet4.Rows.Cells[5 + i, 4].value = "Dodatne operacije";
                            }
                            else
                            {
                                worksheet4.Rows.Cells[5 + i, 4].value = "Tokarenje";
                            }

                            worksheet4.Rows.Cells[5 + i, 5].value = proiz2;
                            worksheet4.Rows.Cells[5 + i, 6].value = norma1; //  reader["norma"]; ;  // norma
                            worksheet4.Rows.Cells[5 + i, 7].value = kolicina1; // reader["kolicinaok"]; ;  // komada                    

                            if (kolicina1 == 0)  // podbačaj radnika
                            {
                                worksheet4.Rows.Cells[5 + i, 29].value = "";
                            }
                            else
                            {
                                worksheet4.Rows.Cells[5 + i, 29].value = norm1 - kolicina1;
                            }

                            i++;

                            ///111...
                            if (1 == 1)
                            {
                                norma = norma + norma1;
                                kolicina = kolicina + kolicina1;
                                erv = erv + erv1;
                                normaukupp = normaukupp + normaukup1;
                                norm = norm + norma1 * erv1 / normaukup1;
                            }

                            if (hala1 == hala2 && smjena1 == smjena2 && proiz1 == proiz2 && linija1 == linija2 && radninalog1 != radninalog2)
                            {
                                norma = norma - norma1;
                                // norm  = norm  - norma1 * erv1 / normaukup1;
                                // erv   = erv   - erv1;
                                normaukupp = normaukupp - normaukup1;
                            }
                            norm = norma * erv / normaukupp;
                            norma1 = norma;
                            erv1 = erv;
                            normaukup1 = normaukupp;
                            norm1 = norm;
                            
                            //hala2 = hala1;
                            //erv2 = erv1;

                            //smjena2 = smjena1;
                            //proiz2 = proiz1;
                            //radninalog2 = radninalog1;
                            //linija2 = linija1;
                            //tt = (reader["obrada3"].ToString()).TrimEnd();
                            //if (tt == "1")
                            //{
                            //    tt = tt;
                            //}
                            //cijena2 = cijena1;
                            //kupac2 = kupac1;
                            prvii = 0;
                           
                            if (!(reader.Read()))   // ako je kraj query-aa
                            {
                                hala1 = "99";
                                jos = 0;
                            }
                            else
                            {
                                hala1 = (reader["hala"].ToString());
                                linija1 = (reader["linija"].ToString()).TrimEnd();
                                proiz1 = (reader["nazivpro"].ToString());
                                kupac1 = (reader["kupac"].ToString());
                                smjena1 = (reader["smjena"].ToString());
                                radninalog1 = (reader["brojrn"].ToString());
                                kolicina1 = (int.Parse)(reader["kolicinaok"].ToString());
                                cijena1 = (double.Parse)(reader["cijena"].ToString());
                                normaukup1 = (int.Parse)(reader["normukup"].ToString());
                                erv1 = (int.Parse)(reader["erv"].ToString());
                                norma1 = (int.Parse)(reader["norma"].ToString());
                                norm1 = norma1 * erv1 / normaukup1;

                                if (hala1 == hala2)
                                {
                                    int t1 = 0;
                                        }

                                if (linija1 == linija2)
                                {
                                    int t1 = 0;
                                }

                                if (proiz1 == proiz2)
                                {
                                    int t1 = 0;
                                }
                                
                                        


                                //tt = (reader["obrada3"].ToString()).Trim();
                            }

                        }
                            prvii = 1;
                        //hala2 = (reader["hala"].ToString());
                        //linija2 = (reader["linija"].ToString());
                        //proiz2 = (reader["nazivpro"].ToString());
                        //kupac2 = (reader["kupac"].ToString());
                        //smjena2 = (reader["smjena"].ToString());
                        //radninalog2 = (reader["brojrn"].ToString());


                        //linija2 = linija1;
                        //norma = norma1;
                        //hala2 = hala1;
                        //proiz2 = proiz1;
                        //kolicina = kolicina1;
                        //kupac2 = kupac1;
                        //cijena2 = cijena1;


                        
                        //    norma1 = norma;

                        //erv = 0;
                        //string brisi1 = (reader["erv"].ToString());
                        //if (brisi1.Length > 0)
                        //{
                        //    erv = (int)((double.Parse)(brisi1));
                        //}
                        //else
                        //{
                        //    erv = 0;
                        //}
                        ostvareno = 0.0; planirano = 0.0;
                        //double normaukup = 0.0,n1=0.0,k1 = 0.0; ;

     //                   if (reader["norma"] != DBNull.Value)
                        {
                            ostvareno = kolicina * cijena2 ;
                            planirano = norma    * cijena2 ;  //  za normu od 480 minuta
                                                          //normaukup = normaukup2; // (double.Parse)(reader["normukup"].ToString());
                        }

                        //double posto1 = 0.00;
                        //double n1 = 0.0, k1 = 0.0;
                        k1 = ostvareno;
                        n1 = (planirano) * erv / normaukupp;    // korigirina norma, normaukup = ukupno normirano vrijeme
                        n1 = norm * cijena2;                        

                        //i++;
                        hala2 = hala1;
                        smjena2 = smjena1;
                        proiz2 = proiz1;
                        radninalog2 = radninalog1;
                        linija2 = linija1;
                        cijena2 = cijena1;
                       // tt = (reader["obrada3"].ToString());
                        norma = norma1;
                        norm = norm1;
                        erv = erv1;
                        normaukupp = normaukup1;
                        kolicina = kolicina1;
                        kupac2 = kupac1;
                        
                    }

                    if (ws < 2)
                        worksheet.Rows.Cells[30, 3].value = uostvareno;

                    if (ws == 3)
                        worksheet.Rows.Cells[30, 3].value = uos1;

                    if (ws == 4)
                        worksheet.Rows.Cells[30, 3].value = uos3;


                    cn.Close();
                }
                Console.WriteLine("Pregled po linijama  trenutno vrijeme " + DateTime.Now);
                // linijeeeeeeeeeeeeeeeeeeeeeeeeeeeee2
                // po linijama

                // Realizacija 2,3,1 smjene - Horvatić V., procjena postotka troškova pločice u odnosu na ukupnu realizaciju
                // dat13 = datLDP + " 06:00:00" ;
                // dat23 = dat3   + " 05:59:00" ;

                if (smj == "1")
                {
                    string datdanas = d2.Year.ToString() + '-' + mm1 + d2.Month.ToString() + '-' + d2.Day.ToString();
                    string datjucer = d2.AddDays(-1).Year.ToString() + '-' + d2.AddDays(-1).Month.ToString() + '-' + d2.AddDays(-1).Day.ToString();
                    Worksheet worksrealizacija = workbook.Worksheets.Item[9] as Worksheet;    // Pregled aktivnosti

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                        // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                        sql1 = "rfind.dbo.realizacija2 '" + datjucer + "','" + datdanas + "',90,1";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        string pec;
                        int i = 7;
                        int ukupn = 0, ukupp = 0, ukupkol = 0, erv = 0;
                        worksrealizacija.Rows.Cells[1, 2].value = datjucer + " - "  + datdanas;

                        while (reader.Read())
                        {
                            worksrealizacija.Rows.Cells[i, 1].value = (reader["hala"].ToString());
                            worksrealizacija.Rows.Cells[i, 2].value = (reader["kupac"].ToString());
                            worksrealizacija.Rows.Cells[i, 3].value = (reader["vrstapro"].ToString());
                            worksrealizacija.Rows.Cells[i, 4].value = reader["kolicina"]; ;  // norma
                            worksrealizacija.Rows.Cells[i, 5].value = reader["vrijednost"]; ;  // komada                    

                            i++;
                        }

                        cn.Close();
                    }
                    Console.WriteLine("Pregled realizacije za 2,3 smjenu ( jucer),1 smjena(danas)   trenutno vrijeme " + DateTime.Now);

                }

                if (1 == 2)
                    {
                        using (SqlConnection cn = new SqlConnection(connectionString))
                        {
                            cn.Open();
                            //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                            // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                            string smjen1 = ((int.Parse)(smj) + 10).ToString();  // smjene su 11,12,13  , u proceduri oduzima 10
                            sql1 = "rfind.dbo.ldp_recalc '" + dat13 + "','" + dat23 + "'," + smjen1;
                            SqlCommand sqlCommand = new SqlCommand(sql1, cn);

                            SqlDataReader reader = sqlCommand.ExecuteReader();
                            //string pec;
                            int i = 0;
                            int ukupp = 0, ukupkol = 0;
                            double posto1 = 0.00, n1 = 0.00, k1 = 0.00, ostvareno = 0.0, planirano = 0.0, uostvareno = 0.0;
                            while (reader.Read())
                            {
                                if (reader["norma"] != DBNull.Value)
                                {
                                    ukupp = ukupp + (int.Parse)(reader["norma"].ToString());
                                    ukupkol = ukupkol + (int.Parse)(reader["kolicinaok"].ToString());
                                }
                                worksheet4.Rows.Cells[5 + i, 1].value = reader["datum"];
                                worksheet4.Rows.Cells[5 + i, 2].value = (reader["smjena"].ToString());
                                worksheet4.Rows.Cells[5 + i, 3].value = (reader["hala"].ToString());
                                worksheet4.Rows.Cells[5 + i, 4].value = (reader["linija"].ToString());
                                worksheet4.Rows.Cells[5 + i, 5].value = (reader["nazivpro"].ToString());
                                worksheet4.Rows.Cells[5 + i, 6].value = reader["norma"];
                                worksheet4.Rows.Cells[5 + i, 7].value = reader["kolicinaok"];
                                string brisi1 = (reader["cijena"].ToString());

                                if (reader["norma"] != DBNull.Value)
                                {
                                    ostvareno = ((double.Parse)(reader["kolicinaok"].ToString()) * ((double.Parse)(reader["cijena"].ToString())));
                                    if (reader["planirano"] == DBNull.Value)
                                    {
                                        planirano = (double.Parse)(reader["norma"].ToString()) * ((double.Parse)(reader["cijena"].ToString()));
                                    }
                                    else
                                    {
                                        planirano = ((double.Parse)(reader["planirano"].ToString()) * ((double.Parse)(reader["cijena"].ToString())));
                                    }
                                }

                                worksheet4.Rows.Cells[5 + i, 8].value = reader["cijena"];
                                worksheet4.Rows.Cells[5 + i, 9].value = ostvareno;
                                worksheet4.Rows.Cells[5 + i, 10].value = planirano;
                                uostvareno = uostvareno + ostvareno;
                                k1 = ostvareno; n1 = planirano;


                                posto1 = k1 / (n1 + 0.0000001);
                                if (k1 < 0.001 || n1 < 0.001)
                                {
                                    posto1 = 0.00;
                                }

                                if (posto1 > 1.00)
                                {
                                    worksheet4.Rows.Cells[5 + i, 11].Interior.Color = XlRgbColor.rgbGreenYellow;
                                }

                                if (posto1 >= 0.900 && posto1 < 1.00)
                                {

                                    worksheet4.Rows.Cells[5 + i, 11].Interior.Color = XlRgbColor.rgbOrange;
                                }
                                if (posto1 >= 0.99 && posto1 < 1.00 && k1 > 0)
                                {
                                    //worksheet3.Rows.Cells[5 + i, 11].Interior.Color = XlRgbColor.rgbYellowGreen;
                                }
                                if (posto1 < 0.900)
                                {

                                    worksheet4.Rows.Cells[5 + i, 11].Interior.Color = XlRgbColor.rgbRed;
                                }


                                worksheet4.Rows.Cells[5 + i, 11].value = posto1;
                                worksheet4.Rows.Cells[5 + i, 12].value = ostvareno - planirano;
                                worksheet4.Rows.Cells[5 + i, 13].value = (reader["napomena"].ToString());
                                worksheet4.Rows.Cells[5 + i, 14].value = (reader["kupac"].ToString());

                                i++;
                            }
                            worksheet.Rows.Cells[29, 3].value = uostvareno;
                            cn.Close();
                        }

                    }
                
            


            }     // end for ws<=4

            string smjenav;
            smjenav = "Smjena";
            Environment.SetEnvironmentVariable(smjenav, "3");

            //fileName = smj + "_" + fileName;
            var fi = new FileInfo(fileName);
                if (fi.Exists) File.Delete(fileName);

            fi = new FileInfo(fileNamebv);
            if (fi.Exists) File.Delete(fileNamebv);
            excel.Application.ActiveWorkbook.SaveAs(fileName);

            // ovja puta brisi sve vrijednosti u EUR
            Worksheet worksheetbv = workbook.Worksheets.Item[1] as Worksheet;
            worksheetbv.Rows.Cells[30, 3].value = "";    // ukupna vrijednost
            // kaliona
            worksheetbv.Rows.Cells[34, 9].value = "";    // realizacija

            Range startCell = worksheetbv.Cells[7, 6];  // kolona sa ostvarenim eurima, T, DO
            Range endCell = worksheetbv.Cells[22, 7];
            worksheetbv.Range[startCell, endCell].Value = "";

            Worksheet worksheetbvt = workbook.Worksheets.Item[2] as Worksheet;  // tjedni
            worksheetbvt.Rows.Cells[30, 3].value = "";    // ukupna vrijednost
            // kaliona
            worksheetbvt.Rows.Cells[34, 9].value = "";    // realizacija

            startCell = worksheetbvt.Cells[7, 6];  // kolona sa ostvarenim eurima, T, DO
            endCell = worksheetbvt.Cells[22, 7];
            worksheetbvt.Range[startCell, endCell].Value = "";

            worksheetbvt.Rows.Cells[30, 3].value = "";    // ukupna vrijednost
            // kaliona
            worksheetbvt.Rows.Cells[34, 9].value = "";    // realizacija


            Worksheet worksheetbv1 = workbook.Worksheets.Item[3] as Worksheet;  // p1
            worksheetbv1.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur
            worksheetbv1.Rows.Cells[34, 9].value = "";    // kaliona realizacija
            worksheetbv1.Rows.Cells[30, 3].value = "";    // ukupna vrijednost
            // kaliona
            worksheetbv1.Rows.Cells[34, 9].value = "";    // realizacija

            startCell = worksheetbv1.Cells[7, 6];  // kolona sa ostvarenim eurima, T, DO
            endCell = worksheetbv1.Cells[22, 7];
            worksheetbv1.Range[startCell, endCell].Value = "";

            worksheetbv1.Rows.Cells[30, 3].value = "";    // ukupna vrijednost
            // kaliona
            worksheetbv1.Rows.Cells[34, 9].value = "";    // realizacija

            Worksheet worksheetbv3 = workbook.Worksheets.Item[4] as Worksheet;  // p3
            worksheetbv3.Rows.Cells[30, 3].value = "";    // ukupna vrijednost u eur
            worksheetbv3.Rows.Cells[34, 9].value = "";    // kaliona realizacija
            worksheetbv3.Rows.Cells[30, 3].value = "";    // ukupna vrijednost
            // kaliona
            worksheetbv3.Rows.Cells[34, 9].value = "";    // realizacija

            startCell = worksheetbv3.Cells[7, 6];  // kolona sa ostvarenim eurima, T, DO
            endCell = worksheetbv3.Cells[22, 7];
            worksheetbv3.Range[startCell, endCell].Value = "";

            worksheetbv3.Rows.Cells[30, 3].value = "";    // ukupna vrijednost
            // kaliona
            worksheetbv3.Rows.Cells[34, 9].value = "";    // realizacija
            

            Worksheet worksheetbvlinija = workbook.Worksheets.Item[6] as Worksheet;  // po linijama
            
            //startCell = worksheetbvlinija.Cells[5, 8];  // kolona sa cijenama
            //endCell = worksheetbvlinija.Cells[110, 8];

            //worksheetbvlinija.Range[startCell, endCell].Value = "";

            startCell = worksheetbvlinija.Cells[5, 9];  // kolona sa ostvarenim eurima
            endCell = worksheetbvlinija.Cells[110, 9];

            worksheetbvlinija.Range[startCell, endCell].Value = "";

            startCell = worksheetbvlinija.Cells[5, 10];  // kolona sa planiranim eurima
            endCell = worksheetbvlinija.Cells[110, 10];

            worksheetbvlinija.Range[startCell, endCell].Value = "";


            startCell = worksheetbvlinija.Cells[5, 11];  // kolona sa  planirano eura
            endCell = worksheetbvlinija.Cells[110, 11];
            worksheetbvlinija.Range[startCell, endCell].Value = "";

            startCell = worksheetbvlinija.Cells[5, 13];  // kolona sa  razlikom eura
            endCell = worksheetbvlinija.Cells[110, 13];
            worksheetbvlinija.Range[startCell, endCell].Value = "";

            excel.Application.ActiveWorkbook.SaveAs(fileNamebv);
            
            Console.WriteLine("Snimljen file ddmmyyyy   trenutno vrijeme " + DateTime.Now);
                //Console.ReadKey();
                //workbook.Close(false);
                //excel.Application.Quit();
                //excel.Quit();

                //workbook.Close(true, Type.Missing, Type.Missing);
                workbook.Close(false, Type.Missing, Type.Missing);
                excel.Application.Quit();
                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                workbook = null;                
                app = null;
                int test = 10;

            //            string fileName = @"C:\brisi\dsr20062017.xlsm";
            MailMessage mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr,cakanic.s@feroimpex.hr");

            if (test==0)
                    mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr,legac.b@feroimpex.hr,legac.h@feroimpex.hr,legac.z@feroimpex.hr,horvatic.v@feroimpex.hr");
// MailMessage mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr, srecckog@gmail.com");
                SmtpClient client = new SmtpClient();
                client.Port = 25;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential("gasparic.s@feroimpex.hr", "gasparic1");
                client.Host = "mail.feroimpex.hr";
                mail.Subject = "Shiftly & weekly production report " + datreps + " shift " + smjenanaziv;
                mail.Body = "Shiftly & weekly production report for " + datreps + " shift " + smjenanaziv;

                Attachment attachment = new Attachment(fileName, System.Net.Mime.MediaTypeNames.Application.Octet);
                System.Net.Mime.ContentDisposition disposition = attachment.ContentDisposition;
                disposition.CreationDate = File.GetCreationTime(fileName);
                disposition.ModificationDate = File.GetLastWriteTime(fileName);
                disposition.ReadDate = File.GetLastAccessTime(fileName);
                disposition.FileName = Path.GetFileName(fileName);
                disposition.Size = new FileInfo(fileName).Length;
                disposition.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;

                mail.Attachments.Add(attachment);
                client.Send(mail);
                client.Dispose();


            // ovaj puta šalji mail bez vrijednosti
            if (test == 0)
               mail = new MailMessage("gasparic.s@feroimpex.hr","barbic.i@feroimpex.hr,grgecic.d@feroimpex.hr , cakanic.s@feroimpex.hr,kicin.d@feroimpex.hr,gasparic.s@feroimpex.hr,srecckog@gmail.com,stefanac.d@feroimpex.hr,hren.h@feroimpex.hr,darapi.i@feroimpex.hr,gradiski.b@feroimpex.hr,vladic.p@feroimpex.hr,igrec.m@feroimpex.hr,jancic.d@feroimpex.hr,golesic.m@feroimpex.hr");
            //           mail = new MailMessage("gasparic.s@feroimpex.hr", "gasparic.s@feroimpex.hr, srecckog@gmail.com");
            SmtpClient client2 = new SmtpClient();
            client2.Port = 25;
            client2.DeliveryMethod = SmtpDeliveryMethod.Network;
            client2.UseDefaultCredentials = false;
            client2.Credentials = new System.Net.NetworkCredential("gasparic.s@feroimpex.hr", "gasparic1");
            client2.Host = "mail.feroimpex.hr";
            mail.Subject = "_Shiftly & weekly production report " + datreps + " shift " + smjenanaziv;
            mail.Body = "Shiftly & weekly production report for " + datreps + " shift " + smjenanaziv;

            attachment = new Attachment(fileNamebv, System.Net.Mime.MediaTypeNames.Application.Octet);
            disposition = attachment.ContentDisposition;
            disposition.CreationDate = File.GetCreationTime(fileNamebv);
            disposition.ModificationDate = File.GetLastWriteTime(fileNamebv);
            disposition.ReadDate = File.GetLastAccessTime(fileNamebv);
            disposition.FileName = Path.GetFileName(fileNamebv);
            disposition.Size = new FileInfo(fileNamebv).Length;
            disposition.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;

            mail.Attachments.Add(attachment);
            client2.Send(mail);
            client2.Dispose();


            var processes = from p in System.Diagnostics.Process.GetProcessesByName("EXCEL")
                                select p;

                foreach (var process in processes)
                {
                    int z = 0;
                    if (process.MainWindowTitle.Contains("Microsoft Excel"))
                        process.Kill();
                }


                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

            }
        }

    }

