using ClosedXML.Excel;
using MaasBordrosuUI.Class;
using MaasBordrosuUI.Class.Abstract;
using System.Reflection;
using System.Text.Json;
using System.Reflection;

namespace MaasBordrosuUI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            lblSayfaBilgisi.Visible = false;
            panelGiris.Visible = true;
            panelDokumantasyon.Visible = false;
            panelPersonelYonetimi.Visible = false;
            panelRaporlar.Visible = false;
        }

        List<Personel> personellerListesi = new List<Personel>();
        List<Personel> memurlarListesi = new List<Personel>();
        List<Personel> yoneticilerListesi = new List<Personel>();

        private void Form1_Load(object sender, EventArgs e)
        {
            lblSayfaBilgisi.Text = "Anasayfa";
            timer1.Start();
            timer1.Interval = 1000;
            panelGiris.Visible = true;
            lstvBilgiCRUD.Columns.Add("Ad", 120);
            lstvBilgiCRUD.Columns.Add("Soyad", 120);
            lstvBilgiCRUD.Columns.Add("Stat�", 100);
            lstvBilgiCRUD.Columns.Add("Telefon Numaras�", 150);
            lstvBilgiCRUD.Columns.Add("Do�um Tarihi", 80);
            lstvBilgiCRUD.Columns.Add("Memur Derecesi(Memur ise)", 50);
            lstvBilgiCRUD.Columns.Add("�al��ma saati", 80);
            lstvBilgiCRUD.Columns.Add("Maa�", 100);
            cmbMemurDerecesi.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbStatu.DropDownStyle = ComboBoxStyle.DropDownList;
            lstvBilgiCRUD.FullRowSelect = true;



            lstvMemurlar.Columns.Add("Ad", 80);
            lstvMemurlar.Columns.Add("Soyad", 80);
            lstvMemurlar.Columns.Add("Stat�", 80);
            lstvMemurlar.Columns.Add("Telefon No", 100);
            lstvMemurlar.Columns.Add("Do�um Tarihi", 100);
            lstvMemurlar.Columns.Add("�al��ma Saati", 120);
            lstvMemurlar.Columns.Add("Maa�", 100);
            lstvMemurlar.Columns.Add("Mesai �creti", 80);
            lstvMemurlar.Columns.Add("Memur Derecesi", 100);


            lstvYoneticiler.Columns.Add("Ad", 100);
            lstvYoneticiler.Columns.Add("Soyad", 100);
            lstvYoneticiler.Columns.Add("Stat�", 100);
            lstvYoneticiler.Columns.Add("Telefon No", 100);
            lstvYoneticiler.Columns.Add("Saatlik �cret", 130);
            lstvYoneticiler.Columns.Add("�al��ma Saati", 150);
            lstvYoneticiler.Columns.Add("Bonus", 80);
            lstvYoneticiler.Columns.Add("Maa�", 100);
        }


        #region Desing
        private void timer1_Tick(object sender, EventArgs e)
        {
            lblTimer.Text = DateTime.Now.ToString("yyyy.MM.dd   HH:mm");
        }

        private void btnAnaSayfa_MouseEnter(object sender, EventArgs e)
        {
            btnAnaSayfa.BackColor = ColorTranslator.FromHtml("34; 33; 74");
        }

        private void btnAnaSayfa_MouseLeave(object sender, EventArgs e)
        {
            btnAnaSayfa.BackColor = ColorTranslator.FromHtml("31; 30; 68");
        }

        private void btnPersonelYonetimi_MouseEnter(object sender, EventArgs e)
        {
            btnPersonelYonetimi.BackColor = ColorTranslator.FromHtml("34; 33; 74");
        }

        private void btnPersonelYonetimi_MouseLeave(object sender, EventArgs e)
        {
            btnPersonelYonetimi.BackColor = ColorTranslator.FromHtml("31; 30; 68");
        }

        private void btnDokumantasyon_MouseEnter(object sender, EventArgs e)
        {
            btnDokumantasyon.BackColor = ColorTranslator.FromHtml("34; 33; 74");
        }

        private void btnDokumantasyon_MouseLeave(object sender, EventArgs e)
        {
            btnDokumantasyon.BackColor = ColorTranslator.FromHtml("31; 30; 68");
        }

        private void btnRaporlar_MouseEnter(object sender, EventArgs e)
        {
            btnRaporlar.BackColor = ColorTranslator.FromHtml("34; 33; 74");
        }

        private void btnRaporlar_MouseLeave(object sender, EventArgs e)
        {
            btnRaporlar.BackColor = ColorTranslator.FromHtml("31; 30; 68");
        }
        private void btnCikisYap_MouseEnter(object sender, EventArgs e)
        {
            btnCikisYap.BackColor = ColorTranslator.FromHtml("34; 33; 74");
        }
        private void btnCikisYap_MouseLeave(object sender, EventArgs e)
        {
            btnCikisYap.BackColor = ColorTranslator.FromHtml("31; 30; 68");
        }
        private void btnSayfayiAsagiAl_MouseEnter(object sender, EventArgs e)
        {
            btnSayfayiAsagiAl.BackColor = ColorTranslator.FromHtml("34; 33; 74");
        }
        private void btnSayfayiAsagiAl_MouseLeave(object sender, EventArgs e)
        {
            btnSayfayiAsagiAl.BackColor = ColorTranslator.FromHtml("31; 30; 68");
        }

        #endregion


        #region Yapt���m Desing i�in Visible ve label text d�zenlemesi

        //Desing'� destekleyen birka� olay.
        //Form yerine panel kullan�p farkl� bir proje ortaya ��karmak istedi�im i�in panellerin g�r�n�rl���n� button'lar�n click event'i ile de�i�tiriyorum.
        //Form ekran�n� kapatacak veya form ekran�n� a�a�� atacak button'u kendim tasarlad���m i�in i�levini yerine getirebilmesi i�in birka� kod blo�u ekledim.

        private void btnCikisYap_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnSayfayiAsagiAl_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void btnAnaSayfa_Click(object sender, EventArgs e)
        {
            lblSayfaBilgisi.Visible = true;
            lblSayfaBilgisi.Text = "Ana Sayfa";
            panelGiris.Visible = true;
            panelDokumantasyon.Visible = false;
            panelPersonelYonetimi.Visible = false;
            panelRaporlar.Visible = false;

        }

        private void btnPersonelYonetimi_Click(object sender, EventArgs e)
        {
            lblSayfaBilgisi.Visible = true;
            lblSayfaBilgisi.Text = "Personel Y�netimi";
            panelGiris.Visible = false;
            panelDokumantasyon.Visible = false;
            panelPersonelYonetimi.Visible = true;
            panelRaporlar.Visible = false;
        }

        private void btnDokumantasyon_Click(object sender, EventArgs e)
        {
            lblSayfaBilgisi.Visible = true;
            panelDokumantasyon.Visible = true;
            lblSayfaBilgisi.Text = "Dok�mantasyon";
            panelGiris.Visible = false;
            panelPersonelYonetimi.Visible = false;
            panelRaporlar.Visible = false;
        }

        private void btnRaporlar_Click(object sender, EventArgs e)
        {
            lblSayfaBilgisi.Visible = true;
            lblSayfaBilgisi.Text = "Raporlar";
            panelGiris.Visible = false;
            panelDokumantasyon.Visible = false;
            panelPersonelYonetimi.Visible = false;
            panelRaporlar.Visible = true;
        }
        #endregion


        private void btnEkle_Click(object sender, EventArgs e)
        {
            ElemanEkle();

        }


        #region Eleman ekleme metodu
        public void ElemanEkle()
        {
            //TextBox'lardan, Combobox'lar, DateTimePicker'dan ki�inin bilgilerini al�yorum,
            string ad = txtAd.Text;
            string soyad = txtSoyad.Text;
            string telefonNo = mtbTelNo.Text;
            DateTime dogumTarihi = dtpDogumTarihi.Value;
            int personelYasi = DateTime.Now.Year - dogumTarihi.Year;
            string calismaSaati = txtCalismaSaati.Text;

            int calismaSaatiKontrol;
            decimal yoneticiSaatlikUcret;
            decimal bonus = 0;

            //Burada stat� bilgisini de al�p onda g�re olas� hatalar� �nceden engelliyorum,
            string statu = cmbStatu.Text;
            if (string.IsNullOrWhiteSpace(statu))
            {
                MessageBox.Show("L�tfen personelin stat�s�n� se�iniz.");
                return;//Mesela stat� Combobox'�n� bo� b�rakt���nda alaca�� hata vb.

            }
            if (string.IsNullOrWhiteSpace(ad) || string.IsNullOrWhiteSpace(soyad))
            {
                MessageBox.Show("Ad ve Soyad k�sm� bo� b�rak�lamaz.");
                return;
            }
            if (!HarfKontrol(ad) || !HarfKontrol(soyad))
            {
                MessageBox.Show("L�tfen ad ve soyad k�sm�n� istenilen formatta giriniz.");
                return;//Ad ve soyad Textbox'lar�na harf harici ba�ka bir �ey girdi�inde alaca�� hata.
            }
            if (personelYasi < 18)
            {
                MessageBox.Show("Personel ya�� 18'den k���k olamaz.");
                return;
            }
            if (TelefonNumarasiGecerliMi(telefonNo))
            {
                MessageBox.Show("L�tfen ge�erli bir telefon numaras� giriniz.");
                return;
            }
            if (!int.TryParse(calismaSaati, out calismaSaatiKontrol))
            {
                MessageBox.Show("L�tfen ge�erli bir �al��ma saati giriniz.");
                return;
            }

            //E�er ekleyece�i ki�i memur ise burada ona g�re i�lemler yap�yorum,
            if (statu == "Memur")
            {
                if (cmbMemurDerecesi.SelectedIndex == -1)
                {
                    MessageBox.Show("Memur derecesi se�iniz.");
                    return;//Stat� memur ise memur derecesi se�mesi gerekti�ini mb ile s�yl�yorum.
                }
                int derece = cmbMemurDerecesi.SelectedIndex + 1;
                //Class'lar�mda olu�turdu�um constructor'a bilgi aktar�m� yap�yorum,
                Memur memur = new Memur(ad, soyad, telefonNo, dogumTarihi, calismaSaatiKontrol, derece);
                //Personel oldu�u i�in ve memur oldu�u i�in personel ve memur dosyas�na ekliyorum.
                personellerListesi.Add(memur);
                memurlarListesi.Add(memur);
                MessageBox.Show("Memur Personel Ba�ar�yla Eklendi.");
                //Personel ekledi�imiz sayfadaki Listview i�erisine bilgi aktar�m� yapmak i�in burada Listview bir de�i�ken olu�turup bilgileri g�nderiyorum,
                ListViewItem memurLSTV = new ListViewItem(new string[]
                {   memur.Ad,
                    memur.Soyad,
                    "Memur",
                    memur.TelefonNo,
                    memur.DogumTarihi.ToShortDateString(),
                    memur.Derece.ToString(),
                    memur.CalismaSaati.ToString(),
                    memur.Maas.ToString("C"),
                    memur.MesaiUcreti.ToString("C"),
                    memur.Derece.ToString()
                });
                //G�z�kmesi i�in de ekliyorum.
                lstvBilgiCRUD.Items.Add(memurLSTV);
                //Dok�mantasyondaki Memur listesine ekleyebilmek i�in tekrar olu�turuyorum ��nk� bir nesne birden fazla Listview'a eklenemez.
                ListViewItem memurlarListview = new ListViewItem(new string[]
                {
                    memur.Ad,
                    memur.Soyad,
                    "Memur",
                    memur.TelefonNo,
                    memur.DogumTarihi.ToShortDateString(),
                    memur.Derece.ToString(),
                    memur.MesaiUcreti.ToString(),
                    memur.CalismaSaati.ToString(),
                    memur.Maas.ToString("C")

                });
                //Ayn� �ekilde burada da Listview'a ekliyorum.
                lstvMemurlar.Items.Add(memurlarListview);

            }
            //E�er ekleyece�i personel Y�netici ise,
            if (statu == "Y�netici")
            {
                if (string.IsNullOrWhiteSpace(txtYoneticiSaatlikUcret.Text) || string.IsNullOrWhiteSpace(txtYoneticiBonusu.Text))
                {
                    MessageBox.Show("Y�netici saatlik �cretini veya bonus �deme bilgisini doldurunuz.");
                    return;//Bonusu ve saatlik �creti girmek zorunda oldu�u i�in burada �nlem al�yorum.
                }
                if (!decimal.TryParse(txtYoneticiSaatlikUcret.Text, out yoneticiSaatlikUcret) || !decimal.TryParse(txtYoneticiBonusu.Text, out bonus))
                {
                    MessageBox.Show("L�tfen y�netici saatlik �cretini veya bonusunu ge�erli formatta giriniz. (�rne�in=500.5)");
                    return;//Y�neticinin saatlik �cretini ge�erli formatta girip girmedi�ini kontrol ediyorum, sadece int de�erleri kabul ediyorum.
                }
                if (Convert.ToDecimal(txtYoneticiSaatlikUcret) < 500)
                {
                    MessageBox.Show("Y�neticinin saatlik �creti 500'den k���k olamaz!");
                    return;//�sterlerde yazan saatlik �cretin 500'den k���k olamayaca��yd�, onun kontrol�n� yap�yorum.
                }
                //Listeye aktar�m yapabilmek i�in Y�netici s�n�f�n�n(Miras ald��� personel s�n�f� ayr�ca) constructor'una bilgi g�nderiyorum
                Yonetici yonetici = new Yonetici(ad, soyad, telefonNo, dogumTarihi, calismaSaatiKontrol, yoneticiSaatlikUcret, bonus);
                //Personel oldu�u i�in ve Y�netici oldu�u i�in listeye ekliyorum.
                personellerListesi.Add(yonetici);
                yoneticilerListesi.Add(yonetici);
                MessageBox.Show("Y�netici Personel Ba�ar�yla Eklendi.");
                //Personel ekledi�imiz sayfadaki Listview i�erisine bilgi aktar�m� yapmak i�in burada Listview bir de�i�ken olu�turup bilgileri g�nderiyorum,
                ListViewItem yoneticiLSTV = new ListViewItem(new string[]
                {
                    yonetici.Ad,
                    yonetici.Soyad,
                    "Yonetici",
                    yonetici.TelefonNo,
                    yonetici.DogumTarihi.ToShortDateString(),
                    "",
                    yonetici.CalismaSaati.ToString(),
                    yonetici.Maas.ToString("C")
                });
                //Dok�mantasyondaki Listview'a atamak i�in tekrar olu�turuyorum.
                ListViewItem yoneticiListview = new ListViewItem(new string[]
                {
                    yonetici.Ad,
                    yonetici.Soyad,
                    "Yonetici",
                    yonetici.TelefonNo,
                    yonetici.SaatlikUcret.ToString(),
                    yonetici.CalismaSaati.ToString(),
                    yonetici.Bonus.ToString(),
                    yonetici.Maas.ToString("C")
                });
                //Sonra g�nderiyorum.
                lstvBilgiCRUD.Items.Add(yoneticiLSTV);
                lstvYoneticiler.Items.Add(yoneticiListview);

            }
        }
        #endregion


        #region Telefon numaras� format� kontrol
        public bool TelefonNumarasiGecerliMi(string telefonNumarasi)
        {
            //Girilen formada uygun olup olmad���n� kontrol etmek i�in regex olu�turuyorum.
            //Bu formak xxx-xxx-xxxx olarak int de�erleri kabul edecek.
            var reg = new System.Text.RegularExpressions.Regex(@"^5\d{2} \d{3} \d{4}$");
            return reg.IsMatch(telefonNumarasi);
        }
        #endregion


        #region Harf kontrol 
        public bool HarfKontrol(string metin)
        {
            //�stedi�im formatta regex olu�turuyorum.
            //Sadece alfabetik harflari k���k ve b�y�k formatta kabul edecek.
            var reg = new System.Text.RegularExpressions.Regex(@"^[a-zA-Z������������]+$");
            return reg.IsMatch(metin);
        }
        #endregion


        #region Combobox'taki se�ili de�er kontrol
        private void cmbStatu_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Burada olu�turdu�um Desing'dan kaynakl� Enable'�n� de�i�tirmem gereken i�lemler vard� onlar� ger�ekle�tiriyorum.
            string statu = cmbStatu.SelectedItem.ToString();
            if (statu == "Memur")
            {
                //��nk� Memur bir ki�i ekleyecekse Y�netici saatlik �cretine ve bonusa eri�memeli
                cmbMemurDerecesi.Enabled = true;
                txtYoneticiSaatlikUcret.Enabled = false;
                txtYoneticiBonusu.Enabled = false;
            }
            if (statu == "Y�netici")
            {
                //Burada da e�er Y�neticiyse Memur derecesine eri�ememeli
                txtYoneticiSaatlikUcret.Enabled = true;
                txtYoneticiBonusu.Enabled = true;
                cmbMemurDerecesi.Enabled = false;
            }
        }
        #endregion


        #region JSON format�nda verileri kaydetme
        private void JsonVeriKaydet()
        {
            //Herhangi bir hatay� yakalamak i�in try-catch bulo�una al�yorum,
            try
            {
                //Projenin oldu�u lokasyonu ve Proje i�erisine olu�acak klas�r�n yolunu belirliyorum. Sonras�nda hedefdizi i�erisine hangi isimle kaydedilece�ini belirliyorum.
                string projeDizi = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string HedefDizi = Path.Combine(projeDizi, @"..\..\..\", "JSON");
                string dosyaYolu = Path.Combine(HedefDizi, "personelVerileri.json");

                if (!Directory.Exists(HedefDizi))
                {
                    Directory.CreateDirectory(HedefDizi);
                }
                //iki de�erli tutmak istedi�im i�in dictionary bir de�i�ken listesi olu�turuyorum.
                List<Dictionary<string, string>> mevcutVeriler = new List<Dictionary<string, string>>();
                if (File.Exists(dosyaYolu))
                {
                    //E�er daha �nce olu�turulmu� bir JSON varda t�m dosyay� okuyorum ve 
                    //E�er bo� de�ilse yeni JSON verilerini �st�ne yazd�r�yorum.
                    string eskiJson = File.ReadAllText(dosyaYolu);
                    var veriListesi = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(eskiJson);
                    if (veriListesi != null)
                    {
                        mevcutVeriler = veriListesi;
                    }
                }

                //lstvBilgiCRUD listesi i�erisindeki her item'� geziyorum
                foreach (ListViewItem item in lstvBilgiCRUD.Items)
                {
                    Dictionary<string, string> satir = new Dictionary<string, string>
                    {
                        //dictionary'nin i�erisindeki ilk string de�eri kendim verip ikinci string de�erini ise item'daki subitems'dan al�yorum.
                        {"Ad",item.SubItems[0].Text},
                        {"Soyad",item.SubItems[1].Text},
                        {"Statu",item.SubItems[2].Text},
                        {"Telefon No",item.SubItems[3].Text},
                        {"Dogum Tarihi",item.SubItems[4].Text},
                    };
                    //E�er memur ise ona g�re i�lem yapmak i�in subitem'daki 3. ��eyi kontrol ediyorum ��nk� 3. ��e stat�'y� belirliyor.
                    if (item.SubItems[2].Text == "Memur")
                    {
                        satir.Add("Memur derecesi", item.SubItems[5].Text);
                        satir.Add("Calisma Saati", item.SubItems[6].Text);
                        satir.Add("Maas", item.SubItems[7].Text);
                    }
                    //Ayn�s�n� Y�netici i�in yap�yorum.
                    if (item.SubItems[2].Text == "Yonetici")
                    {
                        foreach (ListViewItem yoneticiItem in lstvYoneticiler.Items)
                        {
                            //Burada Y�netici listesinin i�inden detayl� bilgi �ekiyorum.
                            if (yoneticiItem.SubItems[0].Text == item.SubItems[0].Text && yoneticiItem.SubItems[1].Text == item.SubItems[1].Text)
                            {
                                satir.Add("Saatlik Ucret", yoneticiItem.SubItems[4].Text);
                                satir.Add("Bonus", yoneticiItem.SubItems[5].Text);
                                satir.Add("Calisma Saati", yoneticiItem.SubItems[6].Text);
                                satir.Add("Maas", yoneticiItem.SubItems[7].Text);
                            }

                        }
                    }
                    //Mevcut verilerin i�erisine sat�r� ekliyorum.
                    mevcutVeriler.Add(satir);
                }
                //Burada sat�r sat�r olmas� i�in JSON ayarlar�n� yap�yorum.
                var jsonAyar = new JsonSerializerOptions { WriteIndented = true };
                string yeniJson = JsonSerializer.Serialize(mevcutVeriler, jsonAyar);
                File.WriteAllText(dosyaYolu, yeniJson);
                MessageBox.Show("Veriler .JSON format�nda kaydedildi.");


            }
            catch (Exception ex)
            {
                //Hatalar� yakal�yorum.
                MessageBox.Show("Bir hata olu�tu.");
            }
        }
        #endregion


        #region Excel olu�tur metodu
        private void ExcelOlustur()
        {
            try
            {
                //Using kullan�yorum ��nk� kaynaklar� otomatik olarak y�netmem laz�m.
                using (var workbook = new XLWorkbook())
                {
                    //Bu k�s�mda ba�l�klar� belirliyorum.
                    var workSheet = workbook.AddWorksheet("PersonelRaporu");
                    workSheet.Cell(1, 1).Value = "Ad";
                    workSheet.Cell(1, 2).Value = "Soyad";
                    workSheet.Cell(1, 3).Value = "Stat�";
                    workSheet.Cell(1, 4).Value = "Telefon No";
                    workSheet.Cell(1, 7).Value = "�al��ma Saati";
                    workSheet.Cell(1, 8).Value = "Maa�";
                    //Sat�r belirtmem laz�m ��nk� ilk sat�rda ba�l�klar var.
                    int satir = 2;

                    //Memur listesini gezip i�erisindeki verileri olu�turdu�um excel i�erisine g�nderiyorum,
                    foreach (ListViewItem item in lstvMemurlar.Items)
                    {
                        workSheet.Cell(satir, 1).Value = item.SubItems[0].Text;
                        workSheet.Cell(satir, 2).Value = item.SubItems[1].Text;
                        workSheet.Cell(satir, 3).Value = "Memur";
                        workSheet.Cell(satir, 4).Value = item.SubItems[3].Text;
                        workSheet.Cell(satir, 7).Value = item.SubItems[5].Text;
                        workSheet.Cell(satir, 8).Value = item.SubItems[6].Text;
                        satir++;

                    }
                    //Ayn� i�lemi Y�netici listesi i�in yap�yorum.
                    foreach (ListViewItem item in lstvYoneticiler.Items)
                    {
                        workSheet.Cell(satir, 1).Value = item.SubItems[0].Text;
                        workSheet.Cell(satir, 2).Value = item.SubItems[1].Text;
                        workSheet.Cell(satir, 3).Value = "Y�netici";
                        workSheet.Cell(satir, 4).Value = item.SubItems[3].Text;
                        workSheet.Cell(satir, 7).Value = item.SubItems[6].Text;
                        workSheet.Cell(satir, 8).Value = item.SubItems[7].Text;
                        satir++;
                    }
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        //Burada format, dosya ad� gibi �zellikleri belirliyorum.
                        saveFileDialog.Filter = "Excel Files |*.xlsx";
                        saveFileDialog.Title = "Excel Dosyas�n� Kaydet";
                        saveFileDialog.FileName = "PersonelRaporu.xlsx";
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string dosyaYolu = saveFileDialog.FileName;
                            workbook.SaveAs(dosyaYolu);
                            MessageBox.Show("Excel Ba�ar�yla Olu�turuldu.");
                        }
                    }

                }

            }
            catch (Exception)
            {
                //Olas� hatalar� yakal�yorum.
                MessageBox.Show("Bir Sorun Olu�tu.");
            }
        }
        #endregion

        private void btnJSON_Click(object sender, EventArgs e)
        {
            JsonVeriKaydet();
        }

        private void btnPDF_Click(object sender, EventArgs e)
        {
            ExcelOlustur();
        }


        #region G�ncelle click event'i
        private void btnGuncelle_Click(object sender, EventArgs e)
        {
            //E�er personel se�iliyse:
            if (lstvBilgiCRUD.SelectedItems.Count > 0)
            {
                //Bilgilerin hepsini tekrar al�yorum,
                var secilenKisi = lstvBilgiCRUD.SelectedItems[0];
                string ad = txtAd.Text;
                string soyad = txtSoyad.Text;
                string telefonNo = mtbTelNo.Text;
                DateTime dogumTarihi = dtpDogumTarihi.Value;
                int personelYasi = DateTime.Now.Year - dogumTarihi.Year;
                string calismaSaati = txtCalismaSaati.Text;

                int calismaSaatiKontrol;
                decimal yoneticiSaatlikUcret;
                decimal bonus = 0;

                //Kontrolleri yap�p stat�'ye g�re tekrar i�lem yapaca��m.
                //Stat� se�ili mi kontrol�, ad soyad bo� mu? bo� de�ilse istedi�im formatta m�? personelin ya�� 18'den b�y�k m�? telefon numaras� istedi�im formatta m�? gibi kontrolleri yap�yorum.
                string statu = cmbStatu.Text;
                if (string.IsNullOrWhiteSpace(statu))
                {
                    MessageBox.Show("L�tfen personelin stat�s�n� se�iniz.");
                    return;

                }
                if (string.IsNullOrWhiteSpace(ad) || string.IsNullOrWhiteSpace(soyad))
                {
                    MessageBox.Show("Ad ve Soyad k�sm� bo� b�rak�lamaz.");
                    return;
                }
                if (!HarfKontrol(ad) || !HarfKontrol(soyad))
                {
                    MessageBox.Show("L�tfen ad ve soyad k�sm�n� istenilen formatta giriniz.");
                    return;
                }
                if (personelYasi < 18)
                {
                    MessageBox.Show("Personel ya�� 18'den k���k olamaz.");
                    return;
                }
                if (TelefonNumarasiGecerliMi(telefonNo))
                {
                    MessageBox.Show("L�tfen ge�erli bir telefon numaras� giriniz.");
                    return;
                }
                if (!int.TryParse(calismaSaati, out calismaSaatiKontrol))
                {
                    MessageBox.Show("L�tfen ge�erli bir �al��ma saati giriniz.");
                    return;
                }
                //secilenKisi i�erisine bilgileri aktar�yorum,
                secilenKisi.SubItems[0].Text = ad;
                secilenKisi.SubItems[1].Text = soyad;
                secilenKisi.SubItems[2].Text = telefonNo;
                secilenKisi.SubItems[3].Text = dogumTarihi.ToShortDateString();
                secilenKisi.SubItems[4].Text = calismaSaati;
                
                //E�er memursa,
                if (statu == "Memur")
                {
                    if (cmbMemurDerecesi.SelectedIndex == -1)
                    {
                        MessageBox.Show("Memur derecesi se�iniz.");
                        return;//derecesi se�ili mi diye kontrol ediyorum.
                    }
                    //Dereceyi belirliyorum
                    int derece = cmbMemurDerecesi.SelectedIndex + 1;
                    secilenKisi.SubItems[5].Text = derece.ToString();
                    //Ve g�ncelliyorum.
                    MessageBox.Show("Memur Bilgileri G�ncellendi.");

                }
                //Y�netici ise,
                if (statu == "Y�netici")
                {
                    if (string.IsNullOrWhiteSpace(txtYoneticiSaatlikUcret.Text) || string.IsNullOrWhiteSpace(txtYoneticiBonusu.Text))
                    {
                        MessageBox.Show("Y�netici saatlik �cretini veya bonus �deme bilgisini doldurunuz.");
                        return;//Tekrardan saatlik �cretini ve bonusunu girip girmedi�ini kontrol ediyorum.
                    }
                    if (!decimal.TryParse(txtYoneticiSaatlikUcret.Text, out yoneticiSaatlikUcret) || !decimal.TryParse(txtYoneticiBonusu.Text, out bonus))
                    {
                        MessageBox.Show("L�tfen y�netici saatlik �cretini veya bonusunu ge�erli formatta giriniz. (�rne�in=500.5)");
                        return;//Format�n� kontrol ediyorum.
                    }
                    secilenKisi.SubItems[7].Text = yoneticiSaatlikUcret.ToString("C");
                    secilenKisi.SubItems[5].Text = bonus.ToString();
                    //ve ekliyorum.
                    MessageBox.Show("Y�netici Bilgileri G�ncellendi.");
                }

            }
            else
            {
                //Personel se�ili de�ilse gerekli uyar�y� yapyorum.
                MessageBox.Show("L�tfen bilgilerini g�ncellemek istedi�iniz personeli se�iniz.");
                return;
            }

        }
        #endregion


        #region Sil click event'i
        private void btnSil_Click(object sender, EventArgs e)
        {
            //lstvBilgiCRUD'tan herhangi bir item se�tiyse, yani ki�i se�tiyse,
            if (lstvBilgiCRUD.SelectedItems.Count > 0)
            {
                //Se�ilen item'� al�yorum,
                var secilenKisi = lstvBilgiCRUD.SelectedItems[0];
                var personelAdi = secilenKisi.SubItems[0].Text;
                var personalSoyad = secilenKisi.SubItems[1].Text;
                //Burada kullan�c�ya silmek isteyip istemedi�ini soruyorum tekrar.
                DialogResult result = MessageBox.Show($"{personelAdi} {personalSoyad} adl� ki�iyi silmek istiyor musunuz?", "Evet", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.OK)
                {
                    //E�er silmemi istiyorsa ilk lstvBilgiCRUD i�erisinden kald�r�yorum
                    lstvBilgiCRUD.Items.Remove(secilenKisi);

                    //E�er bilgiler e�le�iyorsa
                    var personel = personellerListesi.FirstOrDefault(p => p.Ad == personelAdi && p.Soyad == personalSoyad);
                    //ve personel null yani bo� de�ilse,
                    if (personel != null)
                    {
                        //liste i�erisinden de kald�r�yorum
                        personellerListesi.Remove(personel);

                        if (personel is Memur)
                        {
                            //ayn� i�lemi memurlar listesi i�in ve
                            memurlarListesi.Remove(personel as Memur);
                        }
                        else if (personel is Yonetici)
                        {
                            //y�netici listesi i�in yap�yorum.
                            yoneticilerListesi.Remove(personel as Yonetici);
                        }
                        MessageBox.Show("Personel ba�ar�yla silindi.");
                    }
                }
            }
            else
            {
                //Olas� hata i�in buradan mb ile mesaj g�nderiyorum.
                MessageBox.Show("L�tfen silmek i�in personel se�in.");
                return;
            }

        }
        #endregion 


        private void btnJsonOku_Click(object sender, EventArgs e)
        {
            JsonVeriOku();

        }


        #region JSON verilerini oku metodu
        private void JsonVeriOku()
        {
            //Olas� hatalar� yakalamak i�in tekrar try-catch i�erisine al�yorum,
            try
            {
                //Projenin oldu�u lokasyonu al�yorum, dosya ad�n� ve lokasyonunu giriyorum ve okunacak verinin ad�n� giriyorum.
                string projeDizi = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string hedefDizi = Path.Combine(projeDizi, @"..\..\..\", "JSON");
                string dosyaYolu = Path.Combine(hedefDizi, "personelVerileri.json");
                //e�er dosya bulunuyorsa,
                if (File.Exists(dosyaYolu))
                {
                    //T�m dosyay� okuyorum
                    string jsonVeri = File.ReadAllText(dosyaYolu);
                    //ve json verilerini deserialize ediyorum.
                    var veriListesi = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(jsonVeri);
                    //Verilistesi bo� de�ilse,
                    if (veriListesi != null)
                    {
                        //lstvBilgiCRUD ve di�er listview'daki bilgileri temizliyorum.
                        lstvBilgiCRUD.Items.Clear();
                        lstvMemurlar.Items.Clear();
                        lstvYoneticiler.Items.Clear();
                        //Verilistesindeki sat�rlar� tek tek geziyorum,
                        foreach (var satir in veriListesi)
                        {
                            //Stat�'y� al�yorum,
                            string statu = satir["Statu"];
                            ListViewItem lstv = new ListViewItem(new string[]
                            {
                                //Ekleyece�im listview item'� olu�turup bilgileri dolduruyorum.
                                satir["Ad"],
                                satir["Soyad"],
                                statu,
                                satir["Telefon No"],
                                satir["Dogum Tarihi"],
                                "",
                                satir["Calisma Saati"],
                                satir["Maas"]

                            });
                            //Sonra okudu�um bilgileri ana listview'a g�nderiyorum.
                            lstvBilgiCRUD.Items.Add(lstv);
                            //E�er memursa,
                            if (statu == "Memur")
                            {
                                ListViewItem lstvMemur = new ListViewItem(new string[]
                                {
                                    //Ayn� �ekilde listview item olu�turup bilgileri dolduruyorum.
                                    satir["Ad"],
                                    satir["Soyad"],
                                    statu,
                                    satir["Telefon No"],
                                    satir["Dogum Tarihi"],
                                    satir["Calisma Saati"],
                                    satir["Maas"],

                                });
                                //Memurlar listview'�na g�nderiyorum.
                                lstvMemurlar.Items.Add(lstvMemur);
                            }
                            //E�er Y�netici ise,
                            else if (statu == "Yonetici")
                            {
                                ListViewItem lstvYonetici = new ListViewItem(new string[]
                                {
                                    //Y�neticiler listview'�na g�ndermek i�in listview item olu�turuyorum ve bilgileri dolduruyorum.
                                    satir["Ad"],
                                    satir["Soyad"],
                                    statu,
                                    satir["Telefon No"],
                                    satir["Saatlik Ucret"],
                                    satir["Calisma Saati"],
                                    satir["Bonus"],
                                    satir["Maas"],
                                });
                                //Sonda ekliyorum.
                                lstvYoneticiler.Items.Add(lstvYonetici);
                            }

                        }
                        MessageBox.Show("Veriler ba�ar�yla y�klendi.");
                    }
                    else
                    {
                        //E�er JSON verisi bo�sa bilgilendiriyorum.
                        MessageBox.Show("Veri bulunamad�.");
                        return;
                    }
                }

            }
            catch (Exception ex)
            {
                //Olas� bir hatay� burada yakalay�p, hatay� detayland�r�yorum.
                MessageBox.Show("Bir hata olu�tu." + ex.Message);
                return;
            }

        }
        #endregion


        #region Az �al��anlar�n raporunu ald���m method
        private void AzCalisanlarExcel()
        {
            //Olas� hatalar� yakalamak i�in try-catch blo�u a��yorum.
            try
            {
               //Kaynaklar� y�netebilmek i�in tekrardan using kullan�yorum.
                using (var workbook = new XLWorkbook())
                {
                    //WorkSheet'in ad�n� belirliyorum.
                    var workSheet = workbook.AddWorksheet("AzCalisanPersonelRaporu");

                    //Ve ba�l�klar� olu�turuyorum.
                    workSheet.Cell(1, 1).Value = "Ad";
                    workSheet.Cell(1, 2).Value = "Soyad";
                    workSheet.Cell(1, 3).Value = "Stat�";
                    workSheet.Cell(1, 4).Value = "Telefon No";
                    workSheet.Cell(1, 4).Value = "Ki�isel Bilgileri";
                    workSheet.Cell(1, 6).Value = "�al��ma Saati";
                    workSheet.Cell(1, 7).Value = "Maa�";
                    workSheet.Cell(1, 8).Value = "Memur Derecesi";

                    //Burada tekar sat�r� girmem gerekiyor ��nk� ilk sat�rda ba�l�klar yer al�yor.
                    int satir = 2;

                    //Az �al��anlar i�in listviewitem t�r�nde bir list olu�turuyorum.
                    List<ListViewItem> azCalisanlar = new List<ListViewItem>();

                    //Memurlar listview'�n� geziyorum
                    foreach (ListViewItem item in lstvMemurlar.Items)
                    {
                        //�al��ma saatini al�p burada TryParse'l�yorum, t�r�n d�n���p d�n��emedi�ini bilmem gerek.
                        int calismaSaati;
                        if (int.TryParse(item.SubItems[5].Text, out calismaSaati) && calismaSaati < 150)
                        {
                            //E�er d�n���yorsa azcalisanlar listesine ekliyorum. 
                            azCalisanlar.Add(item);
                        }

                    }
                    //Ayn� i�lemi y�netici listview'� i�in de yap�yorum.
                    foreach (ListViewItem item in lstvYoneticiler.Items)
                    {
                        //calismasaati'nin d�n���p d�n��medi�ine tekrar bak�yorum.
                        int calismaSaati;
                        if (int.TryParse(item.SubItems[5].Text, out calismaSaati) && calismaSaati < 150)
                        {
                            //D�n���yorsa ekliyorum.
                            azCalisanlar.Add(item);
                        }

                    }
                    //e�er az �al��an say�s� 0'dan b�y�kse, yani varsa;
                    if (azCalisanlar.Count > 0)
                    {
   
                        //Az �al��anlar listesinin item'lar�n� tek tek geziyorum.
                        foreach (ListViewItem item in azCalisanlar)
                        {
                            //Ve excel i�erisindeki cell'e ekliyorum.
                            workSheet.Cell(satir, 1).Value = item.SubItems[0].Text;
                            workSheet.Cell(satir, 2).Value = item.SubItems[1].Text;
                            workSheet.Cell(satir, 3).Value = item.SubItems[2].Text;
                            workSheet.Cell(satir, 4).Value = item.SubItems[3].Text;
                            workSheet.Cell(satir, 5).Value = item.SubItems[4].Text;
                            workSheet.Cell(satir, 6).Value = item.SubItems[5].Text;
                            workSheet.Cell(satir, 7).Value = item.SubItems[6].Text;
                            workSheet.Cell(satir, 8).Value = item.SubItems.Count > 7 ? item.SubItems[7].Text : "";
                            //Ekledikten sonra sat�r� artt�r�yorum ki �st�ne yazmas�n
                            satir++;

                        }
                    }
                    else
                    {
                        //E�er az �al��an personel yoksa excel i�erisine bunu yazd�raca��m.
                        workSheet.Cell(satir, 1).Value = "150 Saatten az �al��an bulunamad�.";
                    }
                    //Kaynaklar� kontrol edebilmem i�in tekrardan using blo�una al�yorum,
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        //t�r�n�, kaydetme ad� gibi �zellikleri belirtiyorum.
                        saveFileDialog.Filter = "Excel Files |*.xlsx";
                        saveFileDialog.Title = "Excel Dosyas�n� Kaydet";
                        saveFileDialog.FileName = "AzCalisanPersonellerinRaporu.xlsx";
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            //E�er OK'a bast�ysa kaydediyorum.
                            string dosyaYolu = saveFileDialog.FileName;
                            workbook.SaveAs(dosyaYolu);
                            MessageBox.Show("Ba�ar�yla Kaydedildi!");
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                //Yakalad���m hatalar� burada tutup detayl� mesaj ile g�steriyorum.
                MessageBox.Show("Bir hata olu�tu" + ex.Message);
            }

        }
        #endregion

        private void btnAzCalisan_Click(object sender, EventArgs e)
        {
            AzCalisanlarExcel();
        }
    }
}
