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
            lstvBilgiCRUD.Columns.Add("Statü", 100);
            lstvBilgiCRUD.Columns.Add("Telefon Numarasý", 150);
            lstvBilgiCRUD.Columns.Add("Doðum Tarihi", 80);
            lstvBilgiCRUD.Columns.Add("Memur Derecesi(Memur ise)", 50);
            lstvBilgiCRUD.Columns.Add("Çalýþma saati", 80);
            lstvBilgiCRUD.Columns.Add("Maaþ", 100);
            cmbMemurDerecesi.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbStatu.DropDownStyle = ComboBoxStyle.DropDownList;
            lstvBilgiCRUD.FullRowSelect = true;



            lstvMemurlar.Columns.Add("Ad", 80);
            lstvMemurlar.Columns.Add("Soyad", 80);
            lstvMemurlar.Columns.Add("Statü", 80);
            lstvMemurlar.Columns.Add("Telefon No", 100);
            lstvMemurlar.Columns.Add("Doðum Tarihi", 100);
            lstvMemurlar.Columns.Add("Çalýþma Saati", 120);
            lstvMemurlar.Columns.Add("Maaþ", 100);
            lstvMemurlar.Columns.Add("Mesai Ücreti", 80);
            lstvMemurlar.Columns.Add("Memur Derecesi", 100);


            lstvYoneticiler.Columns.Add("Ad", 100);
            lstvYoneticiler.Columns.Add("Soyad", 100);
            lstvYoneticiler.Columns.Add("Statü", 100);
            lstvYoneticiler.Columns.Add("Telefon No", 100);
            lstvYoneticiler.Columns.Add("Saatlik Ücret", 130);
            lstvYoneticiler.Columns.Add("Çalýþma Saati", 150);
            lstvYoneticiler.Columns.Add("Bonus", 80);
            lstvYoneticiler.Columns.Add("Maaþ", 100);
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


        #region Yaptýðým Desing için Visible ve label text düzenlemesi

        //Desing'ý destekleyen birkaç olay.
        //Form yerine panel kullanýp farklý bir proje ortaya çýkarmak istediðim için panellerin görünürlüðünü button'larýn click event'i ile deðiþtiriyorum.
        //Form ekranýný kapatacak veya form ekranýný aþaðý atacak button'u kendim tasarladýðým için iþlevini yerine getirebilmesi için birkaç kod bloðu ekledim.

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
            lblSayfaBilgisi.Text = "Personel Yönetimi";
            panelGiris.Visible = false;
            panelDokumantasyon.Visible = false;
            panelPersonelYonetimi.Visible = true;
            panelRaporlar.Visible = false;
        }

        private void btnDokumantasyon_Click(object sender, EventArgs e)
        {
            lblSayfaBilgisi.Visible = true;
            panelDokumantasyon.Visible = true;
            lblSayfaBilgisi.Text = "Dokümantasyon";
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
            //TextBox'lardan, Combobox'lar, DateTimePicker'dan kiþinin bilgilerini alýyorum,
            string ad = txtAd.Text;
            string soyad = txtSoyad.Text;
            string telefonNo = mtbTelNo.Text;
            DateTime dogumTarihi = dtpDogumTarihi.Value;
            int personelYasi = DateTime.Now.Year - dogumTarihi.Year;
            string calismaSaati = txtCalismaSaati.Text;

            int calismaSaatiKontrol;
            decimal yoneticiSaatlikUcret;
            decimal bonus = 0;

            //Burada statü bilgisini de alýp onda göre olasý hatalarý önceden engelliyorum,
            string statu = cmbStatu.Text;
            if (string.IsNullOrWhiteSpace(statu))
            {
                MessageBox.Show("Lütfen personelin statüsünü seçiniz.");
                return;//Mesela statü Combobox'ýný boþ býraktýðýnda alacaðý hata vb.

            }
            if (string.IsNullOrWhiteSpace(ad) || string.IsNullOrWhiteSpace(soyad))
            {
                MessageBox.Show("Ad ve Soyad kýsmý boþ býrakýlamaz.");
                return;
            }
            if (!HarfKontrol(ad) || !HarfKontrol(soyad))
            {
                MessageBox.Show("Lütfen ad ve soyad kýsmýný istenilen formatta giriniz.");
                return;//Ad ve soyad Textbox'larýna harf harici baþka bir þey girdiðinde alacaðý hata.
            }
            if (personelYasi < 18)
            {
                MessageBox.Show("Personel yaþý 18'den küçük olamaz.");
                return;
            }
            if (TelefonNumarasiGecerliMi(telefonNo))
            {
                MessageBox.Show("Lütfen geçerli bir telefon numarasý giriniz.");
                return;
            }
            if (!int.TryParse(calismaSaati, out calismaSaatiKontrol))
            {
                MessageBox.Show("Lütfen geçerli bir çalýþma saati giriniz.");
                return;
            }

            //Eðer ekleyeceði kiþi memur ise burada ona göre iþlemler yapýyorum,
            if (statu == "Memur")
            {
                if (cmbMemurDerecesi.SelectedIndex == -1)
                {
                    MessageBox.Show("Memur derecesi seçiniz.");
                    return;//Statü memur ise memur derecesi seçmesi gerektiðini mb ile söylüyorum.
                }
                int derece = cmbMemurDerecesi.SelectedIndex + 1;
                //Class'larýmda oluþturduðum constructor'a bilgi aktarýmý yapýyorum,
                Memur memur = new Memur(ad, soyad, telefonNo, dogumTarihi, calismaSaatiKontrol, derece);
                //Personel olduðu için ve memur olduðu için personel ve memur dosyasýna ekliyorum.
                personellerListesi.Add(memur);
                memurlarListesi.Add(memur);
                MessageBox.Show("Memur Personel Baþarýyla Eklendi.");
                //Personel eklediðimiz sayfadaki Listview içerisine bilgi aktarýmý yapmak için burada Listview bir deðiþken oluþturup bilgileri gönderiyorum,
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
                //Gözükmesi için de ekliyorum.
                lstvBilgiCRUD.Items.Add(memurLSTV);
                //Dokümantasyondaki Memur listesine ekleyebilmek için tekrar oluþturuyorum çünkü bir nesne birden fazla Listview'a eklenemez.
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
                //Ayný þekilde burada da Listview'a ekliyorum.
                lstvMemurlar.Items.Add(memurlarListview);

            }
            //Eðer ekleyeceði personel Yönetici ise,
            if (statu == "Yönetici")
            {
                if (string.IsNullOrWhiteSpace(txtYoneticiSaatlikUcret.Text) || string.IsNullOrWhiteSpace(txtYoneticiBonusu.Text))
                {
                    MessageBox.Show("Yönetici saatlik ücretini veya bonus ödeme bilgisini doldurunuz.");
                    return;//Bonusu ve saatlik ücreti girmek zorunda olduðu için burada önlem alýyorum.
                }
                if (!decimal.TryParse(txtYoneticiSaatlikUcret.Text, out yoneticiSaatlikUcret) || !decimal.TryParse(txtYoneticiBonusu.Text, out bonus))
                {
                    MessageBox.Show("Lütfen yönetici saatlik ücretini veya bonusunu geçerli formatta giriniz. (Örneðin=500.5)");
                    return;//Yöneticinin saatlik ücretini geçerli formatta girip girmediðini kontrol ediyorum, sadece int deðerleri kabul ediyorum.
                }
                if (Convert.ToDecimal(txtYoneticiSaatlikUcret) < 500)
                {
                    MessageBox.Show("Yöneticinin saatlik ücreti 500'den küçük olamaz!");
                    return;//Ýsterlerde yazan saatlik ücretin 500'den küçük olamayacaðýydý, onun kontrolünü yapýyorum.
                }
                //Listeye aktarým yapabilmek için Yönetici sýnýfýnýn(Miras aldýðý personel sýnýfý ayrýca) constructor'una bilgi gönderiyorum
                Yonetici yonetici = new Yonetici(ad, soyad, telefonNo, dogumTarihi, calismaSaatiKontrol, yoneticiSaatlikUcret, bonus);
                //Personel olduðu için ve Yönetici olduðu için listeye ekliyorum.
                personellerListesi.Add(yonetici);
                yoneticilerListesi.Add(yonetici);
                MessageBox.Show("Yönetici Personel Baþarýyla Eklendi.");
                //Personel eklediðimiz sayfadaki Listview içerisine bilgi aktarýmý yapmak için burada Listview bir deðiþken oluþturup bilgileri gönderiyorum,
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
                //Dokümantasyondaki Listview'a atamak için tekrar oluþturuyorum.
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
                //Sonra gönderiyorum.
                lstvBilgiCRUD.Items.Add(yoneticiLSTV);
                lstvYoneticiler.Items.Add(yoneticiListview);

            }
        }
        #endregion


        #region Telefon numarasý formatý kontrol
        public bool TelefonNumarasiGecerliMi(string telefonNumarasi)
        {
            //Girilen formada uygun olup olmadýðýný kontrol etmek için regex oluþturuyorum.
            //Bu formak xxx-xxx-xxxx olarak int deðerleri kabul edecek.
            var reg = new System.Text.RegularExpressions.Regex(@"^5\d{2} \d{3} \d{4}$");
            return reg.IsMatch(telefonNumarasi);
        }
        #endregion


        #region Harf kontrol 
        public bool HarfKontrol(string metin)
        {
            //Ýstediðim formatta regex oluþturuyorum.
            //Sadece alfabetik harflari küçük ve büyük formatta kabul edecek.
            var reg = new System.Text.RegularExpressions.Regex(@"^[a-zA-ZÇçÐðÝýÖöÞþÜü]+$");
            return reg.IsMatch(metin);
        }
        #endregion


        #region Combobox'taki seçili deðer kontrol
        private void cmbStatu_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Burada oluþturduðum Desing'dan kaynaklý Enable'ýný deðiþtirmem gereken iþlemler vardý onlarý gerçekleþtiriyorum.
            string statu = cmbStatu.SelectedItem.ToString();
            if (statu == "Memur")
            {
                //Çünkü Memur bir kiþi ekleyecekse Yönetici saatlik ücretine ve bonusa eriþmemeli
                cmbMemurDerecesi.Enabled = true;
                txtYoneticiSaatlikUcret.Enabled = false;
                txtYoneticiBonusu.Enabled = false;
            }
            if (statu == "Yönetici")
            {
                //Burada da eðer Yöneticiyse Memur derecesine eriþememeli
                txtYoneticiSaatlikUcret.Enabled = true;
                txtYoneticiBonusu.Enabled = true;
                cmbMemurDerecesi.Enabled = false;
            }
        }
        #endregion


        #region JSON formatýnda verileri kaydetme
        private void JsonVeriKaydet()
        {
            //Herhangi bir hatayý yakalamak için try-catch buloðuna alýyorum,
            try
            {
                //Projenin olduðu lokasyonu ve Proje içerisine oluþacak klasörün yolunu belirliyorum. Sonrasýnda hedefdizi içerisine hangi isimle kaydedileceðini belirliyorum.
                string projeDizi = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string HedefDizi = Path.Combine(projeDizi, @"..\..\..\", "JSON");
                string dosyaYolu = Path.Combine(HedefDizi, "personelVerileri.json");

                if (!Directory.Exists(HedefDizi))
                {
                    Directory.CreateDirectory(HedefDizi);
                }
                //iki deðerli tutmak istediðim için dictionary bir deðiþken listesi oluþturuyorum.
                List<Dictionary<string, string>> mevcutVeriler = new List<Dictionary<string, string>>();
                if (File.Exists(dosyaYolu))
                {
                    //Eðer daha önce oluþturulmuþ bir JSON varda tüm dosyayý okuyorum ve 
                    //Eðer boþ deðilse yeni JSON verilerini üstüne yazdýrýyorum.
                    string eskiJson = File.ReadAllText(dosyaYolu);
                    var veriListesi = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(eskiJson);
                    if (veriListesi != null)
                    {
                        mevcutVeriler = veriListesi;
                    }
                }

                //lstvBilgiCRUD listesi içerisindeki her item'ý geziyorum
                foreach (ListViewItem item in lstvBilgiCRUD.Items)
                {
                    Dictionary<string, string> satir = new Dictionary<string, string>
                    {
                        //dictionary'nin içerisindeki ilk string deðeri kendim verip ikinci string deðerini ise item'daki subitems'dan alýyorum.
                        {"Ad",item.SubItems[0].Text},
                        {"Soyad",item.SubItems[1].Text},
                        {"Statu",item.SubItems[2].Text},
                        {"Telefon No",item.SubItems[3].Text},
                        {"Dogum Tarihi",item.SubItems[4].Text},
                    };
                    //Eðer memur ise ona göre iþlem yapmak için subitem'daki 3. öðeyi kontrol ediyorum çünkü 3. öðe statü'yü belirliyor.
                    if (item.SubItems[2].Text == "Memur")
                    {
                        satir.Add("Memur derecesi", item.SubItems[5].Text);
                        satir.Add("Calisma Saati", item.SubItems[6].Text);
                        satir.Add("Maas", item.SubItems[7].Text);
                    }
                    //Aynýsýný Yönetici için yapýyorum.
                    if (item.SubItems[2].Text == "Yonetici")
                    {
                        foreach (ListViewItem yoneticiItem in lstvYoneticiler.Items)
                        {
                            //Burada Yönetici listesinin içinden detaylý bilgi çekiyorum.
                            if (yoneticiItem.SubItems[0].Text == item.SubItems[0].Text && yoneticiItem.SubItems[1].Text == item.SubItems[1].Text)
                            {
                                satir.Add("Saatlik Ucret", yoneticiItem.SubItems[4].Text);
                                satir.Add("Bonus", yoneticiItem.SubItems[5].Text);
                                satir.Add("Calisma Saati", yoneticiItem.SubItems[6].Text);
                                satir.Add("Maas", yoneticiItem.SubItems[7].Text);
                            }

                        }
                    }
                    //Mevcut verilerin içerisine satýrý ekliyorum.
                    mevcutVeriler.Add(satir);
                }
                //Burada satýr satýr olmasý için JSON ayarlarýný yapýyorum.
                var jsonAyar = new JsonSerializerOptions { WriteIndented = true };
                string yeniJson = JsonSerializer.Serialize(mevcutVeriler, jsonAyar);
                File.WriteAllText(dosyaYolu, yeniJson);
                MessageBox.Show("Veriler .JSON formatýnda kaydedildi.");


            }
            catch (Exception ex)
            {
                //Hatalarý yakalýyorum.
                MessageBox.Show("Bir hata oluþtu.");
            }
        }
        #endregion


        #region Excel oluþtur metodu
        private void ExcelOlustur()
        {
            try
            {
                //Using kullanýyorum çünkü kaynaklarý otomatik olarak yönetmem lazým.
                using (var workbook = new XLWorkbook())
                {
                    //Bu kýsýmda baþlýklarý belirliyorum.
                    var workSheet = workbook.AddWorksheet("PersonelRaporu");
                    workSheet.Cell(1, 1).Value = "Ad";
                    workSheet.Cell(1, 2).Value = "Soyad";
                    workSheet.Cell(1, 3).Value = "Statü";
                    workSheet.Cell(1, 4).Value = "Telefon No";
                    workSheet.Cell(1, 7).Value = "Çalýþma Saati";
                    workSheet.Cell(1, 8).Value = "Maaþ";
                    //Satýr belirtmem lazým çünkü ilk satýrda baþlýklar var.
                    int satir = 2;

                    //Memur listesini gezip içerisindeki verileri oluþturduðum excel içerisine gönderiyorum,
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
                    //Ayný iþlemi Yönetici listesi için yapýyorum.
                    foreach (ListViewItem item in lstvYoneticiler.Items)
                    {
                        workSheet.Cell(satir, 1).Value = item.SubItems[0].Text;
                        workSheet.Cell(satir, 2).Value = item.SubItems[1].Text;
                        workSheet.Cell(satir, 3).Value = "Yönetici";
                        workSheet.Cell(satir, 4).Value = item.SubItems[3].Text;
                        workSheet.Cell(satir, 7).Value = item.SubItems[6].Text;
                        workSheet.Cell(satir, 8).Value = item.SubItems[7].Text;
                        satir++;
                    }
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        //Burada format, dosya adý gibi özellikleri belirliyorum.
                        saveFileDialog.Filter = "Excel Files |*.xlsx";
                        saveFileDialog.Title = "Excel Dosyasýný Kaydet";
                        saveFileDialog.FileName = "PersonelRaporu.xlsx";
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string dosyaYolu = saveFileDialog.FileName;
                            workbook.SaveAs(dosyaYolu);
                            MessageBox.Show("Excel Baþarýyla Oluþturuldu.");
                        }
                    }

                }

            }
            catch (Exception)
            {
                //Olasý hatalarý yakalýyorum.
                MessageBox.Show("Bir Sorun Oluþtu.");
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


        #region Güncelle click event'i
        private void btnGuncelle_Click(object sender, EventArgs e)
        {
            //Eðer personel seçiliyse:
            if (lstvBilgiCRUD.SelectedItems.Count > 0)
            {
                //Bilgilerin hepsini tekrar alýyorum,
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

                //Kontrolleri yapýp statü'ye göre tekrar iþlem yapacaðým.
                //Statü seçili mi kontrolü, ad soyad boþ mu? boþ deðilse istediðim formatta mý? personelin yaþý 18'den büyük mü? telefon numarasý istediðim formatta mý? gibi kontrolleri yapýyorum.
                string statu = cmbStatu.Text;
                if (string.IsNullOrWhiteSpace(statu))
                {
                    MessageBox.Show("Lütfen personelin statüsünü seçiniz.");
                    return;

                }
                if (string.IsNullOrWhiteSpace(ad) || string.IsNullOrWhiteSpace(soyad))
                {
                    MessageBox.Show("Ad ve Soyad kýsmý boþ býrakýlamaz.");
                    return;
                }
                if (!HarfKontrol(ad) || !HarfKontrol(soyad))
                {
                    MessageBox.Show("Lütfen ad ve soyad kýsmýný istenilen formatta giriniz.");
                    return;
                }
                if (personelYasi < 18)
                {
                    MessageBox.Show("Personel yaþý 18'den küçük olamaz.");
                    return;
                }
                if (TelefonNumarasiGecerliMi(telefonNo))
                {
                    MessageBox.Show("Lütfen geçerli bir telefon numarasý giriniz.");
                    return;
                }
                if (!int.TryParse(calismaSaati, out calismaSaatiKontrol))
                {
                    MessageBox.Show("Lütfen geçerli bir çalýþma saati giriniz.");
                    return;
                }
                //secilenKisi içerisine bilgileri aktarýyorum,
                secilenKisi.SubItems[0].Text = ad;
                secilenKisi.SubItems[1].Text = soyad;
                secilenKisi.SubItems[2].Text = telefonNo;
                secilenKisi.SubItems[3].Text = dogumTarihi.ToShortDateString();
                secilenKisi.SubItems[4].Text = calismaSaati;
                
                //Eðer memursa,
                if (statu == "Memur")
                {
                    if (cmbMemurDerecesi.SelectedIndex == -1)
                    {
                        MessageBox.Show("Memur derecesi seçiniz.");
                        return;//derecesi seçili mi diye kontrol ediyorum.
                    }
                    //Dereceyi belirliyorum
                    int derece = cmbMemurDerecesi.SelectedIndex + 1;
                    secilenKisi.SubItems[5].Text = derece.ToString();
                    //Ve güncelliyorum.
                    MessageBox.Show("Memur Bilgileri Güncellendi.");

                }
                //Yönetici ise,
                if (statu == "Yönetici")
                {
                    if (string.IsNullOrWhiteSpace(txtYoneticiSaatlikUcret.Text) || string.IsNullOrWhiteSpace(txtYoneticiBonusu.Text))
                    {
                        MessageBox.Show("Yönetici saatlik ücretini veya bonus ödeme bilgisini doldurunuz.");
                        return;//Tekrardan saatlik ücretini ve bonusunu girip girmediðini kontrol ediyorum.
                    }
                    if (!decimal.TryParse(txtYoneticiSaatlikUcret.Text, out yoneticiSaatlikUcret) || !decimal.TryParse(txtYoneticiBonusu.Text, out bonus))
                    {
                        MessageBox.Show("Lütfen yönetici saatlik ücretini veya bonusunu geçerli formatta giriniz. (Örneðin=500.5)");
                        return;//Formatýný kontrol ediyorum.
                    }
                    secilenKisi.SubItems[7].Text = yoneticiSaatlikUcret.ToString("C");
                    secilenKisi.SubItems[5].Text = bonus.ToString();
                    //ve ekliyorum.
                    MessageBox.Show("Yönetici Bilgileri Güncellendi.");
                }

            }
            else
            {
                //Personel seçili deðilse gerekli uyarýyý yapyorum.
                MessageBox.Show("Lütfen bilgilerini güncellemek istediðiniz personeli seçiniz.");
                return;
            }

        }
        #endregion


        #region Sil click event'i
        private void btnSil_Click(object sender, EventArgs e)
        {
            //lstvBilgiCRUD'tan herhangi bir item seçtiyse, yani kiþi seçtiyse,
            if (lstvBilgiCRUD.SelectedItems.Count > 0)
            {
                //Seçilen item'ý alýyorum,
                var secilenKisi = lstvBilgiCRUD.SelectedItems[0];
                var personelAdi = secilenKisi.SubItems[0].Text;
                var personalSoyad = secilenKisi.SubItems[1].Text;
                //Burada kullanýcýya silmek isteyip istemediðini soruyorum tekrar.
                DialogResult result = MessageBox.Show($"{personelAdi} {personalSoyad} adlý kiþiyi silmek istiyor musunuz?", "Evet", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.OK)
                {
                    //Eðer silmemi istiyorsa ilk lstvBilgiCRUD içerisinden kaldýrýyorum
                    lstvBilgiCRUD.Items.Remove(secilenKisi);

                    //Eðer bilgiler eþleþiyorsa
                    var personel = personellerListesi.FirstOrDefault(p => p.Ad == personelAdi && p.Soyad == personalSoyad);
                    //ve personel null yani boþ deðilse,
                    if (personel != null)
                    {
                        //liste içerisinden de kaldýrýyorum
                        personellerListesi.Remove(personel);

                        if (personel is Memur)
                        {
                            //ayný iþlemi memurlar listesi için ve
                            memurlarListesi.Remove(personel as Memur);
                        }
                        else if (personel is Yonetici)
                        {
                            //yönetici listesi için yapýyorum.
                            yoneticilerListesi.Remove(personel as Yonetici);
                        }
                        MessageBox.Show("Personel baþarýyla silindi.");
                    }
                }
            }
            else
            {
                //Olasý hata için buradan mb ile mesaj gönderiyorum.
                MessageBox.Show("Lütfen silmek için personel seçin.");
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
            //Olasý hatalarý yakalamak için tekrar try-catch içerisine alýyorum,
            try
            {
                //Projenin olduðu lokasyonu alýyorum, dosya adýný ve lokasyonunu giriyorum ve okunacak verinin adýný giriyorum.
                string projeDizi = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string hedefDizi = Path.Combine(projeDizi, @"..\..\..\", "JSON");
                string dosyaYolu = Path.Combine(hedefDizi, "personelVerileri.json");
                //eðer dosya bulunuyorsa,
                if (File.Exists(dosyaYolu))
                {
                    //Tüm dosyayý okuyorum
                    string jsonVeri = File.ReadAllText(dosyaYolu);
                    //ve json verilerini deserialize ediyorum.
                    var veriListesi = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(jsonVeri);
                    //Verilistesi boþ deðilse,
                    if (veriListesi != null)
                    {
                        //lstvBilgiCRUD ve diðer listview'daki bilgileri temizliyorum.
                        lstvBilgiCRUD.Items.Clear();
                        lstvMemurlar.Items.Clear();
                        lstvYoneticiler.Items.Clear();
                        //Verilistesindeki satýrlarý tek tek geziyorum,
                        foreach (var satir in veriListesi)
                        {
                            //Statü'yü alýyorum,
                            string statu = satir["Statu"];
                            ListViewItem lstv = new ListViewItem(new string[]
                            {
                                //Ekleyeceðim listview item'ý oluþturup bilgileri dolduruyorum.
                                satir["Ad"],
                                satir["Soyad"],
                                statu,
                                satir["Telefon No"],
                                satir["Dogum Tarihi"],
                                "",
                                satir["Calisma Saati"],
                                satir["Maas"]

                            });
                            //Sonra okuduðum bilgileri ana listview'a gönderiyorum.
                            lstvBilgiCRUD.Items.Add(lstv);
                            //Eðer memursa,
                            if (statu == "Memur")
                            {
                                ListViewItem lstvMemur = new ListViewItem(new string[]
                                {
                                    //Ayný þekilde listview item oluþturup bilgileri dolduruyorum.
                                    satir["Ad"],
                                    satir["Soyad"],
                                    statu,
                                    satir["Telefon No"],
                                    satir["Dogum Tarihi"],
                                    satir["Calisma Saati"],
                                    satir["Maas"],

                                });
                                //Memurlar listview'ýna gönderiyorum.
                                lstvMemurlar.Items.Add(lstvMemur);
                            }
                            //Eðer Yönetici ise,
                            else if (statu == "Yonetici")
                            {
                                ListViewItem lstvYonetici = new ListViewItem(new string[]
                                {
                                    //Yöneticiler listview'ýna göndermek için listview item oluþturuyorum ve bilgileri dolduruyorum.
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
                        MessageBox.Show("Veriler baþarýyla yüklendi.");
                    }
                    else
                    {
                        //Eðer JSON verisi boþsa bilgilendiriyorum.
                        MessageBox.Show("Veri bulunamadý.");
                        return;
                    }
                }

            }
            catch (Exception ex)
            {
                //Olasý bir hatayý burada yakalayýp, hatayý detaylandýrýyorum.
                MessageBox.Show("Bir hata oluþtu." + ex.Message);
                return;
            }

        }
        #endregion


        #region Az çalýþanlarýn raporunu aldýðým method
        private void AzCalisanlarExcel()
        {
            //Olasý hatalarý yakalamak için try-catch bloðu açýyorum.
            try
            {
               //Kaynaklarý yönetebilmek için tekrardan using kullanýyorum.
                using (var workbook = new XLWorkbook())
                {
                    //WorkSheet'in adýný belirliyorum.
                    var workSheet = workbook.AddWorksheet("AzCalisanPersonelRaporu");

                    //Ve baþlýklarý oluþturuyorum.
                    workSheet.Cell(1, 1).Value = "Ad";
                    workSheet.Cell(1, 2).Value = "Soyad";
                    workSheet.Cell(1, 3).Value = "Statü";
                    workSheet.Cell(1, 4).Value = "Telefon No";
                    workSheet.Cell(1, 4).Value = "Kiþisel Bilgileri";
                    workSheet.Cell(1, 6).Value = "Çalýþma Saati";
                    workSheet.Cell(1, 7).Value = "Maaþ";
                    workSheet.Cell(1, 8).Value = "Memur Derecesi";

                    //Burada tekar satýrý girmem gerekiyor çünkü ilk satýrda baþlýklar yer alýyor.
                    int satir = 2;

                    //Az çalýþanlar için listviewitem türünde bir list oluþturuyorum.
                    List<ListViewItem> azCalisanlar = new List<ListViewItem>();

                    //Memurlar listview'ýný geziyorum
                    foreach (ListViewItem item in lstvMemurlar.Items)
                    {
                        //Çalýþma saatini alýp burada TryParse'lýyorum, türün dönüþüp dönüþemediðini bilmem gerek.
                        int calismaSaati;
                        if (int.TryParse(item.SubItems[5].Text, out calismaSaati) && calismaSaati < 150)
                        {
                            //Eðer dönüþüyorsa azcalisanlar listesine ekliyorum. 
                            azCalisanlar.Add(item);
                        }

                    }
                    //Ayný iþlemi yönetici listview'ý için de yapýyorum.
                    foreach (ListViewItem item in lstvYoneticiler.Items)
                    {
                        //calismasaati'nin dönüþüp dönüþmediðine tekrar bakýyorum.
                        int calismaSaati;
                        if (int.TryParse(item.SubItems[5].Text, out calismaSaati) && calismaSaati < 150)
                        {
                            //Dönüþüyorsa ekliyorum.
                            azCalisanlar.Add(item);
                        }

                    }
                    //eðer az çalýþan sayýsý 0'dan büyükse, yani varsa;
                    if (azCalisanlar.Count > 0)
                    {
   
                        //Az çalýþanlar listesinin item'larýný tek tek geziyorum.
                        foreach (ListViewItem item in azCalisanlar)
                        {
                            //Ve excel içerisindeki cell'e ekliyorum.
                            workSheet.Cell(satir, 1).Value = item.SubItems[0].Text;
                            workSheet.Cell(satir, 2).Value = item.SubItems[1].Text;
                            workSheet.Cell(satir, 3).Value = item.SubItems[2].Text;
                            workSheet.Cell(satir, 4).Value = item.SubItems[3].Text;
                            workSheet.Cell(satir, 5).Value = item.SubItems[4].Text;
                            workSheet.Cell(satir, 6).Value = item.SubItems[5].Text;
                            workSheet.Cell(satir, 7).Value = item.SubItems[6].Text;
                            workSheet.Cell(satir, 8).Value = item.SubItems.Count > 7 ? item.SubItems[7].Text : "";
                            //Ekledikten sonra satýrý arttýrýyorum ki üstüne yazmasýn
                            satir++;

                        }
                    }
                    else
                    {
                        //Eðer az çalýþan personel yoksa excel içerisine bunu yazdýracaðým.
                        workSheet.Cell(satir, 1).Value = "150 Saatten az çalýþan bulunamadý.";
                    }
                    //Kaynaklarý kontrol edebilmem için tekrardan using bloðuna alýyorum,
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        //türünü, kaydetme adý gibi özellikleri belirtiyorum.
                        saveFileDialog.Filter = "Excel Files |*.xlsx";
                        saveFileDialog.Title = "Excel Dosyasýný Kaydet";
                        saveFileDialog.FileName = "AzCalisanPersonellerinRaporu.xlsx";
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            //Eðer OK'a bastýysa kaydediyorum.
                            string dosyaYolu = saveFileDialog.FileName;
                            workbook.SaveAs(dosyaYolu);
                            MessageBox.Show("Baþarýyla Kaydedildi!");
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                //Yakaladýðým hatalarý burada tutup detaylý mesaj ile gösteriyorum.
                MessageBox.Show("Bir hata oluþtu" + ex.Message);
            }

        }
        #endregion

        private void btnAzCalisan_Click(object sender, EventArgs e)
        {
            AzCalisanlarExcel();
        }
    }
}
