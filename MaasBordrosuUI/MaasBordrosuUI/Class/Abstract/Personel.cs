using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MaasBordrosuUI.Class.Abstract
{
    public abstract class Personel
    {
        public string Ad { get; set; }
        public string Soyad { get; set; }
        public string TelefonNo { get; set; }
        public DateTime DogumTarihi { get; set; }
        public int CalismaSaati { get; set; }

        protected Personel(string ad, string soyad,string telno,DateTime dogumtarihi,int calismasaati)
        {
            Ad = ad;
            Soyad = soyad;
            TelefonNo = telno;
            DogumTarihi = dogumtarihi;
            CalismaSaati= calismasaati;
                
        }

        public abstract decimal MaasHesapla();
    }
}
