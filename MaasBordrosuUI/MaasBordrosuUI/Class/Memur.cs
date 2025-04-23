using MaasBordrosuUI.Class.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MaasBordrosuUI.Class
{
    internal class Memur:Personel
    {
        private decimal mesaiUcreti;
        public int Derece { get; set; }
        public decimal Maas { get; set; }
        public decimal MesaiUcreti { get; set; }
        public decimal SaatlikUcret

        {
            get
            {
                switch (Derece)
                {
                    case 1:
                        return 500m;
                    case 2:
                        return 600m;
                    case 3:
                        return 700m;
                    default:
                        return 500m;
                }
            }

        }
        public Memur(string ad,string soyad, string telno, DateTime dogumtarihi, int calismasaati,int derece) : base(ad,soyad, telno, dogumtarihi, calismasaati)
        {
            Derece = derece;
            Maas = MaasHesapla();
            MesaiUcreti = mesaiUcreti;
        }
        public override decimal MaasHesapla()
        {
            if (CalismaSaati > 180)
            {
                int mesaiSaati = CalismaSaati-180;
                decimal mesaiUcreti = mesaiSaati * 1.5m;
                decimal normalMaas = 180 * SaatlikUcret;
                decimal Maas = (mesaiSaati * SaatlikUcret) * 1.5m + normalMaas;
                return Maas;
            }
            else
            {
                decimal normalMaas = CalismaSaati * SaatlikUcret;
                return normalMaas;
            }
            
        }
    }
}
