using MaasBordrosuUI.Class.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MaasBordrosuUI.Class
{
    internal class Yonetici:Personel
    {
        public decimal SaatlikUcret { get; set; }
        public decimal Bonus { get; set; }
        public decimal Maas { get; set; }

        public Yonetici(string ad, string soyad, string telno, DateTime dogumtarihi, int calismasaati, decimal saatlikucret,decimal bonus): base(ad, soyad, telno, dogumtarihi,calismasaati)
        {
            SaatlikUcret = saatlikucret;
            Bonus = bonus;
            Maas = MaasHesapla();
                
        }

        public override decimal MaasHesapla()
        {
            if (SaatlikUcret < 500)
            {
                throw new Exception("Yöneticinin saatlik ücreti 500'den küçük olamaz.");
            }
            else
            {
                decimal maas = (SaatlikUcret * CalismaSaati) + Bonus;
                return maas;
            }
        }
    }
}
