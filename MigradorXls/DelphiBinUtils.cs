using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigradorXls
{
    class DelphiBinUtils
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public DelphiBinUtils()
        {
        }
        /// <summary>
        /// Metodo que obtiene el valor flotante del rango del bytes
        /// </summary>
        public static decimal getDecFromCurr(byte[] binaryData, int address)
        {
            return decimal.Divide(new decimal(BitConverter.ToInt64(binaryData, address)), new decimal(10000L));
        }
        /// <summary>
        /// Metodo que obtiene el string de rango de byte usando codificacion  ASCII
        /// </summary>
        public static string getStrFromStr(byte[] binaryData, int start, int length)
        {
            ASCIIEncoding enc = new ASCIIEncoding();
            return enc.GetString(binaryData, start, length);
        }
        /// <summary>
        /// Metodo que obtiene bool en base a parametros de entrada
        /// </summary>
        public static bool getBoolFromBool(byte[] binaryData, int address)
        {
            return BitConverter.ToBoolean(binaryData, address);
        }

        public struct RegCostos
        {
            public string CodeCompra;
            public bool VImpuesto1;
            public bool VImpuesto2;
            public decimal CostoAnteriorBs;
            public decimal CostoAnteriorEx;
            public decimal CostoActualBs;
            public decimal CostoActualEx;
            public decimal CostoPromedioBs;
            public decimal CostoPromedioEx;
            public decimal MImpuesto1;
            public decimal MImpuesto2;
            public bool PorcentImp1;
            public bool Excento1;
            public bool PorcentImp2;
            public bool Excento2;
            public DateTime FechaVencimiento;
            public string NumeroDeLote;
            public RegUnPrecio[] Precios;
            public RegCostos(byte[] RegCosto)
            {
                this = default(RegCostos);
                this.CodeCompra = DelphiBinUtils.getStrFromStr(RegCosto, 1, 50);
                this.VImpuesto1 = DelphiBinUtils.getBoolFromBool(RegCosto, 51);
                this.VImpuesto2 = DelphiBinUtils.getBoolFromBool(RegCosto, 52);
                this.CostoAnteriorBs = DelphiBinUtils.getDecFromCurr(RegCosto, 56);
                this.CostoAnteriorEx = DelphiBinUtils.getDecFromCurr(RegCosto, 64);
                this.CostoActualBs = DelphiBinUtils.getDecFromCurr(RegCosto, 72);
                this.CostoActualEx = DelphiBinUtils.getDecFromCurr(RegCosto, 80);
                this.CostoPromedioBs = DelphiBinUtils.getDecFromCurr(RegCosto, 88);
                this.CostoPromedioEx = DelphiBinUtils.getDecFromCurr(RegCosto, 96);
                this.MImpuesto1 = DelphiBinUtils.getDecFromCurr(RegCosto, 104);
                this.MImpuesto2 = DelphiBinUtils.getDecFromCurr(RegCosto, 112);
                this.PorcentImp1 = DelphiBinUtils.getBoolFromBool(RegCosto, 120);
                this.Excento1 = DelphiBinUtils.getBoolFromBool(RegCosto, 121);
                this.PorcentImp2 = DelphiBinUtils.getBoolFromBool(RegCosto, 122);
                this.Excento2 = DelphiBinUtils.getBoolFromBool(RegCosto, 123);
                this.NumeroDeLote = DelphiBinUtils.getStrFromStr(RegCosto, 137, 50);
                this.Precios = new RegUnPrecio[6];
                int[] RegUnPrecioAddr = new int[]
                {
                192,
                264,
                336,
                408,
                480,
                552
                };
                int i = 0;
                checked
                {
                    int arg_169_0;
                    int num;
                    do
                    {
                        byte[] RegUnPrecioData = new byte[73];
                        Array.Copy(RegCosto, RegUnPrecioAddr[i], RegUnPrecioData, 0, 72);
                        this.Precios[i] = new RegUnPrecio(RegUnPrecioData);
                        i++;
                        arg_169_0 = i;
                        num = 5;
                    }
                    while (arg_169_0 <= num);
                }
            }
        }
        public struct RegUnPrecio
        {
            public bool PorcUtil;
            public bool PorcUtilEx;
            public decimal Utilidad;
            public decimal UtilidadEx;
            public decimal SinImpuesto;
            public decimal MtoImpuesto1;
            public decimal MtoImpuesto2;
            public decimal TotalPrecio;
            public decimal TotalPrecioEx;
            public byte TipoRound;
            public RegUnPrecio(byte[] RegUnPrecio)
            {
                this = default(RegUnPrecio);
                this.PorcUtil = DelphiBinUtils.getBoolFromBool(RegUnPrecio, 0);
                this.PorcUtilEx = DelphiBinUtils.getBoolFromBool(RegUnPrecio, 1);
                this.Utilidad = DelphiBinUtils.getDecFromCurr(RegUnPrecio, 8);
                this.UtilidadEx = DelphiBinUtils.getDecFromCurr(RegUnPrecio, 16);
                this.SinImpuesto = DelphiBinUtils.getDecFromCurr(RegUnPrecio, 24);
                this.MtoImpuesto1 = DelphiBinUtils.getDecFromCurr(RegUnPrecio, 32);
                this.MtoImpuesto2 = DelphiBinUtils.getDecFromCurr(RegUnPrecio, 40);
                this.TotalPrecio = DelphiBinUtils.getDecFromCurr(RegUnPrecio, 48);
                this.TotalPrecioEx = DelphiBinUtils.getDecFromCurr(RegUnPrecio, 56);
                this.TipoRound = RegUnPrecio[64];
            }
        }
    }
}
