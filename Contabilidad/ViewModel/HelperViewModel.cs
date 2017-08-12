using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Contabilidad.ViewModel
{
    public class HelperViewModel
    {
        public string tipoDeComprobante { get; set; }
        public string serie { get; set; }
        public string folio { get; set; }
        public string fecha { get; set; }
        public string hora { get; set; }
        
        //Emisor
        public string rfcEmisor { get; set; }
        public string nombreEmisor { get; set; }
        public string RegimenFiscal { get; set; }
        public string calleEmisor { get; set; }
        public string coloniaEmisor { get; set; }
        public string localidadEmisor { get; set; }
        public string municipioEmisor { get; set; }
        public string paisEmisor { get; set; }
        public string estadoEmisor { get; set; }
        public string cpEmisor { get; set; }

        //Receptor
        public string rfcReceptor { get; set; }
        public string nombreReceptor { get; set; }
        public string calleReceptor { get; set; }
        public string coloniaReceptor { get; set; }
        public string localidadReceptor { get; set; }
        public string municipioReceptor { get; set; }
        public string paisReceptor { get; set; }
        public string estadoReceptor { get; set; }
        public string cpReceptor { get; set; }

        //Conceptos
        public List<string> cantidad { get; set; }
        public List<string> unidad { get; set; }
        public List<string> descripcion { get; set; }
        public List<string> valorUnitario { get; set; }
        public List<string> importe { get; set; }

        public string serieEmisor { get; set; }
        public string folioFiscal { get; set; }
        public string noCertificadoSAT { get; set; }
        public string fechaHoraCertificacion { get; set; }
        public string subtotal { get; set; }
        public string iva { get; set; }
        public string descuentos { get; set; }
        public string total { get; set; }

    }
}