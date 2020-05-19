using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;
using Microsoft.AnalysisServices.AdomdClient;

namespace ApiExamenP2.Controllers

{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    [RoutePrefix("v1/Analysis/Northwind")]
public class ApiNorthwindController : ApiController
{
        [HttpGet]
        [Route("Top5DimensionsPie/{dim}/{año}/{mes}/{order}")]
        public HttpResponseMessage Top5DimensionsPie(string dim, string año = "", string mes = "", string order = "DESC")
        {
            string dimension = string.Empty;
            string DimensionMes = "{(";
            string DimensionAño = "";


            dimension = dim + ".CHILDREN";
            DimensionMes = "{([Dim Tiempo].[Mes Espaniol].[" + mes + "],";
            DimensionAño = "[Dim Tiempo].[Anio].[" + año + "],";
            //Consulta MDX concatenando El mes + el año y la dimension que se esta seleccionando.
            string WITH = @"
                WITH
                SET [TopVentas] AS
                NONEMPTY(
                    ORDER(
                       " + DimensionMes + @" " + DimensionAño + @"STRTOSET(@Dimension))},
                            [Measures].[Hec Ventas Ventas], " + order + @"))";


            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Hec Ventas Ventas]
                }
                ON COLUMNS,
            ";

            string ROWS = @"
                NON EMPTY
                {
                    HEAD([TopVentas], 5)
                }
                ON ROWS
            ";

            string CUBO_NAME = "[DWH Northwind]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;
            Debug.Write(MDX_QUERY);
            List<string> clientes = new List<string>();
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();
            dynamic result = new
            {
                datosDimension = clientes,
                datosVenta = ventas,
                datosTabla = lstTabla
            };

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", dimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    { 
                        while (dr.Read())
                        {
                            int conteo = dr.FieldCount;
                            clientes.Add(dr.GetString(conteo - 2));
                            ventas.Add(decimal.Parse(dr.GetString(conteo - 1)));

                        }
                        dr.Close();
                    }
                }
            }
            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

    }
}