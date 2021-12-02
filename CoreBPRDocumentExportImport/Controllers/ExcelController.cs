using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using System.IO;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Syncfusion.XlsIO;
using CoreBPRDocumentExportImport.Models;
using Npgsql;
using Newtonsoft.Json;
using Dapper;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace CoreBPRDocumentExportImport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private IConfiguration Configuration { get; }

        public ExcelController(ApplicationDbContext context, IConfiguration configuration)
        {
            _context = context;
            Configuration = configuration;
        }



        // GET: api/<ExcelController>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<ExcelController>/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        // POST api/<ExcelController>
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/<ExcelController>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/<ExcelController>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }



        [HttpGet("lapkeu/{template}")]
        public async Task<ActionResult<string>> GetExport(string template)
        {
            try
            {
                string filePath = await GenerateLaporanKeuangan(template);
                return Ok(filePath);
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }
        }

        // Generate Laporan Keuangan
        private async Task<string> GenerateLaporanKeuangan(string template)
        {
            string templateFolder = Configuration["Folder:TemplateFolder"];
            string uploadFolder = Configuration["Folder:UploadFolder"];
            string templateFileName = _context.DcxTemplateMasters.Where(x => x.TemplateId == template).Select(x => x.TemplateFilename).FirstOrDefault();
            string templateFile = Path.Combine(templateFolder, templateFileName);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(templateFile);
            string fileExtension = Path.GetExtension(templateFile);
            DateTime today = DateTime.Now;
            string strToday = today.ToString("yyyy-MM-dd");
            string createdFileName = fileNameWithoutExtension + "-" + strToday + fileExtension;
            string createdFile = Path.Combine(uploadFolder, createdFileName);

            try
            {
                if (System.IO.File.Exists(createdFile))
                {
                    int i = 1;
                    string newFileName = fileNameWithoutExtension + "-" + strToday + "(" + i.ToString() + ")" + fileExtension;
                    string newFilePath = Path.Combine(uploadFolder, newFileName);
                    while (System.IO.File.Exists(newFilePath))
                    {
                        i++;
                        newFileName = fileNameWithoutExtension + "-" + strToday + "(" + i.ToString() + ")" + fileExtension;
                        newFilePath = Path.Combine(uploadFolder, newFileName);
                    }
                    createdFileName = newFileName;
                    createdFile = newFilePath;
                }

                List<DcxTemplateMaster> listDcxTemplateMaster;
                List<DcxTemplateDetail> listDcxTemplateDetail;
                string sheetId, sqlSelect, sqlFrom, sqlWhere, sqlQuery, pos, saldo;
                int sequenceId;
                DataTable tableResult;

                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;

                    FileStream templateFileStream = new FileStream(templateFile, FileMode.OpenOrCreate, FileAccess.Read);

                    //Loads or open an existing workbook through Open method of IWorkbooks
                    IWorkbook workbook = excelEngine.Excel.Workbooks.Open(templateFileStream);

                    // Loop for DcxTemplateMaster
                    listDcxTemplateMaster = _context.DcxTemplateMasters.Where(x => x.TemplateId == template).OrderBy(x => x.Id).ToList();

                    foreach (DcxTemplateMaster dcxTemplateMaster in listDcxTemplateMaster)
                    {
                        sheetId = dcxTemplateMaster.SheetId;
                        sequenceId = dcxTemplateMaster.SequenceId;

                        //Access a worksheet from workbook
                        IWorksheet worksheet = workbook.Worksheets[sheetId];

                        // Loop for DcxTemplateDetail
                        listDcxTemplateDetail = _context.DcxTemplateDetails.Where(x => x.TemplateId == template && x.SheetId == sheetId && x.SequenceId == sequenceId).OrderBy(x => x.Id).ToList();

                        sqlSelect = "";
                        sqlFrom = "";
                        sqlWhere = "";
                        sqlQuery = "";

                        foreach (DcxTemplateDetail dcxTemplateDetail in listDcxTemplateDetail)
                        {
                            if ((dcxTemplateDetail.SqlSelect != null) && (dcxTemplateDetail.SqlSelect.Trim() != String.Empty))
                            {
                                sqlSelect = sqlSelect + dcxTemplateDetail.SqlSelect;
                            }

                            if ((dcxTemplateDetail.SqlFrom != null) && (dcxTemplateDetail.SqlFrom.Trim() != String.Empty))
                            {
                                sqlFrom = sqlFrom + dcxTemplateDetail.SqlFrom;
                            }

                            if ((dcxTemplateDetail.SqlWhere != null) && (dcxTemplateDetail.SqlWhere.Trim() != String.Empty))
                            {
                                sqlWhere = sqlWhere + dcxTemplateDetail.SqlWhere;
                            }
                        }

                        sqlQuery = sqlSelect + sqlFrom + sqlWhere;

                        // Dapper Connection
                        using (var dconn = new NpgsqlConnection(Configuration.GetConnectionString("DefaultConnection")))
                        {
                            var queryResult = dconn.Query(sqlQuery).ToList();

                            var jsonResult = JsonConvert.SerializeObject(queryResult);

                            tableResult = (DataTable)JsonConvert.DeserializeObject(jsonResult, typeof(DataTable));

                            for (int i = 0; i < tableResult.Rows.Count; i++)
                            {
                                pos = tableResult.Rows[i][0].ToString();
                                saldo = tableResult.Rows[i][1].ToString();

                                worksheet.Range[pos].Text = saldo;
                            }
                        }
                    }

                    //Saving the workbook to disk in XLSX format
                    FileStream result = new FileStream(createdFile, FileMode.Create, FileAccess.Write);
                    workbook.SaveAs(result);

                    // Clean Process
                    workbook.Close();
                    templateFileStream.Close();
                    result.Close();

                    return createdFile;
                }
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }
    }
}
