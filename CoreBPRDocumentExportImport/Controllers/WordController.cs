using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Dapper;
using Npgsql;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using CoreBPRDocumentExportImport.Models;

namespace CoreBPRDocumentExportImport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WordController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        public IConfiguration Configuration { get; }

        public WordController(ApplicationDbContext context, IConfiguration configuration)
        {
            _context = context;
            Configuration = configuration;
        }

        [HttpGet("export/{template}/{id}")]
        // public async Task<ActionResult<string>> GetExport(string template, string id)
        public async Task<ActionResult<string>> GetExport(string template, string id)
        {
            try
            {
                string filePath = await ExportWord(template, id);
                return Ok(filePath);
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }
        }

        private async Task<string> ExportWord(string template, string id)
        {
            string templateFolder = Configuration.GetSection("Folder").GetValue<string>("TemplateFolder");
            string uploadFolder = Configuration.GetSection("Folder").GetValue<string>("UploadFolder");
            string templateFileName = _context.DcxTemplateMasters.Where(x => x.TemplateId == template).Select(x => x.TemplateFilename).FirstOrDefault();
            string templateFile = Path.Combine(templateFolder, templateFileName);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(templateFile);
            string fileExtension = Path.GetExtension(templateFile);
            string createdFileName = fileNameWithoutExtension + "-" + id + fileExtension;
            string createdFile = Path.Combine(uploadFolder, createdFileName);

            try
            {
                

                if (System.IO.File.Exists(createdFile))
                {
                    int i = 1;
                    string newFileName = fileNameWithoutExtension + "-" + id + "(" + i.ToString() + ")" + fileExtension;
                    string newFilePath = Path.Combine(uploadFolder, newFileName);
                    while (System.IO.File.Exists(newFilePath))
                    {
                        i++;
                        newFileName = fileNameWithoutExtension + "-" + id + "(" + i.ToString() + ")" + fileExtension;
                        newFilePath = Path.Combine(uploadFolder, newFileName);
                    }
                    createdFileName = newFileName;
                    createdFile = newFilePath;
                }

                List<DcxTemplateMaster> listDcxTemplateMaster;
                List<DcxTemplateDetail> listDcxTemplateDetail;
                string sqlSelect, sqlFrom, sqlWhere, sqlQuery;
                DataTable tableResult;

                // Open Template File then Save as Created File
                using (var templateFileStream = new FileStream(templateFile, FileMode.Open, FileAccess.Read))
                {
                    //Loads an existing Word document into DocIO instance
                    WordDocument document = new WordDocument(templateFileStream, FormatType.Docx);

                    using (var createdFileStream = new FileStream(createdFile, FileMode.Create, FileAccess.Write))
                    {
                        document.Save(createdFileStream, FormatType.Docx);
                        createdFileStream.Close();
                    }

                    document.Close();
                    templateFileStream.Close();
                }

                // Process Created File
                using (var createdFileStream = new FileStream(createdFile, FileMode.Open, FileAccess.ReadWrite))
                {
                    //Loads an existing Word document into DocIO instance
                    WordDocument document = new WordDocument(createdFileStream, FormatType.Docx);



                    listDcxTemplateMaster = _context.DcxTemplateMasters.Where(x => x.TemplateId == template).OrderBy(x => x.Id).ToList();

                    foreach (var dcxTemplateMaster in listDcxTemplateMaster)
                    {
                        int sequence = dcxTemplateMaster.SequenceId;
                        string masterQuery = dcxTemplateMaster.SqlCommand;

                        // if Master-Level-Query is not empty, process Query from Master-Level
                        if ((masterQuery != null) && (masterQuery.Trim() != String.Empty))
                        {

                        }
                        // if Template-Level-Query is empty, process Query from Detail-Level
                        else
                        {
                            listDcxTemplateDetail = _context.DcxTemplateDetails.Where(x => x.TemplateId == template && x.SequenceId == sequence).OrderBy(x => x.Id).ToList();

                            sqlSelect = "";
                            sqlFrom = "";
                            sqlWhere = "";
                            sqlQuery = "";

                            foreach (var dcxTemplateDetail in listDcxTemplateDetail)
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

                            sqlWhere = sqlWhere.Replace("{*id*}", id);
                            sqlQuery = sqlSelect + sqlFrom + sqlWhere;

                            // Dapper Connection
                            using (var dconn = new NpgsqlConnection(Configuration.GetConnectionString("DefaultConnection")))
                            {
                                var queryResult = dconn.Query(sqlQuery).ToList();

                                var jsonResult = JsonConvert.SerializeObject(queryResult);

                                tableResult = (DataTable)JsonConvert.DeserializeObject(jsonResult, typeof(DataTable));
                            }

                            foreach (var dcxTemplateDetail in listDcxTemplateDetail)
                            {
                                string bookmarkName = dcxTemplateDetail.CellColumn;

                                string fieldContent = tableResult.Rows[0][dcxTemplateDetail.CellColumn].ToString();

                                //Gets the bookmark instance by using FindByName method of BookmarkCollection with bookmark name
                                Syncfusion.DocIO.DLS.Bookmark bookmark = document.Bookmarks.FindByName(bookmarkName);

                                //Creates the bookmark navigator instance to access the bookmark
                                BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);

                                //Moves the virtual cursor to the location before the end of the bookmark "Northwind"
                                bookmarkNavigator.MoveToBookmark(bookmarkName);

                                //Inserts a new text before the bookmark end of the bookmark
                                bookmarkNavigator.InsertText(fieldContent);
                            }
                        }

                    }



                    document.Save(createdFileStream, FormatType.Docx);
                    document.Close();
                    createdFileStream.Close();
                }

                return createdFileName;
            }
            catch (Exception e)
            {
                if (e.Message.Contains(templateFile))
                {
                    throw new Exception(e.Message.Replace(templateFile, templateFileName));
                }
                if (e.Message.Contains(createdFile))
                {
                    throw new Exception(e.Message.Replace(createdFile, createdFileName));
                }
                throw new Exception(e.Message);
            }
        }
    }
}
