using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace CoreBPRDocumentExportImport.Models
{
    [Table("dcx_templatemaster")]
    public class DcxTemplateMaster
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        [Column("template_group")]
        public string TemplateGroup { get; set; }

        [Column("template_id")]
        public string TemplateId { get; set; }

        [Column("sheet_id")]
        public string SheetId { get; set; }

        [Column("sequence_id")]
        public int SequenceId { get; set; }

        [Column("template_desc")]
        public string TemplateDesc { get; set; }

        [Column("sheet_desc")]
        public string SheetDesc { get; set; }

        [Column("sequence_desc")]
        public string SequenceDesc { get; set; }

        [Column("document_type")]
        public string DocumentType { get; set; }

        [Column("action_type")]
        public string ActionType { get; set; }

        [Column("template_filename")]
        public string TemplateFilename { get; set; }

        [Column("sql_command")]
        public string SqlCommand { get; set; }
    }
}
