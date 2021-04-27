using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace CoreBPRDocumentExportImport.Models
{
    [Table("dcx_templatedetail")]
    public class DcxTemplateDetail
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        [Column("template_id")]
        public string TemplateId { get; set; }

        [Column("sheet_id")]
        public string SheetId { get; set; }

        [Column("sequence_id")]
        public int SequenceId { get; set; }

        [Column("cell_column")]
        public string CellColumn { get; set; }

        [Column("cell_row")]
        public string CellRow { get; set; }

        [Column("sql_select")]
        public string SqlSelect { get; set; }

        [Column("sql_from")]
        public string SqlFrom { get; set; }

        [Column("sql_where")]
        public string SqlWhere { get; set; }
    }
}
